import csv
import requests
import urllib3
import os
import logging
import re
import pandas as pd 
import threading
import itertools
from concurrent.futures import ThreadPoolExecutor, as_completed

# Disable SSL warnings for internal domains
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# --- CONFIGURATION ---
ALM_BASE_URL = "https://rb-alm-06-p.de.bosch.com/ccm"
USERNAME = "lop2cob"
PASSWORD = "shreyansh4991Ab#"

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_CSV_FILE = os.path.join(SCRIPT_DIR, "Release_ID_Ford.csv")
OUTPUT_EXCEL_FILE = os.path.join(SCRIPT_DIR, "Checking_Ford_30_threads.xlsx") 
LOG_FILE = os.path.join(SCRIPT_DIR, "extraction_log_checking_threads_ford.txt")
MAPPING_CSV_FILE = os.path.join(SCRIPT_DIR, "mapping.csv")
ADDED_LOG_FILE = os.path.join(SCRIPT_DIR, "added_time_log_checking_30_threads_ford.txt")
TEAM_ROSTER_FILE = os.path.join(SCRIPT_DIR, "Team_roster_EHN.xlsx")

# --- SETUP LOGGING ---
logger = logging.getLogger("alm_extractor")
logger.setLevel(logging.DEBUG) 
formatter = logging.Formatter('%(message)s')

master_file_handler = logging.FileHandler(LOG_FILE, mode='a', encoding='utf-8')
master_file_handler.setLevel(logging.DEBUG)
master_file_handler.setFormatter(formatter)
logger.addHandler(master_file_handler)

added_file_handler = logging.FileHandler(ADDED_LOG_FILE, mode='a', encoding='utf-8')
added_file_handler.setLevel(logging.INFO)
added_file_handler.setFormatter(formatter)
logger.addHandler(added_file_handler)

console_handler = logging.StreamHandler()
console_handler.setLevel(logging.INFO)
console_handler.setFormatter(formatter)
logger.addHandler(console_handler)

logging.getLogger("urllib3").setLevel(logging.WARNING)

# --- THREAD-LOCAL SESSION SETUP ---
thread_local = threading.local()

def get_session():
    if not hasattr(thread_local, "session"):
        s = requests.Session()
        s.auth = (USERNAME, PASSWORD)
        s.headers.update({"Accept": "application/json"})
        thread_local.session = s
    return thread_local.session

# --- BUFFERED LOGGER FOR THREADS ---
print_lock = threading.Lock()

class BufferedLogger:
    """Holds log messages in memory until a thread finishes, then prints them together."""
    def __init__(self):
        self.logs = []
        
    def info(self, msg):
        self.logs.append((logging.INFO, msg))
        
    def debug(self, msg):
        self.logs.append((logging.DEBUG, msg))
        
    def warning(self, msg):
        self.logs.append((logging.WARNING, msg))
        
    def error(self, msg):
        self.logs.append((logging.ERROR, msg))
        
    def flush(self):
        with print_lock:
            for level, msg in self.logs:
                logger.log(level, msg)

# --- CACHES ---
KNOWN_DEPARTMENT_URLS = {}
KNOWN_RESOLUTIONS = {}
KNOWN_USERS = {}
CATEGORY_MAPPING = {}
TEAM_ROSTER = set()

def load_team_roster():
    """Scans the Excel file and explicitly generates all permutations of names for bulletproof matching."""
    if os.path.exists(TEAM_ROSTER_FILE):
        try:
            xls = pd.ExcelFile(TEAM_ROSTER_FILE)
            sheets_used = []
            for sheet in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet)
                
                name_col = None
                for col in df.columns:
                    if str(col).strip().lower() == "names":
                        name_col = col
                        break
                
                if name_col is None:
                    continue
                    
                sheets_used.append(sheet)
                for val in df[name_col].dropna().astype(str):
                    clean_val = val.strip().lower()
                    if clean_val and len(clean_val) > 2 and clean_val != 'names':
                        # --- THE NEW PERMUTATION LOGIC ---
                        words = clean_val.split()
                        # Safe cap at 6 words to prevent crazy memory usage (720 combos max per person)
                        if len(words) <= 6:
                            for perm in itertools.permutations(words):
                                TEAM_ROSTER.add(" ".join(perm))
                        else:
                            TEAM_ROSTER.add(clean_val)
                            
            if not sheets_used:
                logger.warning(f"⚠️ No sheet in '{TEAM_ROSTER_FILE}' has a 'Names' column! Rescue feature will be skipped.")
            else:
                logger.info(f"✔️ Loaded {len(TEAM_ROSTER)} unique name combinations from sheets {sheets_used} in '{TEAM_ROSTER_FILE}' for Rescue Operations.")
        except Exception as e:
            logger.error(f"❌ Failed to load team roster from '{TEAM_ROSTER_FILE}': {e}")
    else:
        logger.warning(f"⚠️ Roster file '{TEAM_ROSTER_FILE}' not found! The 'Miscategorized' rescue feature will be skipped.")

def is_owner_in_roster(owner_string):
    """Checks if the extracted owner name exists anywhere in the team roster."""
    if not TEAM_ROSTER or not owner_string or owner_string == "Unassigned":
        return False
        
    owner_lower = owner_string.lower()
    name_part = owner_lower.split('(')[0].strip()
    if name_part in TEAM_ROSTER:
        return True
        
    for r_name in TEAM_ROSTER:
        if r_name and r_name in owner_lower:
            return True
            
    return False

def load_category_mapping():
    if os.path.exists(MAPPING_CSV_FILE):
        with open(MAPPING_CSV_FILE, mode="r", encoding="utf-8-sig") as mapfile:
            reader = csv.reader(mapfile)
            for row in reader:
                if len(row) >= 2:
                    cat_name = row[0].strip().lower()
                    bucket = row[1].strip().upper()
                    CATEGORY_MAPPING[cat_name] = bucket
        logger.info(f"✔️ Loaded {len(CATEGORY_MAPPING)} category mappings from '{MAPPING_CSV_FILE}'.")
    else:
        logger.warning(f"⚠️ Mapping file '{MAPPING_CSV_FILE}' not found! Will rely solely on keyword fallback.")

def get_workitem_json(url, blog):
    try:
        response = get_session().get(url, verify=False)
        if response.status_code == 200:
            return response.json()
    except Exception as e:
        blog.error(f"Error fetching {url}: {e}")
    return None

def format_date(iso_date_str):
    if not iso_date_str:
        return "Not Found"
    try:
        date_part = str(iso_date_str).split('T')[0]
        year, month, day = date_part.split('-')
        return f"{day}-{month}-{year}"
    except Exception:
        return str(iso_date_str)

def get_resolution_name(resolution_url, blog):
    if not resolution_url:
        return ""
    if resolution_url in KNOWN_RESOLUTIONS:
        return KNOWN_RESOLUTIONS[resolution_url]
        
    res_data = get_workitem_json(resolution_url, blog)
    title = ""
    if res_data:
        title = res_data.get("dcterms:title") or res_data.get("dc:title") or res_data.get("title") or ""
        
    title_lower = title.lower()
    
    if not title_lower:
        identifier = ""
        if res_data:
            identifier = res_data.get("dcterms:identifier") or res_data.get("dc:identifier") or ""
            
        if "invalid" in identifier.lower() or "resolution.r2" in resolution_url.lower():
            title_lower = "invalid"
        elif "resolution.r1" in resolution_url.lower():
            title_lower = "solved"
        else:
            title_lower = "unknown"

    KNOWN_RESOLUTIONS[resolution_url] = title_lower
    return title_lower

def get_owner_details(user_url):
    if not user_url or "unassigned" in user_url.lower(): 
        return "Unassigned"
        
    if user_url in KNOWN_USERS: return KNOWN_USERS[user_url]
    
    fallback_id = str(user_url).rstrip('/').split('/')[-1]
    
    try:
        api_headers = {"Accept": "application/rdf+xml", "OSLC-Core-Version": "2.0"}
        response = get_session().get(user_url, headers=api_headers, verify=False)
        
        if response.status_code == 200:
            xml_match = re.search(r'<foaf:name[^>]*>([^<]+)</foaf:name>', response.text, re.IGNORECASE)
            if xml_match:
                KNOWN_USERS[user_url] = xml_match.group(1).strip()
                return KNOWN_USERS[user_url]
    except Exception:
        pass
        
    KNOWN_USERS[user_url] = fallback_id
    return fallback_id

def get_country_and_rate(owner_string):
    if owner_string == "Unassigned": country = "Others"
    elif "-AU" in owner_string: country = "Australia"
    elif "-JP" in owner_string: country = "Japan"
    elif "BGSW/ECC" in owner_string: country = "Mexico"
    elif "ETAS-ECM/XPC-Abt1" in owner_string: country = "Germany"
    elif "Technology and Strategy,VM/ESB3-CB" in owner_string: country = "Germany"
    elif "T&S (Technology and Strategy" in owner_string: country = "Germany"
    elif "TS, VM/EAE-SD" in owner_string: country = "Germany"
    elif "Technology and Strategy (TS" in owner_string: country = "Germany"
    elif "VM/EAE1-CB" in owner_string: country = "Germany"
    elif "VM/ESB" in owner_string: country = "Germany"
    elif "VM/ESE1-Brg" in owner_string: country = "Portugal"
    elif "MS/EHV" in owner_string: country = "Vietnam"
    elif "VM/EFO" in owner_string: country = "North America"
    elif "-NA" in owner_string: country = "North America"
    elif "Adecco, MS/EAS5-VM" in owner_string: country = "Vietnam"
    elif "MS/EJH53-VM" in owner_string: country = "Vietnam"
    elif "Adecco, MS/ETA-VOS-FBL" in owner_string: country = "Vietnam"
    elif "MS/" in owner_string: country = "India"
    else: country = "Others"
        
    rate_card = {
        "India": 26.71,
        "Mexico": 38.85,
        "Vietnam": 22.73,
        "China": 42.74,
        "Japan": 42.74,
        "Hungary": 42.74,
        "Romania": 42.74,
        "Portugal": 42.74,
        "Austria": 42.74,
        "France": 93.48,
        "Germany": 93.48,
        "USA": 93.48,
        "North America": 93.48, 
        "Australia": 93.48,
        "Others": 0
    }
    return country, rate_card.get(country, 0)

def get_department_name(department_url, blog):
    if not department_url:
        return ""
    if department_url in KNOWN_DEPARTMENT_URLS:
        return KNOWN_DEPARTMENT_URLS[department_url]
    
    cat_data = get_workitem_json(department_url, blog)
    title = ""
    if cat_data:
        title = cat_data.get("dc:title") or cat_data.get("rtc_cm:hierarchicalName") or ""
    
    title_lower = title.lower()
    KNOWN_DEPARTMENT_URLS[department_url] = title_lower 
    return title_lower

def get_bucket_for_category(category_name):
    name = category_name.lower()
    bucket = CATEGORY_MAPPING.get(name)
    
    if bucket in ["NET", "DCOM", "DSM"]:
        return bucket
    elif bucket == "NA":
        return None 
        
    if "test" in name:
        return None

    if any(kw in name for kw in ["net", "fota", "var", "dpe", "hmi"]):
        return "NET"
    elif any(kw in name for kw in ["diag", "dcom", "sar"]):
        return "DCOM"
    elif any(kw in name for kw in ["obd", "dem", "dsm", "dws"]):
        return "DSM" 
        
    return None 

def process_hierarchy(work_item_url, release_id, blog, visited=None, depth=0, is_rework_branch=False):
    if visited is None:
        visited = set()

    country_efforts = {}

    if work_item_url in visited:
        return country_efforts
    visited.add(work_item_url)

    data = get_workitem_json(work_item_url, blog)
    if not data:
        return country_efforts

    item_id = data.get("dcterms:identifier") or data.get("dc:identifier") or data.get("identifier")
    if not item_id:
        item_id = str(work_item_url).rstrip('/').split('/')[-1]

    item_type_str = ""
    dc_type = data.get("dcterms:type") or data.get("dc:type") or data.get("type") or data.get("rtc_cm:type")
    
    if isinstance(dc_type, dict):
        item_type_str = dc_type.get("rdf:resource", "").lower()
    elif isinstance(dc_type, list) and len(dc_type) > 0:
        first_type = dc_type[0]
        item_type_str = first_type.get("rdf:resource", "").lower() if isinstance(first_type, dict) else str(first_type).lower()
    elif isinstance(dc_type, str):
        item_type_str = dc_type.lower()

    type_name_short = item_type_str.split("/")[-1] if "/" in item_type_str else item_type_str
    type_name_short = type_name_short.split(".")[-1]
    
    if not type_name_short:
        type_name_short = "unknown_type"

    if depth > 0 and "release" in type_name_short.lower():
        indent = "  " * depth
        blog.info(f"[Rel {release_id}] {indent}→ [SKIPPED SUB-RELEASE] ID: {item_id} (Preventing double-counting)")
        return country_efforts 

    time_spent_raw = data.get("rtc_cm:timeSpent")
    time_spent_ms = int(time_spent_raw) if time_spent_raw else 0

    current_is_rework = is_rework_branch or ("defect" in item_type_str)

    if time_spent_ms > 0:
        hours_logged = time_spent_ms / 3600000
        indent = "  " * depth
        
        task_title = data.get("dc:title") or data.get("dcterms:title") or data.get("title") or "Unknown Title"
        task_created_raw = data.get("dc:created") or data.get("dcterms:created") or data.get("created")
        task_created_formatted = format_date(task_created_raw)
        
        task_resolved_raw = data.get("rtc_cm:resolved") or data.get("resolved")
        task_resolved_formatted = format_date(task_resolved_raw)
        if task_resolved_formatted == "Not Found":
            task_resolved_formatted = "No Date"
            
        res_info = data.get("rtc_cm:resolution")
        res_url = ""
        if isinstance(res_info, dict):
            res_url = res_info.get("rdf:resource", "")
        elif isinstance(res_info, str):
            res_url = res_info
            
        child_res_status = get_resolution_name(res_url, blog) if res_url else ""
        is_invalid_status = "invalid" in child_res_status.lower()
        
        skip_due_to_invalid = is_invalid_status and (type_name_short != "defect")
        
        is_valid_type = type_name_short in ["task", "review", "defect"]

        is_valid_date = False
        
        date_to_check = task_resolved_raw if (task_resolved_raw and isinstance(task_resolved_raw, str) and "T" in task_resolved_raw) else task_created_raw
        
        if date_to_check and isinstance(date_to_check, str) and len(date_to_check) >= 4:
            try:
                task_year = int(date_to_check[:4]) 
                if task_year >= 2025:
                    is_valid_date = True
            except ValueError:
                pass 

        filed_against_data = data.get("rtc_cm:filedAgainst", {})
        task_category_url = ""
        if isinstance(filed_against_data, dict):
            task_category_url = filed_against_data.get("rdf:resource", "")
        elif isinstance(filed_against_data, str):
            task_category_url = filed_against_data
            
        task_department_name = get_department_name(task_category_url, blog)
        bucket = get_bucket_for_category(task_department_name)
        
        if skip_due_to_invalid:
            blog.debug(f"[Rel {release_id}] {indent}→ [IGNORED INVALID STATUS] Skipped {hours_logged:.2f} hrs | ID: {item_id} | Type: {type_name_short} | Title: {task_title}")
        elif not is_valid_type:
            blog.debug(f"[Rel {release_id}] {indent}→ [IGNORED TYPE] Skipped {hours_logged:.2f} hrs | ID: {item_id} | Type: {type_name_short} | Title: {task_title}")
        elif not is_valid_date:
            blog.debug(f"[Rel {release_id}] {indent}→ [IGNORED OLD DATE] Skipped {hours_logged:.2f} hrs | ID: {item_id} | Type: {type_name_short} | Created: {task_created_formatted} | Resolved: {task_resolved_formatted} (Failed 2025 filter)")
        else:
            task_owner_info = data.get("rtc_cm:ownedBy") or data.get("ownedBy")
            task_owner_url = task_owner_info.get("rdf:resource", "") if isinstance(task_owner_info, dict) else (task_owner_info or "")
            task_owner_string = get_owner_details(task_owner_url)
            
            if not bucket and is_owner_in_roster(task_owner_string):
                bucket = "GENERAL"
                task_department_name = f"{task_department_name} (Rescued via Roster)"
                
            if bucket:
                task_country, task_rate = get_country_and_rate(task_owner_string)
                
                if task_country not in country_efforts:
                    country_efforts[task_country] = {
                        "NET_DEV": 0, "DCOM_DEV": 0, "DSM_DEV": 0,
                        "NET_Rework": 0, "DCOM_Rework": 0, "DSM_Rework": 0,
                        "Miscategorized": 0,
                        "owner_string": task_owner_string, 
                        "rate": task_rate
                    }
                    
                if bucket == "GENERAL":
                    country_efforts[task_country]["Miscategorized"] += time_spent_ms
                    blog.info(f"[Rel {release_id}] {indent}→ [GENERAL Miscategorized] Added {hours_logged:.2f} hrs | ID: {item_id} | Type: {type_name_short} | Dept: '{task_department_name}' | Owner: {task_owner_string} | Country: {task_country} | Title: {task_title} | Created: {task_created_formatted} | Resolved: {task_resolved_formatted}")
                else:
                    bucket_key = f"{bucket}_Rework" if current_is_rework else f"{bucket}_DEV"
                    country_efforts[task_country][bucket_key] += time_spent_ms
                    rework_str = "REWORK" if current_is_rework else "DEV "
                    blog.info(f"[Rel {release_id}] {indent}→ [{bucket} {rework_str}] Added {hours_logged:.2f} hrs | ID: {item_id} | Type: {type_name_short} | Dept: '{task_department_name}' | Owner: {task_owner_string} | Country: {task_country} | Title: {task_title} | Created: {task_created_formatted} | Resolved: {task_resolved_formatted}")
            else:
                blog.debug(f"[Rel {release_id}] {indent}→ [IGNORED CATEGORY] Skipped {hours_logged:.2f} hrs (Not in Roster) | ID: {item_id} | Type: {type_name_short} | Dept: '{task_department_name}' | Title: {task_title}")

    children_data = data.get("rtc_cm:com.ibm.team.workitem.linktype.parentworkitem.children")
    if not children_data:
        children_data = []
    elif isinstance(children_data, dict):
        children_data = [children_data]

    for child in children_data:
        if isinstance(child, dict) and "rdf:resource" in child:
            child_url = child["rdf:resource"]
            child_efforts = process_hierarchy(child_url, release_id, blog, visited, depth + 1, current_is_rework)
            
            for c_name, c_data in child_efforts.items():
                if c_name not in country_efforts:
                    country_efforts[c_name] = {
                        "NET_DEV": 0, "DCOM_DEV": 0, "DSM_DEV": 0,
                        "NET_Rework": 0, "DCOM_Rework": 0, "DSM_Rework": 0,
                        "Miscategorized": 0,
                        "owner_string": c_data["owner_string"],
                        "rate": c_data["rate"]
                    }
                for k in ["NET_DEV", "DCOM_DEV", "DSM_DEV", "NET_Rework", "DCOM_Rework", "DSM_Rework", "Miscategorized"]:
                    country_efforts[c_name][k] += c_data[k]

    return country_efforts

def process_single_release(row):
    """Worker function for threading. Buffers all logs until finished."""
    release_id = row.get("Id", "").strip()
    pm_id = row.get("PM Interface Element ID", "").strip()
    
    blog = BufferedLogger() 
    country_rows_to_return = []
    
    blog.info(f"[Rel {release_id}] =========================================")
    blog.info(f"[Rel {release_id}] Checking Release ID: {release_id} [PM ID: {pm_id}]...")
    root_url = f"{ALM_BASE_URL}/resource/itemName/com.ibm.team.workitem.WorkItem/{release_id}"
    
    response = get_session().get(root_url, verify=False)
    if response and response.status_code == 200:
        release_data = response.json()
        raw_text = response.text 
        
        dc_type = release_data.get("dcterms:type") or release_data.get("dc:type") or release_data.get("type") or release_data.get("rtc_cm:type")
        item_type_str = ""
        if isinstance(dc_type, dict):
            item_type_str = dc_type.get("rdf:resource", "").lower()
        elif isinstance(dc_type, list) and len(dc_type) > 0:
            first_type = dc_type[0]
            item_type_str = first_type.get("rdf:resource", "").lower() if isinstance(first_type, dict) else str(first_type).lower()
        elif isinstance(dc_type, str):
            item_type_str = dc_type.lower()
            
        root_type_short = item_type_str.split("/")[-1] if "/" in item_type_str else item_type_str
        root_type_short = root_type_short.split(".")[-1]
        
        if not root_type_short:
            root_type_short = "Unknown"
        else:
            root_type_short = root_type_short.capitalize()
            
        blog.info(f"[Rel {release_id}] 📌 Root Item Type: {root_type_short}")
        
        res_info = release_data.get("rtc_cm:resolution")
        res_url = ""
        if isinstance(res_info, dict):
            res_url = res_info.get("rdf:resource", "")
        elif isinstance(res_info, str):
            res_url = res_info
            
        res_status = "Unresolved (New / In Progress)"
        if res_url:
            res_status = get_resolution_name(res_url, blog)
            
        is_invalid_status = "invalid" in res_status.lower()
        skip_due_to_invalid = is_invalid_status and (root_type_short.lower() != "defect")
        
        if skip_due_to_invalid:
            blog.warning(f"[Rel {release_id}] ❌ SKIPPING: Resolution is '{res_status.title()}'")
            blog.info("") 
            blog.flush() 
            return []
        else:
            if res_url:
                if is_invalid_status:
                    blog.info(f"[Rel {release_id}] ✔️ Resolution is '{res_status.title()}' but Type is 'Defect' (Continuing).")
                else:
                    blog.info(f"[Rel {release_id}] ✔️ Resolution is '{res_status.title()}'.")
            else:
                blog.info(f"[Rel {release_id}] ✔️ Status is '{res_status}' (No resolution yet).")

        rel_owner_info = release_data.get("rtc_cm:ownedBy") or release_data.get("ownedBy")
        rel_owner_url = rel_owner_info.get("rdf:resource", "") if isinstance(rel_owner_info, dict) else (rel_owner_info or "")
        release_owner_string = get_owner_details(rel_owner_url)
        
        match = re.search(r'"(?:dc|dcterms):created"\s*:\s*"([^"]+)"', raw_text, re.IGNORECASE)
        created_raw = match.group(1) if match else (release_data.get("dc:created") or release_data.get("dcterms:created"))

        match = re.search(r'"(?:rtc_cm:)?resolved"\s*:\s*"([^"]+)"', raw_text, re.IGNORECASE)
        resolved_raw = match.group(1) if match else (release_data.get("rtc_cm:resolved") or release_data.get("resolved"))
        
        created_formatted = format_date(created_raw)
        resolved_formatted = format_date(resolved_raw)
        if resolved_formatted == "Not Found":
            resolved_formatted = "No Date"
        
        blog.info(f"[Rel {release_id}] 👤 Release Owned By: {release_owner_string} (Tasks may be owned by others)")
        blog.info(f"[Rel {release_id}] 📅 Created: {created_formatted} | Resolved: {resolved_formatted}")

    blog.info(f"[Rel {release_id}] Processing Hierarchy...")
    
    efforts_by_country = process_hierarchy(root_url, release_id, blog)
    
    if not efforts_by_country:
        fallback_country, fallback_rate = get_country_and_rate(release_owner_string)
        efforts_by_country = {
            fallback_country: {
                "NET_DEV": 0, "DCOM_DEV": 0, "DSM_DEV": 0,
                "NET_Rework": 0, "DCOM_Rework": 0, "DSM_Rework": 0,
                "Miscategorized": 0,
                "owner_string": release_owner_string,
                "rate": fallback_rate
            }
        }

    blog.info(f"\n[Rel {release_id}] --- Final Hours for Release {release_id} ---")
    
    for country, data in efforts_by_country.items():
        blog.info(f"[Rel {release_id}]   [🌍 {country}]")
        
        net_dev = data["NET_DEV"] / 3600000
        dcom_dev = data["DCOM_DEV"] / 3600000
        dsm_dev = data["DSM_DEV"] / 3600000
        
        net_rew = data["NET_Rework"] / 3600000
        dcom_rew = data["DCOM_Rework"] / 3600000
        dsm_rew = data["DSM_Rework"] / 3600000
        
        miscategorized = data.get("Miscategorized", 0) / 3600000
        
        if net_dev > 0: blog.info(f"[Rel {release_id}]     NET_DEV: {net_dev:.2f} hrs")
        if dcom_dev > 0: blog.info(f"[Rel {release_id}]     DCOM_DEV: {dcom_dev:.2f} hrs")
        if dsm_dev > 0: blog.info(f"[Rel {release_id}]     DSM_DEV: {dsm_dev:.2f} hrs")
        if net_rew > 0: blog.info(f"[Rel {release_id}]     NET_Rework: {net_rew:.2f} hrs")
        if dcom_rew > 0: blog.info(f"[Rel {release_id}]     DCOM_Rework: {dcom_rew:.2f} hrs")
        if dsm_rew > 0: blog.info(f"[Rel {release_id}]     DSM_Rework: {dsm_rew:.2f} hrs")
        if miscategorized > 0: blog.info(f"[Rel {release_id}]     Miscategorized: {miscategorized:.2f} hrs")
        
        if (net_dev + dcom_dev + dsm_dev + net_rew + dcom_rew + dsm_rew + miscategorized) == 0:
            blog.info(f"[Rel {release_id}]     No Valid Hours Logged (0.00 hrs)")

        net_final = net_dev + net_rew
        dcom_dsm_dev = dcom_dev + dsm_dev
        dcom_dsm_rew = dcom_rew + dsm_rew
        dcom_dsm_final = dcom_dsm_dev + dcom_dsm_rew

        country_row = row.copy() 
        
        country_row["Item Type"] = root_type_short
        country_row["Owner Details"] = data["owner_string"]
        country_row["Country"] = country
        country_row["Rate Card (€)"] = data["rate"]
        country_row["Creation Date"] = created_formatted
        country_row["Resolution Date"] = resolved_formatted

        country_row["NET_DEV (Hours)"] = round(net_dev, 2)
        country_row["DCOM_DEV (Hours)"] = round(dcom_dev, 2)
        country_row["DSM_DEV (Hours)"] = round(dsm_dev, 2)
        
        country_row["NET_FINAL (Hours)"] = round(net_final, 2)
        country_row["DCOM_DSM_DEV (Hours)"] = round(dcom_dsm_dev, 2)
        
        country_row["NET_Rework (Hours)"] = round(net_rew, 2)
        country_row["DCOM_Rework (Hours)"] = round(dcom_rew, 2)
        country_row["DSM_Rework (Hours)"] = round(dsm_rew, 2)
        
        country_row["DCOM_DSM_REWORK_TOTAL (Hours)"] = round(dcom_dsm_rew, 2)
        country_row["DCOM_DSM_FINAL (Hours)"] = round(dcom_dsm_final, 2)
        
        country_row["Miscategorized (Hours)"] = round(miscategorized, 2)
        
        country_rows_to_return.append(country_row)
        
    blog.info("") 
    blog.flush() 
    return country_rows_to_return

if __name__ == "__main__":
    logger.info("\n---> [NEW RUN STARTING: FULL PERMUTATION ROSTER MATCHING (30 THREADS)] <---")
    logger.info(f"Master Log: {LOG_FILE}")
    logger.info(f"Clean Log (Added Only): {ADDED_LOG_FILE}\n")
    
    load_category_mapping()
    load_team_roster() 
    
    processed_rows = []
    
    try:
        with open(INPUT_CSV_FILE, mode="r", encoding="utf-16") as infile:
            reader = csv.DictReader(infile, delimiter="\t")
            all_rows = list(reader)
            
            normal_rows = []
            delayed_rows = []
            
            logger.info(f"Reading and cleaning {len(all_rows)} rows from input file...")
            for row in all_rows:
                release_id = row.get("Id", "").strip()
                if not release_id:
                    continue
                    
                pm_id = row.get("PM Interface Element ID", "").strip()
                
                if pm_id.startswith("BM"):
                    pm_id = pm_id.split('_')[0]
                    row["PM Interface Element ID"] = pm_id 
                    
                if pm_id.startswith("|") or "Official_SW_Plan_Draft" in pm_id:
                    delayed_rows.append(row)
                else:
                    normal_rows.append(row)
            
            ordered_rows = normal_rows + delayed_rows
            logger.info(f"Sorting complete. Starting ThreadPoolExecutor with 30 threads...\n")
            
            with ThreadPoolExecutor(max_workers=30) as executor:
                future_to_row = {executor.submit(process_single_release, row): row for row in ordered_rows}
                
                for future in as_completed(future_to_row):
                    try:
                        country_rows = future.result()
                        if country_rows:
                            processed_rows.extend(country_rows)
                    except Exception as exc:
                        logger.error(f"Thread generated an exception: {exc}")
                
        if processed_rows:
            df = pd.DataFrame(processed_rows)
            
            cols = list(df.columns)
            if "Item Type" in cols and "Id" in cols:
                cols.remove("Item Type")
                id_index = cols.index("Id")
                cols.insert(id_index + 1, "Item Type")
                df = df[cols]
                
            df.to_excel(OUTPUT_EXCEL_FILE, index=False)
            logger.info(f"\n✅ Success! All threads finished. Processed data saved to Excel file: '{OUTPUT_EXCEL_FILE}'.")
        else:
            logger.warning("\n⚠️ No rows were processed.")

    except FileNotFoundError:
        logger.error(f"Error: Could not find '{INPUT_CSV_FILE}'.")
    except ImportError:
        logger.error("\n❌ ERROR: Pandas or openpyxl is not installed!")
        logger.error("Please run: pip install pandas openpyxl")
