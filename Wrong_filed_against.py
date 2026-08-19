import csv
import requests
import urllib3
import os
import logging
import re
import pandas as pd 
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed

# Disable SSL warnings for internal domains
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# --- CONFIGURATION ---
ALM_BASE_URL = "https://rb-alm-06-p.de.bosch.com/ccm"
USERNAME = "lop2cob"
PASSWORD = "shreyansh4991Ab#"

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_CSV_FILE = os.path.join(SCRIPT_DIR, "Release ID_MB.csv")
OUTPUT_EXCEL_FILE = os.path.join(SCRIPT_DIR, "open_Daimler_Data_Miscategorized.xlsx") 
LOG_FILE = os.path.join(SCRIPT_DIR, "extraction_log_threading_.txt")
MAPPING_CSV_FILE = os.path.join(SCRIPT_DIR, "mapping.csv")
ADDED_LOG_FILE = os.path.join(SCRIPT_DIR, "open_added_time_log_threading.txt")

# --- NEW CONFIG FOR ROSTER ---
ROSTER_EXCEL_FILE = os.path.join(SCRIPT_DIR, "team_roster.xlsx")

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
TARGET_NAMES = [] # Store the 134 names here

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
        logger.warning(f"⚠️ Mapping file '{MAPPING_CSV_FILE}' not found!")

def load_target_names():
    if os.path.exists(ROSTER_EXCEL_FILE):
        try:
            # Added sheet_name='EHE' to read the correct sheet. 
            # If your sheet has a different name, replace 'EHE' with the correct name.
            df = pd.read_excel(ROSTER_EXCEL_FILE, sheet_name='EHE')
            
            # Strip any hidden leading/trailing spaces from column names to ensure a match
            df.columns = df.columns.str.strip()
            
            if "Names" in df.columns:
                names = df["Names"].dropna().astype(str).tolist()
                for n in names:
                    TARGET_NAMES.append(n.strip().lower())
                logger.info(f"✔️ Loaded {len(TARGET_NAMES)} target names from '{ROSTER_EXCEL_FILE}'.")
            else:
                logger.error(f"⚠️ Column 'Names' not found in '{ROSTER_EXCEL_FILE}'. Available columns: {list(df.columns)}")
        except Exception as e:
            logger.error(f"⚠️ Error reading '{ROSTER_EXCEL_FILE}': {e}")
    else:
        logger.warning(f"⚠️ Roster file '{ROSTER_EXCEL_FILE}' not found! No General WIs will be matched.")


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
        "India": 26.71, "Mexico": 38.85, "Vietnam": 22.73, "China": 42.74,
        "Japan": 42.74, "Hungary": 42.74, "Romania": 42.74, "Portugal": 42.74,
        "Austria": 42.74, "France": 93.48, "Germany": 93.48, "USA": 93.48,
        "North America": 93.48, "Australia": 93.48, "Others": 0
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
        dept_lower = task_department_name.lower()

        # --- EXCLUSION FILTER LOGIC ---
        is_excluded = False
        
        # 1. Skip if mapped in mapping.csv
        if dept_lower in CATEGORY_MAPPING:
            is_excluded = True
        else:
            # 2. Skip if it contains standard keywords
            keywords = ["net", "fota", "var", "dpe", "hmi", "diag", "dcom", "sar", "obd", "dem", "dsm", "dws"]
            if any(kw in dept_lower for kw in keywords):
                is_excluded = True

        bucket = None
        task_owner_string = "Unknown"
        
        # --- INCLUSION FILTER LOGIC (ROSTER MATCH) ---
        # Only fetch owner details and check roster if it SURVIVES exclusions (e.g., 'test' or unmapped)
        if not is_excluded and is_valid_type and is_valid_date:
            task_owner_info = data.get("rtc_cm:ownedBy") or data.get("ownedBy")
            task_owner_url = ""
            if isinstance(task_owner_info, dict):
                task_owner_url = task_owner_info.get("rdf:resource", "")
            elif isinstance(task_owner_info, str):
                task_owner_url = task_owner_info
            
            task_owner_string = get_owner_details(task_owner_url)
            owner_lower = task_owner_string.lower()
            
            # Check if owner name is in our 134 names roster
            for target_name in TARGET_NAMES:
                if target_name in owner_lower:
                    bucket = "General"
                    break

        if not is_valid_type:
            blog.debug(f"[Rel {release_id}] {indent}→ [IGNORED TYPE] Skipped {hours_logged:.2f} hrs | ID: {item_id} | Type: {type_name_short} | Title: {task_title}")
        elif not is_valid_date:
            blog.debug(f"[Rel {release_id}] {indent}→ [IGNORED OLD DATE] Skipped {hours_logged:.2f} hrs | ID: {item_id} (Failed 2025 filter)")
        elif is_excluded:
            blog.debug(f"[Rel {release_id}] {indent}→ [IGNORED DEPT MAPPED] Skipped {hours_logged:.2f} hrs | ID: {item_id} | Dept: '{task_department_name}'")
        elif bucket == "General":
            # WE FOUND A MATCH!
            task_country, task_rate = get_country_and_rate(task_owner_string)
            
            if task_country not in country_efforts:
                country_efforts[task_country] = {
                    "General_DEV": 0, "General_Rework": 0,
                    "owner_string": task_owner_string, 
                    "rate": task_rate
                }
                
            bucket_key = "General_Rework" if current_is_rework else "General_DEV"
            country_efforts[task_country][bucket_key] += time_spent_ms
            rework_str = "REWORK" if current_is_rework else "DEV "
            
            blog.info(f"[Rel {release_id}] {indent}→ [GENERAL {rework_str}] Added {hours_logged:.2f} hrs | ID: {item_id} | Dept: '{task_department_name}' | Owner: {task_owner_string} | Title: {task_title}")
        else:
            blog.debug(f"[Rel {release_id}] {indent}→ [IGNORED OWNER NOT IN ROSTER] Skipped {hours_logged:.2f} hrs | ID: {item_id} | Owner: {task_owner_string}")

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
                        "General_DEV": 0, "General_Rework": 0,
                        "owner_string": c_data["owner_string"],
                        "rate": c_data["rate"]
                    }
                country_efforts[c_name]["General_DEV"] += c_data["General_DEV"]
                country_efforts[c_name]["General_Rework"] += c_data["General_Rework"]

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
        
        res_info = release_data.get("rtc_cm:resolution")
        res_url = ""
        if isinstance(res_info, dict):
            res_url = res_info.get("rdf:resource", "")
        elif isinstance(res_info, str):
            res_url = res_info
            
        res_status = "Unresolved (New / In Progress)"
        if res_url:
            res_status = get_resolution_name(res_url, blog)
            
        if "invalid" in res_status.lower():
            blog.warning(f"[Rel {release_id}] ❌ SKIPPING: Resolution is '{res_status.title()}'")
            blog.info("") 
            blog.flush() 
            return []
        else:
            if res_url:
                blog.info(f"[Rel {release_id}] ✔️ Resolution is '{res_status.title()}'.")
            else:
                blog.info(f"[Rel {release_id}] ✔️ Status is '{res_status}' (No resolution yet).")

        rel_owner_info = release_data.get("rtc_cm:ownedBy") or release_data.get("ownedBy")
        rel_owner_url = rel_owner_info.get("rdf:resource", "") if isinstance(rel_owner_info, dict) else (rel_owner_info or "")
        release_owner_string = get_owner_details(rel_owner_url)
        
        match = re.search(r'"(?:dc|dcterms):created"\s*:\s*"([^"]+)"', raw_text, re.IGNORECASE)
        created_raw = match.group(1) if match else (release_data.get("dc:created") or release_data.get("dcterms:created"))

        match = re.search(r'"(?:rtc_cm)?resolved"\s*:\s*"([^"]+)"', raw_text, re.IGNORECASE)
        resolved_raw = match.group(1) if match else (release_data.get("rtc_cm:resolved") or release_data.get("resolved"))
        
        created_formatted = format_date(created_raw)
        resolved_formatted = format_date(resolved_raw)
        if resolved_formatted == "Not Found":
            resolved_formatted = "No Date"
        
        blog.info(f"[Rel {release_id}] 👤 Release Owned By: {release_owner_string} (Tasks may be owned by others)")
        blog.info(f"[Rel {release_id}] 📅 Created: {created_formatted} | Resolved: {resolved_formatted}")

    blog.info(f"[Rel {release_id}] Processing Hierarchy for Miscategorized WIs...")
    
    efforts_by_country = process_hierarchy(root_url, release_id, blog)
    
    if not efforts_by_country:
        fallback_country, fallback_rate = get_country_and_rate(release_owner_string)
        efforts_by_country = {
            fallback_country: {
                "General_DEV": 0, "General_Rework": 0,
                "owner_string": release_owner_string,
                "rate": fallback_rate
            }
        }

    blog.info(f"\n[Rel {release_id}] --- Final MISCATEGORIZED Hours for Release {release_id} ---")
    
    for country, data in efforts_by_country.items():
        blog.info(f"[Rel {release_id}]   [🌍 {country}]")
        
        gen_dev = data["General_DEV"] / 3600000
        gen_rew = data["General_Rework"] / 3600000
        
        if gen_dev > 0: blog.info(f"[Rel {release_id}]     General_DEV: {gen_dev:.2f} hrs")
        if gen_rew > 0: blog.info(f"[Rel {release_id}]     General_Rework: {gen_rew:.2f} hrs")
        
        if (gen_dev + gen_rew) == 0:
            blog.info(f"[Rel {release_id}]     No Miscategorized Hours Found (0.00 hrs)")

        gen_final = gen_dev + gen_rew

        country_row = row.copy() 
        
        country_row["Owner Details"] = data["owner_string"]
        country_row["Country"] = country
        country_row["Rate Card (€)"] = data["rate"]
        country_row["Creation Date"] = created_formatted
        country_row["Resolution Date"] = resolved_formatted

        # Only adding General columns to keep the Excel output clean and relevant to this specific task
        country_row["General_DEV (Hours)"] = round(gen_dev, 2)
        country_row["General_Rework (Hours)"] = round(gen_rew, 2)
        country_row["General_FINAL (Hours)"] = round(gen_final, 2)
        
        country_rows_to_return.append(country_row)
        
    blog.info("") 
    blog.flush() 
    return country_rows_to_return

if __name__ == "__main__":
    logger.info("\n---> [NEW RUN STARTING: SEARCHING FOR MISCATEGORIZED WIs BASED ON ROSTER] <---")
    logger.info(f"Master Log: {LOG_FILE}")
    logger.info(f"Clean Log (Added Only): {ADDED_LOG_FILE}\n")
    
    load_category_mapping()
    load_target_names() # Load the 134 names
    
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
            logger.info(f"Sorting complete. Starting ThreadPoolExecutor with 15 threads...\n")
            
            with ThreadPoolExecutor(max_workers=15) as executor:
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
            df.to_excel(OUTPUT_EXCEL_FILE, index=False)
            logger.info(f"\n✅ Success! All threads finished. Processed data saved to Excel file: '{OUTPUT_EXCEL_FILE}'.")
        else:
            logger.warning("\n⚠️ No rows were processed.")

    except FileNotFoundError:
        logger.error(f"Error: Could not find '{INPUT_CSV_FILE}'.")
    except ImportError:
        logger.error("\n❌ ERROR: Pandas or openpyxl is not installed!")
        logger.error("Please run: pip install pandas openpyxl")
