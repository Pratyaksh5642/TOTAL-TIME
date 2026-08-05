import csv
import requests
import urllib3
import os
import logging
import re
import pandas as pd # <--- Added for Excel export

# Disable SSL warnings for internal domains
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# --- CONFIGURATION ---
ALM_BASE_URL = "https://rb-alm-06-p.de.bosch.com/ccm"
USERNAME = "lop2cob"
PASSWORD = "shreyansh4991Ab#"

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_CSV_FILE = os.path.join(SCRIPT_DIR, "input.csv")
OUTPUT_EXCEL_FILE = os.path.join(SCRIPT_DIR, "output_categorized.xlsx") # <--- Changed to .xlsx
LOG_FILE = os.path.join(SCRIPT_DIR, "extraction_log.txt")
MAPPING_CSV_FILE = os.path.join(SCRIPT_DIR, "mapping.csv")

# --- SETUP LOGGING ---
logger = logging.getLogger("alm_extractor")
logger.setLevel(logging.DEBUG) 
formatter = logging.Formatter('%(message)s')

file_handler = logging.FileHandler(LOG_FILE, mode='a', encoding='utf-8')
file_handler.setLevel(logging.DEBUG)
file_handler.setFormatter(formatter)
logger.addHandler(file_handler)

console_handler = logging.StreamHandler()
console_handler.setLevel(logging.INFO)
console_handler.setFormatter(formatter)
logger.addHandler(console_handler)

logging.getLogger("urllib3").setLevel(logging.WARNING)

# --- SETUP SESSION ---
session = requests.Session()
session.auth = (USERNAME, PASSWORD)
headers = {
    "Accept": "application/json",
    "OSLC-Core-Version": "2.0"
}

# --- CACHES ---
KNOWN_DEPARTMENT_URLS = {}
KNOWN_RESOLUTIONS = {}
KNOWN_USERS = {}
CATEGORY_MAPPING = {}

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

def get_workitem_json(url):
    try:
        response = session.get(url, headers=headers, verify=False)
        if response.status_code == 200:
            return response.json()
    except Exception as e:
        logger.error(f"Error fetching {url}: {e}")
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

def extract_date_field(json_data, field_keyword):
    if isinstance(json_data, dict):
        for key, value in json_data.items():
            if field_keyword in key.lower():
                if isinstance(value, str) and "T" in value and "-" in value:
                    return value
            if isinstance(value, dict):
                result = extract_date_field(value, field_keyword)
                if result: return result
    return None

def get_resolution_name(resolution_url):
    if not resolution_url:
        return ""
    if resolution_url in KNOWN_RESOLUTIONS:
        return KNOWN_RESOLUTIONS[resolution_url]
        
    res_data = get_workitem_json(resolution_url)
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
    if not user_url: return "Unassigned"
    if user_url in KNOWN_USERS: return KNOWN_USERS[user_url]
    
    fallback_id = str(user_url).rstrip('/').split('/')[-1]
    
    try:
        api_headers = {"Accept": "application/rdf+xml, text/xml", "OSLC-Core-Version": "2.0"}
        response = session.get(user_url, headers=api_headers, verify=False)
        
        if response.status_code == 200:
            xml_match = re.search(r'<(?:foaf:name|dc:title|name)[^>]*>([^<]+)</', response.text, re.IGNORECASE)
            if xml_match:
                KNOWN_USERS[user_url] = xml_match.group(1).strip()
                return KNOWN_USERS[user_url]
    except Exception:
        pass
        
    KNOWN_USERS[user_url] = fallback_id
    return fallback_id

def get_country_and_rate(owner_string):
    if "-AU" in owner_string: country = "Australia"
    elif "-JP" in owner_string: country = "Japan"
    elif "BGSW/ECC" in owner_string: country = "Mexico"
    elif "VM/ESB" in owner_string: country = "Germany"
    elif "MS/EHV" in owner_string: country = "Vietnam"
    elif "VM/EFO" in owner_string: country = "North America"
    elif "-NA" in owner_string: country = "North America"
    elif "Adecco, MS/EAS5-VM" in owner_string: country = "Vietnam"
    elif "MS/EJH53-VM" in owner_string: country = "Vietnam"
    elif "Adecco, MS/ETA-VOS-FBL" in owner_string: country = "Vietnam"
    elif "MS/" in owner_string: country = "India"
    else: country = "Others"
        
    rate_card = {
        "India": 50000, "Mexico": 72727, "Vietnam": 42545,
        "China": 80000, "Germany": 175000, "Hungary": 80000,
        "Japan": 80000, "North America": 175000, "Romania": 80000,
        "Portugal": 80000, "Austria": 80000, "Australia": 175000,
        "Others": 0
    }
    return country, rate_card.get(country, 0)

def get_department_name(department_url):
    if not department_url:
        return ""
    if department_url in KNOWN_DEPARTMENT_URLS:
        return KNOWN_DEPARTMENT_URLS[department_url]
    
    cat_data = get_workitem_json(department_url)
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
        
    if any(kw in name for kw in ["net", "fota", "var", "dpe", "hmi"]):
        return "NET"
    elif any(kw in name for kw in ["diag", "dcom", "sar"]):
        return "DCOM"
    elif any(kw in name for kw in ["obd", "dem", "dsm", "dws"]):
        return "DSM" 
        
    return None 

def process_hierarchy(work_item_url, visited=None, depth=0, is_rework_branch=False):
    if visited is None:
        visited = set()

    efforts = {
        "NET_DEV": 0, "DCOM_DEV": 0, "DSM_DEV": 0,
        "NET_Rework": 0, "DCOM_Rework": 0, "DSM_Rework": 0
    }

    if work_item_url in visited:
        return efforts
    visited.add(work_item_url)

    data = get_workitem_json(work_item_url)
    if not data:
        return efforts

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
    if not type_name_short:
        type_name_short = "unknown_type"

    time_spent_raw = data.get("rtc_cm:timeSpent")
    time_spent_ms = int(time_spent_raw) if time_spent_raw else 0

    current_is_rework = is_rework_branch or ("defect" in item_type_str)

    if time_spent_ms > 0:
        hours_logged = time_spent_ms / 3600000
        indent = "  " * depth
        
        filed_against_data = data.get("rtc_cm:filedAgainst", {})
        task_category_url = ""
        if isinstance(filed_against_data, dict):
            task_category_url = filed_against_data.get("rdf:resource", "")
        elif isinstance(filed_against_data, str):
            task_category_url = filed_against_data
            
        task_department_name = get_department_name(task_category_url)
        bucket = get_bucket_for_category(task_department_name)
        
        if bucket:
            bucket_key = f"{bucket}_Rework" if current_is_rework else f"{bucket}_DEV"
            efforts[bucket_key] += time_spent_ms
            rework_str = "REWORK" if current_is_rework else "DEV "
            logger.info(f"{indent}→ [{bucket} {rework_str}] Added {hours_logged:.2f} hrs | ID: {item_id} | Type: {type_name_short} | Dept: {task_department_name}")
        else:
            logger.debug(f"{indent}→ [IGNORED CATEGORY] Skipped {hours_logged:.2f} hrs | ID: {item_id} | Type: {type_name_short} | Dept: '{task_department_name}'")

    children_data = data.get("rtc_cm:com.ibm.team.workitem.linktype.parentworkitem.children")
    if not children_data:
        children_data = []
    elif isinstance(children_data, dict):
        children_data = [children_data]

    for child in children_data:
        if isinstance(child, dict) and "rdf:resource" in child:
            child_url = child["rdf:resource"]
            child_efforts = process_hierarchy(child_url, visited, depth + 1, current_is_rework)
            
            for key in efforts.keys():
                efforts[key] += child_efforts[key]

    return efforts

if __name__ == "__main__":
    logger.info("\n---> [NEW RUN STARTING: EXCEL EXPORT] <---")
    logger.info(f"Log file appending to: {LOG_FILE}\n")
    
    load_category_mapping()
    
    processed_rows = []
    
    try:
        # We still READ from the original CSV format
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
            logger.info(f"Sorting complete: {len(normal_rows)} standard rows, {len(delayed_rows)} delayed rows.\n")
            
            for row in ordered_rows:
                release_id = row.get("Id", "").strip()
                pm_id = row.get("PM Interface Element ID", "").strip()
                
                logger.info(f"=========================================")
                logger.info(f"Checking Release ID: {release_id} [PM ID: {pm_id}]...")
                root_url = f"{ALM_BASE_URL}/resource/itemName/com.ibm.team.workitem.WorkItem/{release_id}"
                
                release_data = get_workitem_json(root_url)
                if release_data:
                    res_info = release_data.get("rtc_cm:resolution")
                    res_url = ""
                    if isinstance(res_info, dict):
                        res_url = res_info.get("rdf:resource", "")
                    elif isinstance(res_info, str):
                        res_url = res_info
                        
                    res_status = "Unresolved/None"
                    if res_url:
                        res_status = get_resolution_name(res_url)
                        
                    if "invalid" in res_status:
                        logger.warning(f"❌ SKIPPING {release_id}: Resolution is '{res_status.title()}'")
                        logger.info(f"=========================================\n")
                        continue 
                    else:
                        logger.info(f"✔️ Resolution is '{res_status.title()}'.")

                    owner_info = release_data.get("rtc_cm:ownedBy") or release_data.get("dc:creator") or release_data.get("dcterms:creator") or release_data.get("ownedBy")
                    owner_url = ""
                    if isinstance(owner_info, dict):
                        owner_url = owner_info.get("rdf:resource", "")
                    elif isinstance(owner_info, str):
                        owner_url = owner_info
                    
                    full_owner_string = get_owner_details(owner_url)
                    country, rate = get_country_and_rate(full_owner_string)
                    
                    created_raw = extract_date_field(release_data, "created")
                    resolved_raw = extract_date_field(release_data, "resolved")
                    if not resolved_raw:
                        resolved_raw = extract_date_field(release_data, "modified")
                    
                    created_formatted = format_date(created_raw)
                    resolved_formatted = format_date(resolved_raw)
                    
                    logger.info(f"👤 Owned By: {full_owner_string}")
                    logger.info(f"🌍 Country: {country} | Rate: € {rate}")
                    logger.info(f"📅 Created: {created_formatted} | Resolved/Modified: {resolved_formatted}")
                    
                    row["Owner Details"] = full_owner_string
                    row["Country"] = country
                    row["Rate Card (€)"] = rate
                    row["Creation Date"] = created_formatted
                    row["Resolution Date"] = resolved_formatted

                logger.info(f"Processing Hierarchy for {release_id}...")
                efforts_ms = process_hierarchy(root_url)
                
                logger.info(f"\n--- Final Hours for Release {release_id} ---")
                
                net_dev = efforts_ms["NET_DEV"] / 3600000
                dcom_dev = efforts_ms["DCOM_DEV"] / 3600000
                dsm_dev = efforts_ms["DSM_DEV"] / 3600000
                
                net_rew = efforts_ms["NET_Rework"] / 3600000
                dcom_rew = efforts_ms["DCOM_Rework"] / 3600000
                dsm_rew = efforts_ms["DSM_Rework"] / 3600000
                
                if net_dev > 0: logger.info(f"  NET_DEV: {net_dev:.2f} hrs")
                if dcom_dev > 0: logger.info(f"  DCOM_DEV: {dcom_dev:.2f} hrs")
                if dsm_dev > 0: logger.info(f"  DSM_DEV: {dsm_dev:.2f} hrs")
                if net_rew > 0: logger.info(f"  NET_Rework: {net_rew:.2f} hrs")
                if dcom_rew > 0: logger.info(f"  DCOM_Rework: {dcom_rew:.2f} hrs")
                if dsm_rew > 0: logger.info(f"  DSM_Rework: {dsm_rew:.2f} hrs")

                net_final = net_dev + net_rew
                dcom_dsm_dev = dcom_dev + dsm_dev
                dcom_dsm_rew = dcom_rew + dsm_rew
                dcom_dsm_final = dcom_dsm_dev + dcom_dsm_rew

                row["NET_DEV (Hours)"] = round(net_dev, 2)
                row["DCOM_DEV (Hours)"] = round(dcom_dev, 2)
                row["DSM_DEV (Hours)"] = round(dsm_dev, 2)
                
                row["NET_FINAL (Hours)"] = round(net_final, 2)
                row["DCOM_DSM_DEV (Hours)"] = round(dcom_dsm_dev, 2)
                
                row["NET_Rework (Hours)"] = round(net_rew, 2)
                row["DCOM_Rework (Hours)"] = round(dcom_rew, 2)
                row["DSM_Rework (Hours)"] = round(dsm_rew, 2)
                
                row["DCOM_DSM_REWORK_TOTAL (Hours)"] = round(dcom_dsm_rew, 2)
                row["DCOM_DSM_FINAL (Hours)"] = round(dcom_dsm_final, 2)
                
                processed_rows.append(row)
                
        # --- EXPORT TO EXCEL ---
        if processed_rows:
            # Convert list of dictionaries to a Pandas DataFrame
            df = pd.DataFrame(processed_rows)
            # Export to Excel (.xlsx), removing the index column
            df.to_excel(OUTPUT_EXCEL_FILE, index=False)
                
            logger.info(f"\n✅ Success! Processed data saved to Excel file: '{OUTPUT_EXCEL_FILE}'.")
        else:
            logger.warning("\n⚠️ No rows were processed.")

    except FileNotFoundError:
        logger.error(f"Error: Could not find '{INPUT_CSV_FILE}'.")
    except ImportError:
        logger.error("\n❌ ERROR: Pandas or openpyxl is not installed!")
        logger.error("Please run: pip install pandas openpyxl")
