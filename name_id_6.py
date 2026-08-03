import csv
import requests
import urllib3
import os
import logging
import re  # <--- Added for HTML scraping

# Disable SSL warnings for internal domains
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# --- CONFIGURATION ---
ALM_BASE_URL = "https://rb-alm-06-p.de.bosch.com/ccm"
USERNAME = "lop2cob"
PASSWORD = "shreyansh4991Ab#"

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_CSV_FILE = os.path.join(SCRIPT_DIR, "input.csv")
OUTPUT_CSV_FILE = os.path.join(SCRIPT_DIR, "output_categorized.csv")
LOG_FILE = os.path.join(SCRIPT_DIR, "extraction_log.txt")

# --- SETUP LOGGING ---
logger = logging.getLogger("alm_extractor")
logger.setLevel(logging.DEBUG) 
formatter = logging.Formatter('%(message)s')

file_handler = logging.FileHandler(LOG_FILE, mode='w', encoding='utf-8')
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

def get_workitem_json(url):
    try:
        response = session.get(url, headers=headers, verify=False)
        if response.status_code == 200:
            return response.json()
    except Exception as e:
        logger.error(f"Error fetching {url}: {e}")
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


    """TEST 1: Pure JSON Extraction"""
    if not user_url: return "Unassigned"
    if user_url in KNOWN_USERS: return KNOWN_USERS[user_url]
    
    fallback_id = str(user_url).rstrip('/').split('/')[-1]
    
    try:
        api_headers = {"Accept": "application/json", "OSLC-Core-Version": "2.0"}
        response = session.get(user_url, headers=api_headers, verify=False)
        
        if response.status_code == 200:
            user_data = response.json()
            found_name = user_data.get("name") or user_data.get("dc:title") or user_data.get("foaf:name")
            if found_name:
                KNOWN_USERS[user_url] = found_name.strip()
                return KNOWN_USERS[user_url]
    except Exception:
        pass 
        
    KNOWN_USERS[user_url] = fallback_id
    return fallback_id

def get_owner_details(user_url):
    """TEST 2: Pure XML/RDF Extraction"""
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
    if any(kw in name for kw in ["net", "fota", "var", "dpe", "hmi"]):
        return "NET"
    elif any(kw in name for kw in ["diag", "dcom", "sar"]):
        return "DCOM"
    elif any(kw in name for kw in ["obd", "dem", "dsm", "dws"]):
        return "DEM"
    return None 

def process_hierarchy(work_item_url, visited=None, depth=0, is_rework_branch=False):
    if visited is None:
        visited = set()

    efforts = {
        "NET_Total": 0, "DCOM_Total": 0, "DEM_Total": 0,
        "NET_Rework": 0, "DCOM_Rework": 0, "DEM_Rework": 0
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
            bucket_key = f"{bucket}_Rework" if current_is_rework else f"{bucket}_Total"
            efforts[bucket_key] += time_spent_ms
            rework_str = "REWORK" if current_is_rework else "TOTAL "
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
    logger.info("\n---> [RUNNING HTML SCRAPER VERSION] <---")
    logger.info(f"Log file being generated at: {LOG_FILE}\n")
    
    processed_rows = []
    
    try:
        with open(INPUT_CSV_FILE, mode="r", encoding="utf-16") as infile:
            reader = csv.DictReader(infile, delimiter="\t")
            
            for row in reader:
                release_id = row.get("Id", "").strip()
                if not release_id:
                    continue
                
                logger.info(f"=========================================")
                logger.info(f"Checking Release ID: {release_id}...")
                root_url = f"{ALM_BASE_URL}/resource/itemName/com.ibm.team.workitem.WorkItem/{release_id}"
                
                release_data = get_workitem_json(root_url)
                if release_data:
                    # 1. Check Resolution
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

                    # 2. Extract Owner via Regex Scraping
                    owner_info = release_data.get("rtc_cm:ownedBy") or release_data.get("dc:creator") or release_data.get("dcterms:creator") or release_data.get("ownedBy")
                    owner_url = ""
                    if isinstance(owner_info, dict):
                        owner_url = owner_info.get("rdf:resource", "")
                    elif isinstance(owner_info, str):
                        owner_url = owner_info
                    
                    full_owner_string = get_owner_details(owner_url)
                    logger.info(f"👤 Owned By: {full_owner_string}")
                    
                    row["Owner Details"] = full_owner_string

                logger.info(f"Processing Hierarchy for {release_id}...")
                efforts_ms = process_hierarchy(root_url)
                
                logger.info(f"\n--- Final Hours for Release {release_id} ---")
                
                for key, ms_val in efforts_ms.items():
                    hours = ms_val / 3600000
                    row[f"{key} (Hours)"] = round(hours, 2)
                    
                    if hours > 0:
                        logger.info(f"  {key}: {hours:.2f} hrs")
                
                processed_rows.append(row)
                
        if processed_rows:
            fieldnames = list(processed_rows[0].keys())
            with open(OUTPUT_CSV_FILE, mode="w", newline="", encoding="utf-8") as outfile:
                writer = csv.DictWriter(outfile, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(processed_rows)
                
            logger.info(f"\n✅ Success! Processed data saved to '{OUTPUT_CSV_FILE}'.")
        else:
            logger.warning("\n⚠️ No rows were processed.")

    except FileNotFoundError:
        logger.error(f"Error: Could not find '{INPUT_CSV_FILE}'.")
