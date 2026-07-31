import csv
import requests
import urllib3
import os

# Disable SSL warnings for internal domains
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# --- CONFIGURATION ---
ALM_BASE_URL = "https://rb-alm-06-p.de.bosch.com/ccm"
USERNAME = "lop2cob"
PASSWORD = "shreyansh4991Ab#"

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_CSV_FILE = os.path.join(SCRIPT_DIR, "input.csv")
OUTPUT_CSV_FILE = os.path.join(SCRIPT_DIR, "output_Solved_Invalid.csv")

session = requests.Session()
session.auth = (USERNAME, PASSWORD)
headers = {
    "Accept": "application/json",
    "OSLC-Core-Version": "2.0"
}

# Caches to prevent duplicate network requests
KNOWN_DEPARTMENT_URLS = {}
KNOWN_RESOLUTIONS = {}

def get_workitem_json(url):
    try:
        response = session.get(url, headers=headers, verify=False)
        if response.status_code == 200:
            return response.json()
    except Exception as e:
        print(f"Error fetching {url}: {e}")
    return None

def get_resolution_name(resolution_url):
    """Fetches the exact resolution status name (e.g. 'Solved', 'Invalid')"""
    if not resolution_url:
        return ""
        
    if resolution_url in KNOWN_RESOLUTIONS:
        return KNOWN_RESOLUTIONS[resolution_url]
        
    res_data = get_workitem_json(resolution_url)
    title = ""
    if res_data:
        title = res_data.get("dc:title", "")
        
    title_lower = title.lower()
    KNOWN_RESOLUTIONS[resolution_url] = title_lower
    return title_lower

def get_department_name(department_url):
    """Fetches the exact department/category name for a specific task's URL"""
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
    """Sorts the department into your specific buckets."""
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
            print(f"{indent}→ [{bucket} {rework_str}] Added {hours_logged:.2f} hrs | ID: {item_id} | Type: {type_name_short} | Dept: {task_department_name}")
        else:
            pass # Silently skip unmatched departments to keep console clean

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
    print("\n---> [RUNNING CATEGORIZED & FILTERED VERSION] <---")
    print("Starting ALM Effort Extraction...\n")
    
    processed_rows = []
    
    try:
        with open(INPUT_CSV_FILE, mode="r", encoding="utf-16") as infile:
            reader = csv.DictReader(infile, delimiter="\t")
            
            for row in reader:
                release_id = row.get("Id", "").strip()
                if not release_id:
                    continue
                
                print(f"=========================================")
                print(f"Checking Release ID: {release_id}...")
                root_url = f"{ALM_BASE_URL}/resource/itemName/com.ibm.team.workitem.WorkItem/{release_id}"
                
                # --- NEW: CHECK RESOLUTION BEFORE PROCESSING ---
                release_data = get_workitem_json(root_url)
                if release_data:
                    res_info = release_data.get("rtc_cm:resolution")
                    res_url = ""
                    if isinstance(res_info, dict):
                        res_url = res_info.get("rdf:resource", "")
                    elif isinstance(res_info, str):
                        res_url = res_info
                        
                    if res_url:
                        res_status = get_resolution_name(res_url)
                        if "invalid" in res_status:
                            print(f"❌ Skipping {release_id}: Resolution is '{res_status.title()}'")
                            continue # Skip to the next row in the CSV entirely!
                
                print(f"✅ Processing Hierarchy for {release_id}...")
                print(f"=========================================")
                
                # Do the recursive calculation
                efforts_ms = process_hierarchy(root_url)
                
                print(f"\n--- Final Hours for Release {release_id} ---")
                
                for key, ms_val in efforts_ms.items():
                    hours = ms_val / 3600000
                    row[f"{key} (Hours)"] = round(hours, 2)
                    
                    if hours > 0:
                        print(f"  {key}: {hours:.2f} hrs")
                
                processed_rows.append(row)
                
        if processed_rows:
            fieldnames = list(processed_rows[0].keys())
            with open(OUTPUT_CSV_FILE, mode="w", newline="", encoding="utf-8") as outfile:
                writer = csv.DictWriter(outfile, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(processed_rows)
                
            print(f"\n✅ Success! Filtered & Categorized data saved to '{OUTPUT_CSV_FILE}'.")
        else:
            print("\n⚠️ No rows were processed.")

    except FileNotFoundError:
        print(f"Error: Could not find '{INPUT_CSV_FILE}'.")
