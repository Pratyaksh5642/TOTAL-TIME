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

# Get the script's directory to load files relative to it
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_CSV_FILE = os.path.join(SCRIPT_DIR, "input.csv")
OUTPUT_CSV_FILE = os.path.join(SCRIPT_DIR, "output.csv")

# --- SETUP SESSION ---
session = requests.Session()
session.auth = (USERNAME, PASSWORD)
headers = {
    "Accept": "application/json",
    "OSLC-Core-Version": "2.0"
}

def get_workitem_json(url):
    """Fetches the JSON data for a specific work item."""
    try:
        response = session.get(url, headers=headers, verify=False)
        if response.status_code == 200:
            return response.json()
    except Exception as e:
        print(f"Error fetching {url}: {e}")
    return None

def process_hierarchy(work_item_url, visited=None):
    """
    Recursively crawls down the work item tree.
    Returns: (total_effort_ms, rework_effort_ms)
    """
    if visited is None:
        visited = set()

    # Prevent infinite loops if ALM has cyclical links
    if work_item_url in visited:
        return 0, 0
    visited.add(work_item_url)

    data = get_workitem_json(work_item_url)
    if not data:
        return 0, 0

    total_effort_ms = 0
    rework_effort_ms = 0

    # 1. Extract Time Spent
    # Ensure we handle None or missing values safely
    time_spent_raw = data.get("rtc_cm:timeSpent")
    time_spent_ms = int(time_spent_raw) if time_spent_raw else 0

    # 2. Check the Item Type (to separate Defects from Tasks/Epics/Releases)
    item_type_str = ""
    dc_type = data.get("dc:type", {})
    if isinstance(dc_type, dict):
        item_type_str = dc_type.get("rdf:resource", "").lower()
    elif isinstance(dc_type, str):
        item_type_str = dc_type.lower()

    # Apply calculation logic
    if "defect" in item_type_str:
        rework_effort_ms += time_spent_ms
    else:
        total_effort_ms += time_spent_ms

    # 3. Find and recursively crawl Child Links
    children_data = data.get("rtc_cm:com.ibm.team.workitem.linktype.parentworkitem.children", [])
    
    # ALM JSON usually returns a list of dictionaries with 'rdf:resource' for the child URLs
    if isinstance(children_data, dict):
        children_data = [children_data] # Normalize to list if ALM returns a single dict

    for child in children_data:
        if isinstance(child, dict) and "rdf:resource" in child:
            child_url = child["rdf:resource"]
            
            # Recursive drill-down
            child_total, child_rework = process_hierarchy(child_url, visited)
            
            # Bubble up the sums from the bottom of the tree
            total_effort_ms += child_total
            rework_effort_ms += child_rework

    return total_effort_ms, rework_effort_ms

if __name__ == "__main__":
    print("Starting ALM Effort Extraction...")
    
    # Prepare the output CSV data
    processed_rows = []
    
    try:
        # Open your input CSV
        with open(INPUT_CSV_FILE, mode="r", encoding="utf-16") as infile:
            reader = csv.DictReader(infile, delimiter="\t")
            
            # Ensure your input CSV has these exact column headers!
            # Example headers: "Release ID", "PM Interface Element ID"
            for row in reader:
                release_id = row.get("Id", "").strip()
                pm_id = row.get("PM Interface Element ID", "").strip()
                
                if not release_id:
                    continue
                
                print(f"Processing Release ID: {release_id}...")
                root_url = f"{ALM_BASE_URL}/resource/itemName/com.ibm.team.workitem.WorkItem/{release_id}"
                
                # Do the recursive calculation
                total_ms, rework_ms = process_hierarchy(root_url)
                
                # Convert milliseconds to readable hours
                total_hours = total_ms / 3600000
                rework_hours = rework_ms / 3600000
                
                # Append to our new row data
                row["Total Effort (Hours)"] = round(total_hours, 2)
                row["Rework Effort (Hours)"] = round(rework_hours, 2)
                processed_rows.append(row)
                
        # Save the results to the output CSV
        if processed_rows:
            # We grab the column headers from the first processed row
            fieldnames = list(processed_rows[0].keys())
            
            with open(OUTPUT_CSV_FILE, mode="w", newline="", encoding="utf-8") as outfile:
                writer = csv.DictWriter(outfile, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(processed_rows)
                
            print(f"\n✅ Success! All data processed and saved to '{OUTPUT_CSV_FILE}'.")
        else:
            print("\n⚠️ No rows were processed. Please check your input CSV format.")

    except FileNotFoundError:
        print(f"Error: Could not find '{INPUT_CSV_FILE}'. Please ensure the file is in the same folder as this script.")
