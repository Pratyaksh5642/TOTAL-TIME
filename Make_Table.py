import os
import pandas as pd

# 1. Dynamically get the directory where this script is located
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# 2. Join it with just the file name
# Make sure this matches your actual file name!
INPUT_EXCEL_FILE = os.path.join(SCRIPT_DIR, "Fiat_Final_with_countOfTask_hours.xlsx")

# 3. Read the Detailed_Data sheet using the dynamic path
df = pd.read_excel(INPUT_EXCEL_FILE, sheet_name='Detailed_Data')

# ---------------------------------------------------------
# FIXES FOR DATA INCONSISTENCIES:
# ---------------------------------------------------------
# A. Parse the creation date columns from DD-MM-YYYY format into Python datetime objects
df['Parsed_Task_Date'] = pd.to_datetime(df['Task Created'], dayfirst=True, errors='coerce')
df['Parsed_Release_Date'] = pd.to_datetime(df['Release Created'], dayfirst=True, errors='coerce')
# Use Task Created, and if it's missing, fall back to Release Created
df['Unified_Creation_Date'] = df['Parsed_Task_Date'].fillna(df['Parsed_Release_Date'])

# B. To avoid case sensitivity issues (e.g. 'solved' vs 'Solved'), force the column to lowercase
df['Release Status'] = df['Release Status'].astype(str).str.strip().str.lower()
# ---------------------------------------------------------

# Define target categories 
target_categories = [
    'DCOM DEV', 'NET DEV', 'DCOM REWORK', 
    'NET REWORK', 'DSM DEV', 'DSM REWORK'
]

orphan_roots = ['Defectfix', 'Task', 'Review', 'Epic', 'Story']

# 4. Automatically find all unique countries in your data
# We use dropna() to ignore any empty blank rows
countries = df['Country'].dropna().unique().tolist()
print(f"Found {len(countries)} countries to process: {', '.join(countries)}")

# Add 'WW' (Worldwide) to the end of our list to process all countries combined
countries_and_ww = countries + ['WW']
print(f"Adding 'WW' sheet for all countries combined...\n")

# --- INITIALIZE THE MASTER WW_TABLE ---
# Since you only want the totals, we initialize a simple 1-row dataframe
ww_table_data = {
    "Metric": ["Total Effort (hrs)"]
}
ww_df = pd.DataFrame(ww_table_data)

# 5. Open the Excel writer ONCE to save multiple sheets
with pd.ExcelWriter(INPUT_EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    
    # Loop through each country automatically, ending with 'WW'
    for country in countries_and_ww:
        
        # If it's WW, our country mask is just True (selects everything). 
        # Otherwise, we filter by the specific country.
        if country == 'WW':
            c_mask = True
            c_desc = "WW (All Countries)"
            col_name = "Total" # Call the WW column "Total" in our final WW_Table
        else:
            c_mask = (df['Country'] == country)
            c_desc = f"'{country}'"
            col_name = country

        # Base filters 
        mask_base = (
            c_mask &
            (df['Year'] == 2026) &
            (df['Month'].isin([1, 2, 3, 4, 5, 6])) &
            (df['Category'].isin(target_categories))
        )

        # --- Row 1: Closed Release IDs only ---
        closed_statuses = ['solved'] 
        mask_closed = mask_base & (df['Root Item Type'] == 'Release') & (df['Release Status'].isin(closed_statuses))
        closed_hours = df.loc[mask_closed, 'Hours'].sum()
        desc_closed = f"Country={c_desc}, Year=2026, Month in 1-6, Category in {target_categories}, Root Item Type='Release', Release Status in {closed_statuses}"

        # --- Row 2: New + In Progress Release IDs ---
        unresolved_statuses = ['unresolved (new / in progress)']
        mask_unresolved = mask_base & (df['Root Item Type'] == 'Release') & (df['Release Status'].isin(unresolved_statuses))
        unresolved_hours = df.loc[mask_unresolved, 'Hours'].sum()
        desc_unresolved = f"Country={c_desc}, Year=2026, Month in 1-6, Category in {target_categories}, Root Item Type='Release', Release Status in {unresolved_statuses}"

        # --- Row 3: Linked to release ID (Updated with Datetime logic) ---
        mask_linked = (
            c_mask & 
            (df['Year'] == 2027) & 
            (df['Root Item Type'] == 'Release') &
            (df['Unified_Creation_Date'].dt.year == 2026) &
            (df['Unified_Creation_Date'].dt.month.isin([1, 2, 3, 4, 5, 6]))
        )
        hours_linked = df.loc[mask_linked, 'Hours'].sum()
        desc_linked = f"Country={c_desc}, Year=2027, Root Item Type='Release', Creation Date in 2026 (Jan-Jun)"

        # --- Row 4: Linked to release ID- & Not linked to release ID (Miscategorized) ---
        mask_misc = (
            c_mask & 
            (df['Year'] == 2026) & 
            (df['Month'].isin([1, 2, 3, 4, 5, 6])) & 
            (df['Category'] == 'GENERAL Miscategorized')
        )
        hours_misc = df.loc[mask_misc, 'Hours'].sum()
        desc_misc = f"Country={c_desc}, Year=2026, Month in 1-6, Category='GENERAL Miscategorized'"

        # --- Row 5: Not linked to release ID (Defect) ---
        mask_defect_1 = (
            c_mask & 
            (df['Year'] == 2026) & 
            (df['Month'].isin([1, 2, 3, 4, 5, 6])) & 
            (df['Root Item Type'] == 'Defect') &
            (df['Category'] != 'GENERAL Miscategorized')
        )
        mask_defect_2 = (
            c_mask & 
            (df['Year'] == 2027) & 
            (df['Root Item Type'] == 'Defect')
        )
        hours_defect = df.loc[mask_defect_1 | mask_defect_2, 'Hours'].sum()
        desc_defect = f"(Country={c_desc}, Year=2026, Month 1-6, Root Item='Defect', Category != 'GENERAL Miscategorized') + (Country={c_desc}, Year=2027, Root Item='Defect')"

        # --- Row 6: Not linked to release ID (Orphan) ---
        mask_orphan_1 = (
            c_mask & 
            (df['Year'] == 2026) & 
            (df['Month'].isin([1, 2, 3, 4, 5, 6])) & 
            (df['Root Item Type'].isin(orphan_roots)) & 
            (df['Category'] != 'GENERAL Miscategorized')
        )
        mask_orphan_2 = (
            c_mask & 
            (df['Year'] == 2027) & 
            (df['Root Item Type'].isin(orphan_roots))
        )
        hours_orphan = df.loc[mask_orphan_1 | mask_orphan_2, 'Hours'].sum()
        desc_orphan = f"(Country={c_desc}, Year=2026, Month 1-6, Root Item in {orphan_roots}, Category != 'GENERAL Miscategorized') + (Country={c_desc}, Year=2027, Root Item in {orphan_roots})"

        # Initialize the Results Table for this specific country / WW
        results = [
            {"Scenario": "Closed Release IDs only", "Release ID Status / Explanation": "Closed", "Effort in hrs": closed_hours, "Description": desc_closed},
            {"Scenario": "All relevant Release IDs", "Release ID Status / Explanation": "New + In Progress (Release Ids)", "Effort in hrs": unresolved_hours, "Description": desc_unresolved},
            {"Scenario": "Linked to release ID", "Release ID Status / Explanation": "New + In Progress (Task Ids)", "Effort in hrs": hours_linked, "Description": desc_linked},
            {"Scenario": "Linked to release ID- & Not linked to release ID", "Release ID Status / Explanation": "WIs are not mapped to the correct filed against", "Effort in hrs": hours_misc, "Description": desc_misc},
            {"Scenario": "Not linked to release ID", "Release ID Status / Explanation": "Defect", "Effort in hrs": hours_defect, "Description": desc_defect},
            {"Scenario": "Not linked to release ID", "Release ID Status / Explanation": "Orphan", "Effort in hrs": hours_orphan, "Description": desc_orphan}
        ]

        # Create DataFrame
        results_df = pd.DataFrame(results)

        # KEEP AS ACTUAL NUMBERS: Round to 0 decimals instead of converting to text strings
        results_df['Effort in hrs'] = results_df['Effort in hrs'].round(0)
        
        # SUM THE COLUMN AND ADD IT TO THE WW_TABLE (Just storing the Total)
        total_effort = results_df['Effort in hrs'].sum()
        ww_df[col_name] = [total_effort]

        # Clean up the sheet name (Excel sheets have a 31 character limit and ban certain special characters)
        sheet_name = str(country)[:31]
        for char in ['[', ']', ':', '*', '?', '/', '\\']:
            sheet_name = sheet_name.replace(char, '_')

        # Save this table to its own sheet
        results_df.to_excel(writer, sheet_name=sheet_name, index=False)
        print(f"Saved sheet for: {sheet_name} (Total: {total_effort})")
    
    # Finally, save the new summary WW_Table to its own sheet
    ww_df.to_excel(writer, sheet_name='WW_Table', index=False)
    print("Saved master summary sheet for: WW_Table")

print("\nSuccess! All sheets and the final 'WW_Table' summary have been successfully added to your Excel file.")
