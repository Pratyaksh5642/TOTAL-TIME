import os
import pandas as pd

# 1. Dynamically get the directory where this script is located
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# 2. Join it with just the file name
INPUT_EXCEL_FILE = os.path.join(SCRIPT_DIR, "Final_Corrected_hours.xlsx")

# 3. Read the Detailed_Data sheet using the dynamic path
df = pd.read_excel(INPUT_EXCEL_FILE, sheet_name='Detailed_Data')

# Define target categories (used ONLY in rows 1 and 2)
target_categories = [
    'DCOM DEV', 'NET DEV', 'DCOM REWORK', 
    'NET REWORK', 'DSM DEV', 'DSM REWORK'
]

# Base filters for Rows 1 and 2
mask_base = (
    (df['Country'] == 'India') &
    (df['Year'] == 2026) &
    (df['Month'].isin([1, 2, 3, 4, 5, 6])) &
    (df['Category'].isin(target_categories))
)

# --- Row 1: Closed Release IDs only ---
closed_statuses = ['Solved'] 
mask_closed = mask_base & (df['Root Item Type'] == 'Release') & (df['Release Status'].isin(closed_statuses))
closed_hours = df.loc[mask_closed, 'Hours'].sum()

desc_closed = (
    f"Country='India', Year=2026, Month in 1-6, "
    f"Category in {target_categories}, "
    f"Root Item Type='Release', Release Status in {closed_statuses}"
)

# --- Row 2: New + In Progress Release IDs ---
unresolved_statuses = ['Unresolved (New / In Progress)']
mask_unresolved = mask_base & (df['Root Item Type'] == 'Release') & (df['Release Status'].isin(unresolved_statuses))
unresolved_hours = df.loc[mask_unresolved, 'Hours'].sum()

desc_unresolved = (
    f"Country='India', Year=2026, Month in 1-6, "
    f"Category in {target_categories}, "
    f"Root Item Type='Release', Release Status in {unresolved_statuses}"
)

# --- Row 3: Linked to release ID (Updated) ---
mask_linked = (
    (df['Country'] == 'India') & 
    (df['Year'] == 2027) & 
    (df['Root Item Type'] == 'Release')
)
hours_linked = df.loc[mask_linked, 'Hours'].sum()

desc_linked = "Country='India', Year=2027, Root Item Type='Release'"

# --- Row 4: Linked to release ID- & Not linked to release ID (Miscategorized) ---
mask_misc = (
    (df['Country'] == 'India') & 
    (df['Year'] == 2026) & 
    (df['Month'].isin([1, 2, 3, 4, 5, 6])) & 
    (df['Category'] == 'GENERAL Miscategorized')
)
hours_misc = df.loc[mask_misc, 'Hours'].sum()

desc_misc = "Country='India', Year=2026, Month in 1-6, Category='GENERAL Miscategorized'"

# --- Row 5: Not linked to release ID (Defect) ---
# Condition 1: 2026 without Miscategorized
mask_defect_1 = (
    (df['Country'] == 'India') & 
    (df['Year'] == 2026) & 
    (df['Month'].isin([1, 2, 3, 4, 5, 6])) & 
    (df['Root Item Type'] == 'Defect') &
    (df['Category'] != 'GENERAL Miscategorized')
)
# Condition 2: 2027
mask_defect_2 = (
    (df['Country'] == 'India') & 
    (df['Year'] == 2027) & 
    (df['Root Item Type'] == 'Defect')
)
# Add hours from both conditions
hours_defect = df.loc[mask_defect_1 | mask_defect_2, 'Hours'].sum()

desc_defect = (
    "(Country='India', Year=2026, Month 1-6, Root Item='Defect', Category != 'GENERAL Miscategorized') "
    "+ (Country='India', Year=2027, Root Item='Defect')"
)

# --- Row 6: Not linked to release ID (Orphan) ---
orphan_roots = ['Defectfix', 'Task', 'Review', 'Epic', 'Story']
# Condition 1: 2026 without Miscategorized
mask_orphan_1 = (
    (df['Country'] == 'India') & 
    (df['Year'] == 2026) & 
    (df['Month'].isin([1, 2, 3, 4, 5, 6])) & 
    (df['Root Item Type'].isin(orphan_roots)) & 
    (df['Category'] != 'GENERAL Miscategorized')
)
# Condition 2: 2027
mask_orphan_2 = (
    (df['Country'] == 'India') & 
    (df['Year'] == 2027) & 
    (df['Root Item Type'].isin(orphan_roots))
)
# Add hours from both conditions
hours_orphan = df.loc[mask_orphan_1 | mask_orphan_2, 'Hours'].sum()

desc_orphan = (
    f"(Country='India', Year=2026, Month 1-6, Root Item in {orphan_roots}, Category != 'GENERAL Miscategorized') "
    f"+ (Country='India', Year=2027, Root Item in {orphan_roots})"
)

# Initialize the Results Table with the reordered scenarios
results = [
    {
        "Scenario": "Closed Release IDs only",
        "Release ID Status / Explanation": "Closed",
        "Effort in hrs": closed_hours,
        "Description": desc_closed
    },
    {
        "Scenario": "All relevant Release IDs", 
        "Release ID Status / Explanation": "New + In Progress (Release Ids)",
        "Effort in hrs": unresolved_hours,
        "Description": desc_unresolved
    },
    {
        "Scenario": "Linked to release ID",
        "Release ID Status / Explanation": "New + In Progress (Task Ids)",
        "Effort in hrs": hours_linked,
        "Description": desc_linked
    },
    {
        "Scenario": "Linked to release ID- & Not linked to release ID",
        "Release ID Status / Explanation": "WIs are not mapped to the correct filed against",
        "Effort in hrs": hours_misc,
        "Description": desc_misc
    },
    {
        "Scenario": "Not linked to release ID",
        "Release ID Status / Explanation": "Defect",
        "Effort in hrs": hours_defect,
        "Description": desc_defect
    },
    {
        "Scenario": "Not linked to release ID",
        "Release ID Status / Explanation": "Orphan",
        "Effort in hrs": hours_orphan,
        "Description": desc_orphan
    }
]

# Create DataFrame
results_df = pd.DataFrame(results)

# Format the Output Column with thousands separator and remove decimals
results_df['Effort in hrs'] = results_df['Effort in hrs'].apply(lambda x: f"{x:,.0f}")

# Display the Table in the console
print("Generated Table:\n")
print(results_df.to_string(index=False))

# Save the table to a new sheet named "table" in the SAME Excel file
with pd.ExcelWriter(INPUT_EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    results_df.to_excel(writer, sheet_name='table', index=False)
    
print("\nSuccess! The table has been saved to the 'table' sheet in your Excel file.")
