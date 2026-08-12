import pandas as pd
import os

# ==========================================
# PUT YOUR EXCEL FILE NAME HERE
# ==========================================
FILE_NAME = "Book3 (1).xlsx"

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_FILE = os.path.join(SCRIPT_DIR, FILE_NAME)

# Read sheets
export_df = pd.read_excel(INPUT_FILE, sheet_name="Export")
sheet1_df = pd.read_excel(INPUT_FILE, sheet_name="Sheet1")

# Keep only valid numeric IDs
export_df = export_df[
    pd.to_numeric(export_df["TaskOrReview ID"], errors="coerce").notna()
].copy()

sheet1_df = sheet1_df[
    pd.to_numeric(sheet1_df["Task ID"], errors="coerce").notna()
].copy()

# Convert IDs so:
# 5433090
# 5433090.0
# 5433090.00
# all become "5433090"
export_df["COMPARE_ID"] = (
    pd.to_numeric(export_df["TaskOrReview ID"], errors="coerce")
    .astype("Int64")
    .astype(str)
)

sheet1_df["COMPARE_ID"] = (
    pd.to_numeric(sheet1_df["Task ID"], errors="coerce")
    .astype("Int64")
    .astype(str)
)

# Get unique IDs
export_ids = set(export_df["COMPARE_ID"])
sheet1_ids = set(sheet1_df["COMPARE_ID"])

# IDs present only in one sheet
export_only = export_ids - sheet1_ids
sheet1_only = sheet1_ids - export_ids

# Rows from Export whose TaskOrReview ID is not in Sheet1
export_unmatched = export_df[
    export_df["COMPARE_ID"].isin(export_only)
].copy()

export_unmatched["Source Sheet"] = "Export"

# Rows from Sheet1 whose Task ID is not in Export
sheet1_unmatched = sheet1_df[
    sheet1_df["COMPARE_ID"].isin(sheet1_only)
].copy()

sheet1_unmatched["Source Sheet"] = "Sheet1"

# Remove helper column
export_unmatched.drop(columns=["COMPARE_ID"], inplace=True)
sheet1_unmatched.drop(columns=["COMPARE_ID"], inplace=True)

# Combine results
result_df = pd.concat(
    [export_unmatched, sheet1_unmatched],
    ignore_index=True,
    sort=False
)

# Save output
OUTPUT_FILE = os.path.join(SCRIPT_DIR, "Compare_task.xlsx")

with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    result_df.to_excel(
        writer,
        sheet_name="TaskID_Differences",
        index=False
    )

print("=" * 50)
print("Comparison Complete")
print("=" * 50)
print(f"Export-only IDs : {len(export_only)}")
print(f"Sheet1-only IDs : {len(sheet1_only)}")
print(f"Total rows saved: {len(result_df)}")
print(f"Output file     : {OUTPUT_FILE}")