import re
import pandas as pd
from collections import defaultdict
import os

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_FILE = os.path.join(SCRIPT_DIR, "added_time_log.txt")
OUTPUT_FILE = os.path.join(SCRIPT_DIR, "release_hours.xlsx")

# ==========================
# REGEX PATTERNS
# ==========================

release_pattern = re.compile(
    r"Checking Release ID:\s*(\d+)\s*\[PM ID:\s*([^\]]+)\]"
)

owner_pattern = re.compile(
    r"👤 Release Owned By:\s*(.+?)\s*\(Tasks may be owned by others\)"
)

date_pattern = re.compile(
    r"📅 Created:\s*(\d{2}-\d{2}-\d{4})\s*\|\s*Resolved:\s*(\d{2}-\d{2}-\d{4})"
)

task_pattern = re.compile(
    r"\[(.*?)\]\s*Added\s*([\d.]+)\s*hrs\s*\|\s*"
    r"ID:\s*(\d+)\s*\|\s*"
    r"Type:\s*(.*?)\s*\|\s*"
    r"Dept:\s*'(.*?)'\s*\|\s*"
    r"Country:\s*(.*?)\s*\|\s*"
    r"Title:\s*(.*?)\s*\|\s*"
    r"Created:\s*(\d{2}-\d{2}-\d{4})"
)

# ==========================
# READ FILE
# ==========================

with open(INPUT_FILE, "r", encoding="utf-8") as f:
    lines = f.readlines()

# ==========================
# PARSE LOG
# ==========================

rows = []

current_release = None
current_pm = None
current_owner = None
current_created = None
current_resolved = None

for line in lines:

    line = line.strip()

    # Release
    m = release_pattern.search(line)
    if m:
        current_release = m.group(1)
        current_pm = m.group(2)
        continue

    # Owner
    m = owner_pattern.search(line)
    if m:
        current_owner = m.group(1)
        continue

    # Dates
    m = date_pattern.search(line)
    if m:
        current_created = m.group(1)
        current_resolved = m.group(2)
        continue

    # Task row
    m = task_pattern.search(line)
    if m:
        rows.append({
            "Release ID": current_release,
            "PM ID": current_pm,
            "Release Owner": current_owner,
            "Release Created": current_created,
            "Release Resolved": current_resolved,
            "Category": m.group(1).strip(),
            "Hours": float(m.group(2)),
            "Task ID": m.group(3),
            "Type": m.group(4),
            "Department": m.group(5),
            "Country": m.group(6),
            "Title": m.group(7),
            "Task Created": m.group(8)
        })

# ==========================
# DETAILED SHEET
# ==========================

detail_df = pd.DataFrame(rows)

if detail_df.empty:
    print("No task entries found.")
    exit()

# ==========================
# SUMMARY SHEET
# ==========================

summary_df = (
    detail_df
    .pivot_table(
        index=["Release ID"],
        columns="Category",
        values="Hours",
        aggfunc="sum",
        fill_value=0
    )
    .reset_index()
)

summary_df["Total Hours"] = (
    summary_df
    .drop(columns=["Release ID"])
    .sum(axis=1)
)

# ==========================
# WRITE EXCEL
# ==========================

with pd.ExcelWriter(
    OUTPUT_FILE,
    engine="openpyxl"
) as writer:

    detail_df.to_excel(
        writer,
        sheet_name="Detailed_Data",
        index=False
    )

    summary_df.to_excel(
        writer,
        sheet_name="Release_Summary",
        index=False
    )

print(f"Excel created: {OUTPUT_FILE}")