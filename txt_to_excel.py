import os
import re
import pandas as pd

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

INPUT_FILE = os.path.join(SCRIPT_DIR, "Add_time.txt")
OUTPUT_FILE = os.path.join(SCRIPT_DIR, "release_hours_with test.xlsx")

# ====================================================
# PATTERNS
# ====================================================

release_pattern = re.compile(
    r"Checking Release ID:\s*(\d+)\s*\[PM ID:\s*([^\]]+)\]"
)

owner_pattern = re.compile(
    r"👤 Release Owned By:\s*(.+?)\s*"
    r"\(Tasks may be owned by others\)"
)

date_pattern = re.compile(
    r"📅 Created:\s*(\d{2}-\d{2}-\d{4})\s*\|\s*"
    r"Resolved:\s*(\d{2}-\d{2}-\d{4})"
)

task_pattern = re.compile(
    r"\[(.*?)\]\s*Added\s*([\d.]+)\s*hrs\s*\|\s*"
    r"ID:\s*(\d+)\s*\|\s*"
    r"Type:\s*(.*?)\s*\|\s*"
    r"Dept:\s*'(.*?)'\s*\|\s*"
    r"Owner:\s*(.*?)\s*\|\s*"
    r"Country:\s*(.*?)\s*\|\s*"
    r"Title:\s*(.*?)\s*\|\s*"
    r"Created:\s*(\d{2}-\d{2}-\d{4})"
)

# ====================================================
# READ FILE
# ====================================================

rows = []

current_release = ""
current_pm = ""
current_release_owner = ""
current_release_created = ""
current_release_resolved = ""

with open(INPUT_FILE, "r", encoding="utf-8") as f:

    for line in f:

        line = (
            line.replace("<br>", "")
            .replace("&lt;br&gt;", "")
            .replace("&nbsp;", " ")
            .strip()
        )

        if not line:
            continue

        # --------------------------------------------
        # RELEASE ID
        # --------------------------------------------

        match = release_pattern.search(line)

        if match:
            current_release = match.group(1)
            current_pm = match.group(2)
            continue

        # --------------------------------------------
        # RELEASE OWNER
        # --------------------------------------------

        match = owner_pattern.search(line)

        if match:
            current_release_owner = match.group(1)
            continue

        # --------------------------------------------
        # RELEASE DATES
        # --------------------------------------------

        match = date_pattern.search(line)

        if match:
            current_release_created = match.group(1)
            current_release_resolved = match.group(2)
            continue

        # --------------------------------------------
        # TASKS
        # --------------------------------------------

        match = task_pattern.search(line)

        if match:
            rows.append({
                "Release ID": current_release,
                "PM ID": current_pm,
                "Release Owner": current_release_owner,
                "Release Created": current_release_created,
                "Release Resolved": current_release_resolved,
                "Category": match.group(1).strip(),
                "Hours": float(match.group(2)),
                "Task ID": match.group(3),
                "Type": match.group(4).strip(),
                "Department": match.group(5).strip(),
                "Task Owner": match.group(6).strip(),
                "Country": match.group(7).strip(),
                "Title": match.group(8).strip(),
                "Task Created": match.group(9)
            })

# ====================================================
# DATAFRAME
# ====================================================

detail_df = pd.DataFrame(rows)

if detail_df.empty:
    print("No data found.")
    raise SystemExit

# ====================================================
# TASK CREATED DATE
# ====================================================

task_created_date = pd.to_datetime(
    detail_df["Task Created"],
    format="%d-%m-%Y",
    errors="coerce"
)

# Day, month and year are taken from Task Created
detail_df["Created_Day"] = task_created_date.dt.day.astype("Int64")
detail_df["Created_Month"] = task_created_date.dt.month.astype("Int64")
detail_df["Created_Year"] = task_created_date.dt.year.astype("Int64")

# Keep Task Created as DD-MM-YYYY without 00:00:00
detail_df["Task Created"] = task_created_date.dt.strftime("%d-%m-%Y")

# ====================================================
# SUMMARY
# ====================================================

summary_df = (
    detail_df
    .pivot_table(
        index="Release ID",
        columns="Category",
        values="Hours",
        aggfunc="sum",
        fill_value=0
    )
    .reset_index()
)

numeric_cols = summary_df.select_dtypes(
    include="number"
).columns

summary_df["Total Hours"] = summary_df[
    numeric_cols
].sum(axis=1)

# ====================================================
# WRITE EXCEL
# ====================================================

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

print()
print("Excel created successfully")
print(f"Rows exported: {len(detail_df)}")
print(OUTPUT_FILE)