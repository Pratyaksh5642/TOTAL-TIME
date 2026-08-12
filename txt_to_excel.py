import os
import re
import pandas as pd

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

INPUT_FILE = os.path.join(SCRIPT_DIR, "added_time_log.txt")
OUTPUT_FILE = os.path.join(SCRIPT_DIR, "release_hours.xlsx")

rows = []

current_release = ""
current_pm = ""
current_release_owner = ""
current_release_created = ""
current_release_resolved = ""

# ==================================================
# RELEASE PATTERNS
# ==================================================

release_pattern = re.compile(
    r"Checking Release ID:\s*(\d+)\s*\[PM ID:\s*([^\]]+)\]"
)

owner_pattern = re.compile(
    r"👤 Release Owned By:\s*(.*?)\s*\(Tasks may be owned by others\)"
)

date_pattern = re.compile(
    r"📅 Created:\s*(\d{2}-\d{2}-\d{4})\s*\|\s*Resolved:\s*(\d{2}-\d{2}-\d{4})"
)

# ==================================================
# READ FILE
# ==================================================

with open(INPUT_FILE, "r", encoding="utf-8") as f:

    for line in f:

        line = (
            line.replace("<br>", "")
                .replace("&nbsp;", " ")
                .strip()
        )

        if not line:
            continue

        # ==========================================
        # RELEASE
        # ==========================================

        m = release_pattern.search(line)

        if m:
            current_release = m.group(1)
            current_pm = m.group(2)
            continue

        # ==========================================
        # RELEASE OWNER
        # ==========================================

        m = owner_pattern.search(line)

        if m:
            current_release_owner = m.group(1)
            continue

        # ==========================================
        # RELEASE DATES
        # ==========================================

        m = date_pattern.search(line)

        if m:
            current_release_created = m.group(1)
            current_release_resolved = m.group(2)
            continue

        # ==========================================
        # TASK RECORD
        # ==========================================

        if (
            "Added" in line
            and "ID:" in line
            and "Title:" in line
        ):

            try:

                category = re.search(
                    r"\[(.*?)\]\s*Added",
                    line
                ).group(1).strip()

                hours = float(
                    re.search(
                        r"Added\s*([\d.]+)\s*hrs",
                        line
                    ).group(1)
                )

                task_id = re.search(
                    r"ID:\s*(\d+)",
                    line
                ).group(1)

                task_type = re.search(
                    r"Type:\s*(.*?)\s*\|\s*Dept:",
                    line
                ).group(1)

                department = re.search(
                    r"Dept:\s*'(.*?)'",
                    line
                ).group(1)

                task_owner = re.search(
                    r"Owner:\s*(.*?)\s*\|\s*Country:",
                    line
                ).group(1)

                country = re.search(
                    r"Country:\s*(.*?)\s*\|\s*Title:",
                    line
                ).group(1)

                title = re.search(
                    r"Title:\s*(.*?)\s*\|\s*Created:",
                    line
                ).group(1)

                task_created = re.search(
                    r"Created:\s*(\d{2}-\d{2}-\d{4})",
                    line
                ).group(1)

                rows.append({
                    "Release ID": current_release,
                    "PM ID": current_pm,
                    "Release Owner": current_release_owner,
                    "Release Created": current_release_created,
                    "Release Resolved": current_release_resolved,

                    "Category": category,
                    "Hours": hours,

                    "Task ID": task_id,
                    "Type": task_type,

                    "Department": department,
                    "Task Owner": task_owner,

                    "Country": country,
                    "Title": title,

                    "Task Created": task_created
                })

            except Exception as e:

                print("\nFAILED TO PARSE:")
                print(line)
                print(e)

# ==================================================
# DATAFRAME
# ==================================================

detail_df = pd.DataFrame(rows)

if detail_df.empty:
    print("No task records found.")
    raise SystemExit

# ==================================================
# DATE CONVERSIONS
# ==================================================

for col in [
    "Release Created",
    "Release Resolved",
    "Task Created"
]:
    detail_df[col] = pd.to_datetime(
        detail_df[col],
        format="%d-%m-%Y",
        errors="coerce"
    )

# ==================================================
# RELEASE DATE SPLITS
# ==================================================

detail_df["Release Day"] = detail_df["Release Created"].dt.day
detail_df["Release Month"] = detail_df["Release Created"].dt.month
detail_df["Release Year"] = detail_df["Release Created"].dt.year

# ==================================================
# TASK DATE SPLITS
# ==================================================

detail_df["Task Day"] = detail_df["Task Created"].dt.day
detail_df["Task Month"] = detail_df["Task Created"].dt.month
detail_df["Task Year"] = detail_df["Task Created"].dt.year

# ==================================================
# SUMMARY
# ==================================================

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

hour_cols = [
    c for c in summary_df.columns
    if c != "Release ID"
]

summary_df["Total Hours"] = summary_df[
    hour_cols
].sum(axis=1)

# ==================================================
# WRITE OUTPUT
# ==================================================

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

print("\nDone!")
print(f"Records Parsed : {len(detail_df)}")
print(f"Output File    : {OUTPUT_FILE}")