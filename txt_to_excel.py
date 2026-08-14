import os
import re
import html
import pandas as pd

# ====================================================
# FILE PATHS
# ====================================================

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

INPUT_FILE = os.path.join(
    SCRIPT_DIR,
    "open_added_time_log_threading.txt"
)

OUTPUT_FILE = os.path.join(
    SCRIPT_DIR,
    "open_Dailmer_tasks.xlsx"
)

# ====================================================
# REGEX PATTERNS
# ====================================================

# Removes prefix such as:
# [Rel 7121903]
prefix_pattern = re.compile(
    r"^\[Rel\s+\d+\]\s*"
)

# Example:
# Checking Release ID: 7121903 [PM ID: BM-00105824]...
release_pattern = re.compile(
    r"Checking Release ID:\s*(\d+)\s*"
    r"\[PM ID:\s*([^\]]+)\]"
)

# Example:
# Status is 'Unresolved (New / In Progress)'
# Resolution is 'Solved'
status_pattern = re.compile(
    r"(?:Status|Resolution)\s+is\s+'([^']+)'",
    re.IGNORECASE
)

# Example:
# SKIPPING: Resolution is 'Invalid'
skipped_status_pattern = re.compile(
    r"SKIPPING:\s*Resolution\s+is\s+'([^']+)'",
    re.IGNORECASE
)

# Example:
# Release Owned By: Glitz Ute (...) (Tasks may be owned by others)
release_owner_pattern = re.compile(
    r"Release Owned By:\s*(.+?)\s*"
    r"\(Tasks may be owned by others\)"
)

# Supports:
# Created: 15-06-2026 | Resolved: No Date
# Created: 15-06-2026 | Resolved: 20-06-2026
release_date_pattern = re.compile(
    r"Created:\s*(\d{2}-\d{2}-\d{4})\s*\|\s*"
    r"Resolved:\s*(No Date|\d{2}-\d{2}-\d{4})",
    re.IGNORECASE
)

# Example task:
# [DSM DEV ] Added 1.00 hrs | ID: ... | ... |
# Created: 18-05-2026 | Resolved: 21-05-2026
task_pattern = re.compile(
    r"\[(.*?)\]\s*"
    r"Added\s*([\d.]+)\s*hrs\s*\|\s*"
    r"ID:\s*(\d+)\s*\|\s*"
    r"Type:\s*(.*?)\s*\|\s*"
    r"Dept:\s*'(.*?)'\s*\|\s*"
    r"Owner:\s*(.*?)\s*\|\s*"
    r"Country:\s*(.*?)\s*\|\s*"
    r"Title:\s*(.*?)\s*\|\s*"
    r"Created:\s*(\d{2}-\d{2}-\d{4})\s*\|\s*"
    r"Resolved:\s*(No Date|\d{2}-\d{2}-\d{4})",
    re.IGNORECASE
)

# ====================================================
# STORAGE
# ====================================================

rows = []
release_rows = []

current_release = ""
current_pm = ""
current_release_status = ""
current_release_owner = ""
current_release_created = ""
current_release_resolved = ""

# Prevents the same release from being added more than once
saved_release_ids = set()

# ====================================================
# HELPER FUNCTION
# ====================================================

def save_current_release():
    """
    Saves one row per release in Release_Info.
    This includes unresolved, invalid, and releases
    without valid task hours.
    """

    if not current_release:
        return

    if current_release in saved_release_ids:
        return

    release_rows.append({
        "Release ID": current_release,
        "PM ID": current_pm,
        "Release Status": current_release_status,
        "Release Owner": current_release_owner,
        "Release Created": current_release_created,
        "Release Resolved": current_release_resolved
    })

    saved_release_ids.add(current_release)


# ====================================================
# CHECK INPUT FILE
# ====================================================

if not os.path.exists(INPUT_FILE):
    print("Input file was not found:")
    print(INPUT_FILE)
    raise SystemExit

# ====================================================
# READ AND PARSE TEXT FILE
# ====================================================

with open(
    INPUT_FILE,
    "r",
    encoding="utf-8"
) as file:

    for raw_line in file:

        # Convert HTML entities:
        # &nbsp; becomes a space
        # &amp; becomes &
        line = html.unescape(raw_line)

        line = (
            line.replace("<br>", "")
            .replace("&lt;br&gt;", "")
            .replace("\u00a0", " ")
            .strip()
        )

        if not line:
            continue

        # Remove [Rel 7121903] prefix
        line = prefix_pattern.sub("", line).strip()

        if not line:
            continue

        # --------------------------------------------
        # NEW RELEASE
        # --------------------------------------------

        match = release_pattern.search(line)

        if match:

            # Save the previous release before starting
            # the next release
            save_current_release()

            current_release = match.group(1).strip()
            current_pm = match.group(2).strip()

            # Reset release-specific information
            current_release_status = ""
            current_release_owner = ""
            current_release_created = ""
            current_release_resolved = ""

            continue

        # --------------------------------------------
        # SKIPPED / INVALID RELEASE STATUS
        # --------------------------------------------

        match = skipped_status_pattern.search(line)

        if match:
            current_release_status = match.group(1).strip()
            continue

        # --------------------------------------------
        # RELEASE STATUS
        # --------------------------------------------

        match = status_pattern.search(line)

        if match:
            current_release_status = match.group(1).strip()
            continue

        # --------------------------------------------
        # RELEASE OWNER
        # --------------------------------------------

        match = release_owner_pattern.search(line)

        if match:
            current_release_owner = match.group(1).strip()
            continue

        # --------------------------------------------
        # RELEASE CREATED / RESOLVED DATES
        # --------------------------------------------

        # Only use this pattern for the release-level
        # date line, not task lines
        if "Added" not in line:

            match = release_date_pattern.search(line)

            if match:
                current_release_created = match.group(1).strip()
                current_release_resolved = match.group(2).strip()
                continue

        # --------------------------------------------
        # TASK / REVIEW WORK ITEM
        # --------------------------------------------

        match = task_pattern.search(line)

        if match:

            rows.append({
                "Release ID": current_release,
                "PM ID": current_pm,
                "Release Status": current_release_status,
                "Release Owner": current_release_owner,
                "Release Created": current_release_created,
                "Release Resolved": current_release_resolved,

                "Category": match.group(1).strip(),
                "Hours": float(match.group(2)),
                "Task ID": match.group(3).strip(),
                "Type": match.group(4).strip(),
                "Department": match.group(5).strip(),
                "Task Owner": match.group(6).strip(),
                "Country": match.group(7).strip(),
                "Title": match.group(8).strip(),
                "Task Created": match.group(9).strip(),
                "Task Resolved": match.group(10).strip()
            })

# Save the final release in the file
save_current_release()

# ====================================================
# CREATE DATAFRAMES
# ====================================================

detail_df = pd.DataFrame(rows)
release_df = pd.DataFrame(release_rows)

if detail_df.empty:
    print("No task records were found.")
    print("Check whether the text format matches the examples.")
    raise SystemExit

# ====================================================
# REMOVE EXACT DUPLICATE WORK ITEMS
# ====================================================

# If the same Task ID appears twice under the same
# Release ID, only the first occurrence is retained.
#
# The same Task ID can still appear under a different
# Release ID.

rows_before = len(detail_df)

detail_df = detail_df.drop_duplicates(
    subset=[
        "Release ID",
        "Task ID"
    ],
    keep="first"
).reset_index(drop=True)

duplicates_removed = rows_before - len(detail_df)

# ====================================================
# TASK CREATED DATE COLUMNS
# ====================================================

task_created_date = pd.to_datetime(
    detail_df["Task Created"],
    format="%d-%m-%Y",
    errors="coerce"
)

# These values come only from Task Created
detail_df["Created_Day"] = (
    task_created_date.dt.day.astype("Int64")
)

detail_df["Created_Month"] = (
    task_created_date.dt.month.astype("Int64")
)

detail_df["Created_Year"] = (
    task_created_date.dt.year.astype("Int64")
)

# Keep Task Created as DD-MM-YYYY
# This prevents 00:00:00 from appearing
detail_df["Task Created"] = (
    task_created_date.dt.strftime("%d-%m-%Y")
)

# ====================================================
# FORMAT TASK RESOLVED DATE
# ====================================================

# Preserve "No Date".
# Format real dates as DD-MM-YYYY.

task_resolved_original = detail_df[
    "Task Resolved"
].copy()

task_resolved_date = pd.to_datetime(
    task_resolved_original.replace(
        "No Date",
        pd.NA
    ),
    format="%d-%m-%Y",
    errors="coerce"
)

detail_df["Task Resolved"] = (
    task_resolved_date.dt.strftime("%d-%m-%Y")
)

detail_df["Task Resolved"] = (
    detail_df["Task Resolved"]
    .fillna("No Date")
)

# ====================================================
# SUMMARY BY RELEASE AND CATEGORY
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

# Remove the pivot table column label
summary_df.columns.name = None

hour_columns = [
    column
    for column in summary_df.columns
    if column != "Release ID"
]

summary_df["Total Hours"] = (
    summary_df[hour_columns].sum(axis=1)
)

# ====================================================
# SUMMARY BY RELEASE, COUNTRY AND CATEGORY
# ====================================================

country_summary_df = (
    detail_df
    .pivot_table(
        index=[
            "Release ID",
            "Country"
        ],
        columns="Category",
        values="Hours",
        aggfunc="sum",
        fill_value=0
    )
    .reset_index()
)

country_summary_df.columns.name = None

country_hour_columns = [
    column
    for column in country_summary_df.columns
    if column not in [
        "Release ID",
        "Country"
    ]
]

country_summary_df["Total Hours"] = (
    country_summary_df[
        country_hour_columns
    ].sum(axis=1)
)

# ====================================================
# WRITE EXCEL
# ====================================================

with pd.ExcelWriter(
    OUTPUT_FILE,
    engine="openpyxl"
) as writer:

    # One row per task/review
    detail_df.to_excel(
        writer,
        sheet_name="Detailed_Data",
        index=False
    )

    # One row per release
    release_df.to_excel(
        writer,
        sheet_name="Release_Info",
        index=False
    )

    # Hours grouped by Release ID and Category
    summary_df.to_excel(
        writer,
        sheet_name="Release_Summary",
        index=False
    )

    # Hours grouped by Release ID, Country and Category
    country_summary_df.to_excel(
        writer,
        sheet_name="Country_Summary",
        index=False
    )

# ====================================================
# FINISHED
# ====================================================

print()
print("Excel created successfully")
print(f"Task rows exported: {len(detail_df)}")
print(f"Duplicate task rows removed: {duplicates_removed}")
print(f"Releases found: {len(release_df)}")
print(f"Output file: {OUTPUT_FILE}")