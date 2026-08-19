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
    "added_time_log_2023_2024.txt"
)

OUTPUT_FILE = os.path.join(
    SCRIPT_DIR,
    "2023_2024_task_hours.xlsx"
)


# ====================================================
# REGEX PATTERNS
# ====================================================

# Removes:
# [Rel 7276394]
prefix_pattern = re.compile(
    r"^\[Rel\s+\d+\]\s*"
)

# Supports both:
# [PM ID: BM-00105824]
# [PM ID: ]
#
# The * allows the PM ID to be empty.
release_pattern = re.compile(
    r"Checking Release ID:\s*(\d+)\s*"
    r"\[PM ID:\s*([^\]]*)\]"
)

# Supports:
# Root Item Type: Task
# Root Item Type: Epic
# Root Item Type: Release
root_item_type_pattern = re.compile(
    r"Root Item Type:\s*(.+?)\s*$",
    re.IGNORECASE
)

# Supports:
# Status is 'Unresolved (New / In Progress)'
# Resolution is 'Solved'
status_pattern = re.compile(
    r"(?:Status|Resolution)\s+is\s+'([^']+)'",
    re.IGNORECASE
)

# Supports:
# SKIPPING: Resolution is 'Invalid'
skipped_status_pattern = re.compile(
    r"SKIPPING:\s*Resolution\s+is\s+'([^']+)'",
    re.IGNORECASE
)

# Supports:
# Release Owned By: Name (...) (Tasks may be owned by others)
release_owner_pattern = re.compile(
    r"Release Owned By:\s*(.+?)\s*"
    r"\(Tasks may be owned by others\)",
    re.IGNORECASE
)

# Supports:
# Created: 24-07-2026 | Resolved: 24-07-2026
# Created: 24-07-2026 | Resolved: No Date
release_date_pattern = re.compile(
    r"Created:\s*(\d{2}-\d{2}-\d{4})\s*\|\s*"
    r"Resolved:\s*(No Date|\d{2}-\d{2}-\d{4})",
    re.IGNORECASE
)

# Supports task lines such as:
# [NET DEV ] Added 12.50 hrs | ID: 7276394 |
# Type: task |
# Dept: 'var - variant coding and variant handling' |
# Owner: Name (...) |
# Country: India |
# Title: Example |
# Created: 24-07-2026 |
# Resolved: No Date
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

# Supports:
# [India]
# [Germany]
# [Others]
final_country_pattern = re.compile(
    r"^\[(?:🌍\s*)?(.*?)\]\s*$"
)

# Supports:
# NET_DEV: 12.50 hrs
# DCOM_DEV: 80.50 hrs
# DSM_DEV: 48.75 hrs
final_hours_pattern = re.compile(
    r"([A-Za-z0-9_ ]+):\s*"
    r"([\d.]+)\s*hrs",
    re.IGNORECASE
)

# Supports:
# No Valid Hours Logged (0.00 hrs)
no_hours_pattern = re.compile(
    r"No Valid Hours Logged\s*"
    r"\(([\d.]+)\s*hrs\)",
    re.IGNORECASE
)


# ====================================================
# DATA STORAGE
# ====================================================

task_rows = []
release_rows = []
failed_task_lines = []
log_summary_rows = []

current_release = ""
current_pm = ""
current_root_item_type = ""
current_release_status = ""
current_release_owner = ""
current_release_created = ""
current_release_resolved = ""

current_summary_country = ""

release_started = False


# ====================================================
# SAVE CURRENT RELEASE
# ====================================================

def save_current_release():

    if not release_started:
        return

    if not current_release:
        return

    release_rows.append({
        "Release ID": current_release,
        "PM ID": current_pm,
        "Root Item Type": current_root_item_type,
        "Release Status": current_release_status,
        "Release Owner": current_release_owner,
        "Release Created": current_release_created,
        "Release Resolved": current_release_resolved
    })


# ====================================================
# DATE FORMAT FUNCTION
# ====================================================

def format_date_column(series):

    original_values = series.astype("string").str.strip()

    converted_dates = pd.to_datetime(
        original_values.replace(
            {
                "No Date": pd.NA,
                "": pd.NA,
                "nan": pd.NA,
                "None": pd.NA
            }
        ),
        format="%d-%m-%Y",
        errors="coerce"
    )

    formatted_values = converted_dates.dt.strftime(
        "%d-%m-%Y"
    )

    formatted_values = formatted_values.fillna("No Date")

    return formatted_values


# ====================================================
# CHECK INPUT FILE
# ====================================================

if not os.path.exists(INPUT_FILE):

    print()
    print("Input file not found:")
    print(INPUT_FILE)
    print()
    print(
        "Put added_time_log.txt in the same folder "
        "as this Python script."
    )

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
        # &nbsp; becomes a normal space
        # &amp; becomes &
        line = html.unescape(raw_line)

        line = (
            line
            .replace("<br>", "")
            .replace("&lt;br&gt;", "")
            .replace("\u00a0", " ")
            .strip()
        )

        if not line:
            continue

        # Remove prefix:
        # [Rel 7276394]
        line = prefix_pattern.sub("", line).strip()

        if not line:
            continue

        # --------------------------------------------
        # NEW RELEASE
        # --------------------------------------------

        match = release_pattern.search(line)

        if match:

            if release_started:
                save_current_release()

            current_release = match.group(1).strip()
            current_pm = match.group(2).strip()

            current_root_item_type = ""
            current_release_status = ""
            current_release_owner = ""
            current_release_created = ""
            current_release_resolved = ""
            current_summary_country = ""

            release_started = True

            continue

        # Ignore information before first release
        if not release_started:
            continue

        # --------------------------------------------
        # ROOT ITEM TYPE
        # --------------------------------------------

        match = root_item_type_pattern.search(line)

        if match:
            current_root_item_type = match.group(1).strip()
            continue

        # --------------------------------------------
        # INVALID / SKIPPED RELEASE
        # --------------------------------------------

        match = skipped_status_pattern.search(line)

        if match:
            current_release_status = match.group(1).strip()
            continue

        # --------------------------------------------
        # RELEASE STATUS / RESOLUTION
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

        # Do not treat a task date as a release date.
        if "Added" not in line:

            match = release_date_pattern.search(line)

            if match:
                current_release_created = match.group(1).strip()
                current_release_resolved = match.group(2).strip()
                continue

        # --------------------------------------------
        # TASK OR REVIEW RECORD
        # --------------------------------------------

        if (
            "Added" in line
            and "ID:" in line
            and "Title:" in line
        ):

            match = task_pattern.search(line)

            if match:

                task_rows.append({
                    "Release ID": current_release,
                    "PM ID": current_pm,
                    "Root Item Type": current_root_item_type,
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

            else:

                failed_task_lines.append({
                    "Release ID": current_release,
                    "Unparsed Line": line
                })

            continue

        # --------------------------------------------
        # START OF FINAL HOURS SECTION
        # --------------------------------------------

        if "Final Hours for Release" in line:
            current_summary_country = ""
            continue

        # --------------------------------------------
        # FINAL HOURS COUNTRY
        # --------------------------------------------

        match = final_country_pattern.search(line)

        if match:

            possible_country = match.group(1).strip()

            # Do not treat task category brackets as
            # summary countries.
            if possible_country in [
                "India",
                "Germany",
                "Others"
            ]:
                current_summary_country = possible_country

            continue

        # --------------------------------------------
        # NO VALID HOURS
        # --------------------------------------------

        match = no_hours_pattern.search(line)

        if match:

            log_summary_rows.append({
                "Release ID": current_release,
                "Country": current_summary_country,
                "Category": "No Valid Hours",
                "Hours": float(match.group(1))
            })

            continue

        # --------------------------------------------
        # FINAL HOURS CATEGORY
        # --------------------------------------------

        match = final_hours_pattern.search(line)

        if match and current_summary_country:

            category = match.group(1).strip()
            hours = float(match.group(2))

            log_summary_rows.append({
                "Release ID": current_release,
                "Country": current_summary_country,
                "Category": category,
                "Hours": hours
            })

            continue


# Save final release in the text file
if release_started:
    save_current_release()


# ====================================================
# CREATE DATAFRAMES
# ====================================================

detail_columns = [
    "Release ID",
    "PM ID",
    "Root Item Type",
    "Release Status",
    "Release Owner",
    "Release Created",
    "Release Resolved",
    "Category",
    "Hours",
    "Task ID",
    "Type",
    "Department",
    "Task Owner",
    "Country",
    "Title",
    "Task Created",
    "Task Resolved"
]

release_columns = [
    "Release ID",
    "PM ID",
    "Root Item Type",
    "Release Status",
    "Release Owner",
    "Release Created",
    "Release Resolved"
]

failed_columns = [
    "Release ID",
    "Unparsed Line"
]

log_summary_columns = [
    "Release ID",
    "Country",
    "Category",
    "Hours"
]

detail_df = pd.DataFrame(
    task_rows,
    columns=detail_columns
)

release_df = pd.DataFrame(
    release_rows,
    columns=release_columns
)

failed_df = pd.DataFrame(
    failed_task_lines,
    columns=failed_columns
)

log_summary_df = pd.DataFrame(
    log_summary_rows,
    columns=log_summary_columns
)


# ====================================================
# FORMAT RELEASE INFO
# ====================================================

if not release_df.empty:

    release_df["Release Created"] = format_date_column(
        release_df["Release Created"]
    )

    release_df["Release Resolved"] = format_date_column(
        release_df["Release Resolved"]
    )

    # Release ID remains text so Excel does not add .0
    release_df["Release ID"] = release_df[
        "Release ID"
    ].astype("string")

    release_df["PM ID"] = release_df[
        "PM ID"
    ].fillna("").astype("string")


# ====================================================
# FORMAT DETAILED TASK DATA
# ====================================================

if not detail_df.empty:

    # Convert Task Created for day/month/year columns
    task_created_date = pd.to_datetime(
        detail_df["Task Created"],
        format="%d-%m-%Y",
        errors="coerce"
    )

    # These three columns come only from Task Created
    detail_df["Created_Day"] = (
        task_created_date.dt.day.astype("Int64")
    )

    detail_df["Created_Month"] = (
        task_created_date.dt.month.astype("Int64")
    )

    detail_df["Created_Year"] = (
        task_created_date.dt.year.astype("Int64")
    )

    # Task Created is displayed as DD-MM-YYYY
    # without 00:00:00.
    detail_df["Task Created"] = (
        task_created_date.dt.strftime("%d-%m-%Y")
    )

    detail_df["Task Created"] = (
        detail_df["Task Created"].fillna("No Date")
    )

    # Format Task Resolved and preserve No Date
    detail_df["Task Resolved"] = format_date_column(
        detail_df["Task Resolved"]
    )

    # Format release dates in task details
    detail_df["Release Created"] = format_date_column(
        detail_df["Release Created"]
    )

    detail_df["Release Resolved"] = format_date_column(
        detail_df["Release Resolved"]
    )

    # Keep IDs as text
    detail_df["Release ID"] = detail_df[
        "Release ID"
    ].astype("string")

    detail_df["Task ID"] = detail_df[
        "Task ID"
    ].astype("string")

    detail_df["PM ID"] = detail_df[
        "PM ID"
    ].fillna("").astype("string")

    # Keep all task rows exactly as parsed.
    # No automatic duplicate removal is performed.


# ====================================================
# RELEASE SUMMARY FROM TASK ROWS
# ====================================================

if not detail_df.empty:

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

    summary_df.columns.name = None

    summary_hour_columns = [
        column
        for column in summary_df.columns
        if column != "Release ID"
    ]

    summary_df["Total Hours"] = (
        summary_df[summary_hour_columns].sum(axis=1)
    )

else:

    summary_df = pd.DataFrame(
        columns=[
            "Release ID",
            "Total Hours"
        ]
    )


# ====================================================
# COUNTRY SUMMARY FROM TASK ROWS
# ====================================================

if not detail_df.empty:

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

else:

    country_summary_df = pd.DataFrame(
        columns=[
            "Release ID",
            "Country",
            "Total Hours"
        ]
    )


# ====================================================
# RELEASES WITH NO TASK HOURS
# ====================================================

if not release_df.empty:

    releases_with_tasks = set(
        detail_df["Release ID"].astype(str)
    ) if not detail_df.empty else set()

    no_task_hours_df = release_df[
        ~release_df["Release ID"]
        .astype(str)
        .isin(releases_with_tasks)
    ].copy()

    no_task_hours_df["Total Hours"] = 0.0

else:

    no_task_hours_df = pd.DataFrame(
        columns=release_columns + ["Total Hours"]
    )


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

    release_df.to_excel(
        writer,
        sheet_name="Release_Info",
        index=False
    )

    summary_df.to_excel(
        writer,
        sheet_name="Release_Summary",
        index=False
    )

    country_summary_df.to_excel(
        writer,
        sheet_name="Country_Summary",
        index=False
    )

    no_task_hours_df.to_excel(
        writer,
        sheet_name="No_Valid_Hours",
        index=False
    )

    if not log_summary_df.empty:

        log_summary_df.to_excel(
            writer,
            sheet_name="Log_Final_Hours",
            index=False
        )

    if not failed_df.empty:

        failed_df.to_excel(
            writer,
            sheet_name="Failed_Lines",
            index=False
        )


# ====================================================
# FINISHED
# ====================================================

print()
print("Excel created successfully")
print(f"Task rows exported: {len(detail_df)}")
print(f"Releases found: {len(release_df)}")
print(
    "Releases with no task hours: "
    f"{len(no_task_hours_df)}"
)
print(f"Unparsed task lines: {len(failed_df)}")
print(f"Output file: {OUTPUT_FILE}")