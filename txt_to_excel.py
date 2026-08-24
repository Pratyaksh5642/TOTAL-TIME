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
    "added_time_log_Final_Corrected.txt"
)

OUTPUT_FILE = os.path.join(
    SCRIPT_DIR,
    "Final_Corrected_hours.xlsx"
)


# ====================================================
# REGEX PATTERNS
# ====================================================

prefix_pattern = re.compile(
    r"^\[Rel\s+\d+\]\s*"
)

release_pattern = re.compile(
    r"Checking Release ID:\s*(\d+)\s*"
    r"\[PM ID:\s*([^\]]*)\]"
)

root_item_type_pattern = re.compile(
    r"Root Item Type:\s*(.+?)\s*$",
    re.IGNORECASE
)

status_pattern = re.compile(
    r"(?:Status|Resolution)\s+is\s+'([^']+)'",
    re.IGNORECASE
)

skipped_status_pattern = re.compile(
    r"SKIPPING:\s*Resolution\s+is\s+'([^']+)'",
    re.IGNORECASE
)

release_owner_pattern = re.compile(
    r"Release Owned By:\s*(.+?)\s*"
    r"\(Tasks may be owned by others\)",
    re.IGNORECASE
)

release_date_pattern = re.compile(
    r"Created:\s*(\d{2}-\d{2}-\d{4})\s*\|\s*"
    r"Resolved:\s*(No Date|\d{2}-\d{2}-\d{4})",
    re.IGNORECASE
)

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

task_rows = []
failed_task_lines = []

current_release = ""
current_pm = ""
current_root_item_type = ""
current_release_status = ""
current_release_owner = ""
current_release_created = ""
current_release_resolved = ""


# ====================================================
# EXCEL-SAFE TEXT FUNCTION
# ====================================================

def make_excel_safe(value):
    """
    Prevent imported text from being interpreted as an
    Excel formula and remove invalid control characters.
    """

    if pd.isna(value):
        return value

    if not isinstance(value, str):
        return value

    # Remove control characters that Excel cannot store.
    value = re.sub(
        r"[\x00-\x08\x0B\x0C\x0E-\x1F]",
        "",
        value
    )

    # Excel interprets values beginning with these
    # characters as formulas.
    if value.startswith(("=", "+", "-", "@")):
        value = "'" + value

    return value


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
# DELETE OLD OUTPUT FILE
# ====================================================

if os.path.exists(OUTPUT_FILE):

    try:
        os.remove(OUTPUT_FILE)

    except PermissionError:

        print()
        print("Cannot replace the Excel file.")
        print("Close checking_hours.xlsx in Excel and run again.")
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

        # Convert HTML values:
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

        # Remove prefix such as:
        # [Rel 7276394]
        line = prefix_pattern.sub("", line).strip()

        if not line:
            continue

        # --------------------------------------------
        # NEW RELEASE
        # --------------------------------------------

        match = release_pattern.search(line)

        if match:

            current_release = match.group(1).strip()
            current_pm = match.group(2).strip()

            current_root_item_type = ""
            current_release_status = ""
            current_release_owner = ""
            current_release_created = ""
            current_release_resolved = ""

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
        # RELEASE CREATED / RESOLVED
        # --------------------------------------------

        if "Added" not in line:

            match = release_date_pattern.search(line)

            if match:
                current_release_created = match.group(1).strip()
                current_release_resolved = match.group(2).strip()
                continue

        # --------------------------------------------
        # TASK / REVIEW RECORD
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


# ====================================================
# CREATE DETAILED DATAFRAME
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

detail_df = pd.DataFrame(
    task_rows,
    columns=detail_columns
)


if detail_df.empty:

    print()
    print("No task records were found.")
    print("Check the format of added_time_log.txt.")

    if failed_task_lines:
        print()
        print("First unparsed task line:")
        print(failed_task_lines[0]["Unparsed Line"])

    raise SystemExit


# ====================================================
# FORMAT TASK CREATED
# ====================================================

task_created_date = pd.to_datetime(
    detail_df["Task Created"],
    format="%d-%m-%Y",
    errors="coerce"
)

detail_df["Task Created"] = (
    task_created_date.dt.strftime("%d-%m-%Y")
)

detail_df["Task Created"] = (
    detail_df["Task Created"].fillna("No Date")
)


# ====================================================
# MONTH AND YEAR FROM TASK RESOLVED
# ====================================================

task_resolved_original = (
    detail_df["Task Resolved"]
    .astype("string")
    .str.strip()
)

no_task_resolution = (
    task_resolved_original.isna()
    | task_resolved_original.eq("")
    | task_resolved_original.str.casefold().eq("no date")
)

task_resolved_date = pd.to_datetime(
    task_resolved_original.mask(no_task_resolution),
    format="%d-%m-%Y",
    errors="coerce"
)

# Blank month when Task Resolved is No Date
detail_df["Month"] = (
    task_resolved_date.dt.month.astype("Int64")
)

# Year 2027 when Task Resolved is No Date
detail_df["Year"] = (
    task_resolved_date.dt.year
    .where(~no_task_resolution, 2027)
    .astype("Int64")
)

detail_df["Task Resolved"] = (
    task_resolved_date.dt.strftime("%d-%m-%Y")
)

detail_df.loc[
    no_task_resolution,
    "Task Resolved"
] = "No Date"


# ====================================================
# FORMAT RELEASE CREATED
# ====================================================

release_created_date = pd.to_datetime(
    detail_df["Release Created"],
    format="%d-%m-%Y",
    errors="coerce"
)

detail_df["Release Created"] = (
    release_created_date.dt.strftime("%d-%m-%Y")
)

detail_df["Release Created"] = (
    detail_df["Release Created"].fillna("No Date")
)


# ====================================================
# FORMAT RELEASE RESOLVED
# ====================================================

release_resolved_original = (
    detail_df["Release Resolved"]
    .astype("string")
    .str.strip()
)

release_resolved_missing = (
    release_resolved_original.isna()
    | release_resolved_original.eq("")
    | release_resolved_original.str.casefold().eq("no date")
)

release_resolved_date = pd.to_datetime(
    release_resolved_original.mask(
        release_resolved_missing
    ),
    format="%d-%m-%Y",
    errors="coerce"
)

detail_df["Release Resolved"] = (
    release_resolved_date.dt.strftime("%d-%m-%Y")
)

detail_df.loc[
    release_resolved_missing,
    "Release Resolved"
] = "No Date"


# ====================================================
# KEEP IDs AS TEXT
# ====================================================

detail_df["Release ID"] = (
    detail_df["Release ID"].astype("string")
)

detail_df["Task ID"] = (
    detail_df["Task ID"].astype("string")
)

detail_df["PM ID"] = (
    detail_df["PM ID"]
    .fillna("")
    .astype("string")
)


# ====================================================
# FINAL COLUMN ORDER
# ====================================================

detail_df = detail_df[
    [
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
        "Task Resolved",
        "Month",
        "Year"
    ]
]


# ====================================================
# MAKE ALL TEXT SAFE FOR EXCEL
# ====================================================

text_columns = detail_df.select_dtypes(
    include=[
        "object",
        "string"
    ]
).columns

for column in text_columns:

    detail_df[column] = detail_df[column].map(
        make_excel_safe
    )


# ====================================================
# WRITE ONLY DETAILED_DATA
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

    worksheet = writer.sheets["Detailed_Data"]

    # Force imported text cells to remain plain text.
    for column_number, column_name in enumerate(
        detail_df.columns,
        start=1
    ):

        if column_name in text_columns:

            for row_number in range(
                2,
                len(detail_df) + 2
            ):

                worksheet.cell(
                    row=row_number,
                    column=column_number
                ).number_format = "@"


# ====================================================
# FINISHED
# ====================================================

print()
print("Excel created successfully")
print(f"Task rows exported: {len(detail_df)}")
print(f"Unparsed task lines: {len(failed_task_lines)}")
print(f"Output file: {OUTPUT_FILE}")