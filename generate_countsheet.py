"""Generate sea otter survey count sheet template from field log summaries."""

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import os
import re
import sys

YEAR = 2024
INPUT_DIR = os.path.join(os.path.dirname(__file__), "inputs")
SITES_FILE = os.path.join(INPUT_DIR, "SITES.xlsx")
LOG_FILE = os.path.join(INPUT_DIR, f"{YEAR}_LOGSummary.xlsx")
OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "outputs")
os.makedirs(OUTPUT_DIR, exist_ok=True)
OUTPUT_FILE = os.path.join(OUTPUT_DIR, f"COUNTSHEET_TEMPLATE_{YEAR}.xlsx")

# --- 1. Input Validation ---
for f, label in [(SITES_FILE, "SITES.xlsx"), (LOG_FILE, f"{YEAR}_LOGSummary.xlsx")]:
    if not os.path.exists(f):
        print(f"Error: Required file not found: {label}")
        sys.exit(1)

# --- 2. Data Loading ---
sites = pd.read_excel(SITES_FILE)
log = pd.read_excel(LOG_FILE)

print(f"SITES columns: {list(sites.columns)}")
print(f"LOGSummary columns: {list(log.columns)}")
print(f"SITES rows: {len(sites)}")
print(f"LOGSummary rows: {len(log)}")

# Normalize SUBSITE for matching
sites["_subsite_key"] = sites["SUBSITE"].astype(str).str.strip().str.upper()
log["_subsite_key"] = log["SUBSITE"].astype(str).str.strip().str.upper()

# --- 3A. Handle Duplicate Surveys ---
# Remove "DO NOT USE" rows
log_clean = log[~log["SUBSITE"].astype(str).str.contains("DO NOT USE", case=False, na=False)].copy()

# Flag duplicates and keep most recent
dup_subsites = log_clean[log_clean.duplicated(subset="_subsite_key", keep=False)]["_subsite_key"].unique()
needs_review_dupes = set(dup_subsites)

# Keep most recent DATE for each subsite
log_clean["DATE"] = pd.to_datetime(log_clean["DATE"], errors="coerce")
log_clean = log_clean.sort_values("DATE", ascending=False).drop_duplicates(subset="_subsite_key", keep="first")

# Build set of log subsites
log_subsites = set(log_clean["_subsite_key"])
sites_subsites = set(sites["_subsite_key"])

# --- 3B-D: Build output rows ---
rows = []

# Process all SITES entries
for _, site_row in sites.iterrows():
    key = site_row["_subsite_key"]
    log_match = log_clean[log_clean["_subsite_key"] == key]
    has_log = len(log_match) > 0
    log_row = log_match.iloc[0] if has_log else None

    # Survey status
    if has_log and pd.notna(log_row["DATE"]):
        survey = "OTTER"
    elif has_log and pd.isna(log_row["DATE"]):
        survey = "MISSED"
    else:
        survey = "OUTSIDE"

    # COUNTTYPE
    counttype = ""
    if survey == "OTTER":
        if has_log and pd.notna(log_row.get("COUNT")) and str(log_row.get("COUNT", "")).strip() != "":
            counttype = 4
        elif has_log and pd.notna(log_row.get("PASS")) and str(log_row.get("PASS", "")).strip() != "":
            counttype = 3

    # PHOTO
    photo = ""
    if has_log and pd.notna(log_row.get("PASS")) and str(log_row.get("PASS", "")).strip() != "":
        photo = "Y"

    # FLAGS
    flags = ""
    if key in needs_review_dupes:
        flags = "NEEDS_REVIEW"
    if has_log:
        site_mml = str(site_row.get("MML_ID", "")).strip()
        log_mml = str(log_row.get("MML_ID", "")).strip()
        # Compare only numeric prefix (e.g., "248A" -> "248")
        site_mml_num = re.match(r"(\d+)", site_mml)
        log_mml_num = re.match(r"(\d+)", log_mml)
        site_mml_num = site_mml_num.group(1) if site_mml_num else site_mml
        log_mml_num = log_mml_num.group(1) if log_mml_num else log_mml
        if site_mml_num and log_mml_num and site_mml_num != log_mml_num:
            flags = "NEEDS_REVIEW"

    def _val(series_row, col):
        """Get value from row, return blank string if null."""
        v = series_row.get(col)
        if pd.isna(v):
            return ""
        return v

    row = {
        "FLAGS": flags,
        "SUBSITE": log_row["SUBSITE"] if has_log else site_row["SUBSITE"],
        "SUBSITE_ID": _val(site_row, "SUBSITE_ID"),
        "PARENTSITE": _val(site_row, "PARENTSITE"),
        "PARENTSITE_ID": _val(site_row, "PARENTSITE_ID"),
        "MML_ID": _val(site_row, "MML_ID"),
        "REGION": _val(log_row, "REGION") if (has_log and survey in ("OTTER", "MISSED")) else _val(site_row, "REGION"),
        "REGNO": _val(log_row, "REGNO") if (has_log and survey in ("OTTER", "MISSED")) else _val(site_row, "REGNO"),
        "RCA": _val(log_row, "RCA") if (has_log and survey in ("OTTER", "MISSED")) else _val(site_row, "RCA"),
        "ROOK": _val(log_row, "ROOK") if (has_log and survey in ("OTTER", "MISSED")) else _val(site_row, "ROOK"),
        "LAT": _val(site_row, "LAT"),
        "LON": _val(site_row, "LON"),
        "PRIORITY": _val(log_row, "Priority") if has_log else "",
        "DATE": _val(log_row, "DATE") if has_log else "",
        "SURVEY": survey,
        "COUNTTYPE": counttype,
        "TIME": _val(log_row, "TIME") if has_log else "",
        "PHOTO": photo,
        "LOG_COUNT": _val(log_row, "COUNT") if has_log else "",
        "ADD": _val(log_row, "ADD") if has_log else "",
        "FRAME": "",
        "BULL": "",
        "SAM": "",
        "FEM": "",
        "JUV": "",
        "PUP": "",
        "PUP_DEAD": "",
        "NP_DEAD": "",
        "NP_TOTAL": "",
        "ALL_COUNT": "",
        "COUNTER_NOTES": "",
        "DISTURBANCE": _val(log_row, "DISTURBANCE") if has_log else "",
        "BRANDS": "",
        "COUNTER": "",
        "SURVEY NOTES": _val(log_row, "PASS DESCRIPTION") if has_log else "",
    }
    rows.append(row)

# Check for NEW SITE entries (in log but not in sites)
new_site_keys = log_subsites - sites_subsites
# Remove any NaN/empty keys
new_site_keys = {k for k in new_site_keys if isinstance(k, str) and k.strip() and k != "NAN"}
print(f"New site keys ({len(new_site_keys)}): {new_site_keys}")
for key in new_site_keys:
    log_row = log_clean[log_clean["_subsite_key"] == key].iloc[0]
    survey = "OTTER" if pd.notna(log_row["DATE"]) else "MISSED"

    counttype = ""
    if survey == "OTTER":
        if pd.notna(log_row.get("COUNT")) and str(log_row.get("COUNT", "")).strip() != "":
            counttype = 4
        elif pd.notna(log_row.get("PASS")) and str(log_row.get("PASS", "")).strip() != "":
            counttype = 3

    photo = ""
    if pd.notna(log_row.get("PASS")) and str(log_row.get("PASS", "")).strip() != "":
        photo = "Y"

    row = {
        "FLAGS": "NEW SITE",
        "SUBSITE": log_row["SUBSITE"],
        "SUBSITE_ID": "",
        "PARENTSITE": "",
        "PARENTSITE_ID": "",
        "MML_ID": _val(log_row, "MML_ID"),
        "REGION": _val(log_row, "REGION"),
        "REGNO": _val(log_row, "REGNO"),
        "RCA": _val(log_row, "RCA"),
        "ROOK": _val(log_row, "ROOK"),
        "LAT": "",
        "LON": "",
        "PRIORITY": _val(log_row, "Priority"),
        "DATE": _val(log_row, "DATE"),
        "SURVEY": survey,
        "COUNTTYPE": counttype,
        "TIME": _val(log_row, "TIME"),
        "PHOTO": photo,
        "LOG_COUNT": _val(log_row, "COUNT"),
        "ADD": _val(log_row, "ADD"),
        "FRAME": "",
        "BULL": "",
        "SAM": "",
        "FEM": "",
        "JUV": "",
        "PUP": "",
        "PUP_DEAD": "",
        "NP_DEAD": "",
        "NP_TOTAL": "",
        "ALL_COUNT": "",
        "COUNTER_NOTES": "",
        "DISTURBANCE": _val(log_row, "DISTURBANCE"),
        "BRANDS": "",
        "COUNTER": "",
        "SURVEY NOTES": _val(log_row, "PASS DESCRIPTION"),
    }
    rows.append(row)

# --- 5. Build DataFrame ---
df = pd.DataFrame(rows)

# Sort by SURVEY (OTTER → MISSED → OUTSIDE), then DATE (earliest → latest), then SUBSITE (A→Z)
survey_order = {"OTTER": 0, "MISSED": 1, "OUTSIDE": 2}
df["_sort_survey"] = df["SURVEY"].map(survey_order)
df["_sort_date"] = pd.to_datetime(df["DATE"], errors="coerce")
df["_sort_subsite"] = df["SUBSITE"].astype(str).str.upper()
df = df.sort_values(["_sort_survey", "_sort_date", "_sort_subsite"], na_position="last").drop(columns=["_sort_survey", "_sort_date", "_sort_subsite"])

# --- 6. Output Generation ---
df.to_excel(OUTPUT_FILE, index=False, engine="openpyxl")

# Apply formatting
wb = load_workbook(OUTPUT_FILE)
ws = wb.active

# Bold header
bold_font = Font(bold=True)
for cell in ws[1]:
    cell.font = bold_font

# Freeze top row
ws.freeze_panes = "A2"

# Auto-fit column widths
for col_idx, col_cells in enumerate(ws.columns, 1):
    max_len = 0
    for cell in col_cells:
        val = str(cell.value) if cell.value else ""
        max_len = max(max_len, len(val))
    ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 40)

# Highlight flagged rows in yellow
yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
flags_col_idx = 1  # FLAGS is column A
for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
    if row[0].value and str(row[0].value).strip():
        for cell in row:
            cell.fill = yellow_fill

# Format DATE column as m/dd display
date_col_idx = list(df.columns).index("DATE") + 1  # 1-based
for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=date_col_idx, max_col=date_col_idx):
    for cell in row:
        if cell.value:
            cell.number_format = "m/dd"

wb.save(OUTPUT_FILE)

# --- 7. Summary Report ---
total = len(df)
otter_count = len(df[df["SURVEY"] == "OTTER"])
missed_count = len(df[df["SURVEY"] == "MISSED"])
outside_count = len(df[df["SURVEY"] == "OUTSIDE"])
new_site_count = len(df[df["FLAGS"] == "NEW SITE"])
review_count = len(df[df["FLAGS"] == "NEEDS_REVIEW"])

print(f"""
Count Sheet Generation Summary for {YEAR}
==========================================
Total sites in template: {total}
  - OTTER (surveyed): {otter_count}
  - MISSED (planned but not surveyed): {missed_count}
  - OUTSIDE (not planned): {outside_count}

Flags raised:
  - NEW SITE: {new_site_count} sites
  - NEEDS_REVIEW: {review_count} sites

Column count: {len(df.columns)} (expected 35)

Output file: {OUTPUT_FILE}
""")

# Quality checks
assert df["SURVEY"].isin(["OTTER", "MISSED", "OUTSIDE"]).all(), "Invalid SURVEY values found"
assert df["COUNTTYPE"].isin(["", 3, 4]).all(), "Invalid COUNTTYPE values found"
assert len(df.columns) == 35, f"Expected 35 columns, got {len(df.columns)}"

# Check no duplicate subsites
dup_check = df[df["SUBSITE"].astype(str).str.strip().str.upper().duplicated(keep=False)]
if len(dup_check) > 0:
    print(f"WARNING: {len(dup_check)} duplicate SUBSITE entries found")
    print(dup_check[["FLAGS", "SUBSITE", "SURVEY"]].to_string())

print("Count sheet generation complete.")
