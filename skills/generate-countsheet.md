---
name: generate-countsheet
description: Generate sea otter survey count sheet templates from field log summaries. Reads SITES.xlsx master list and year-specific LOGSummary files, applies survey status logic, validates data integrity, and produces formatted COUNTSHEET_TEMPLATE with automated quality checks. Use when asked to create, generate, or build a count sheet from survey data.
---

# Sea Otter Survey Count Sheet Generator

## Overview

Generate count sheet templates for sea otter aerial surveys by merging master site data with field log summaries. This skill automates the manual data transformation process, applying business rules for survey status classification, data validation, and template formatting.

## Workflow

### 1. Input Validation

**Required files in the working directory:**
- `SITES.xlsx` - Master list of all possible survey sites (502 sites)
- `{YEAR}_LOGSummary.xlsx` - Field log for the survey year (e.g., `2024_LOGSummary.xlsx`)

**Pre-flight checks:**
- Verify both files exist
- Check for expected columns in each file
- Validate file readability
- If files missing, provide clear error message with expected file names

### 2. Data Loading

**Load SITES.xlsx:**
- Expected columns: SUBSITE, SUBSITE_ID, PARENTSITE, PARENTSITE_ID, MML_ID, REGION, REGNO, RCA, ROOK, LAT, LON
- Use SUBSITE as primary key
- Preserve all 502 sites

**Load LOGSummary:**
- Expected columns: DATE, MML_ID, SUBSITE, PARENTSITE, TIME, COUNT, PASS, START FRAME, END FRAME, PASS DESCRIPTION, ADD, DISTURBANCE, Priority, REGION, REGNO, RCA, ROOK
- Use SUBSITE as primary key
- Handle duplicates per rules below

### 3. Data Processing Rules

#### **Step 3A: Handle Duplicate Surveys**

```python
# Remove rows marked as "DO NOT USE"
# If same SUBSITE appears multiple times (excluding DO NOT USE):
#   - Keep the most recent DATE
#   - Flag as "NEEDS_REVIEW" if multiple valid entries exist
```

#### **Step 3B: Determine Survey Status**

For each site, calculate SURVEY column:

```
IF SUBSITE contains "DO NOT USE":
    → SKIP entirely (don't include in output)

ELSE IF SUBSITE in LOGSUMMARY AND DATE is not null/empty:
    → SURVEY = "OTTER" (site was successfully surveyed)
    
ELSE IF SUBSITE in LOGSUMMARY AND DATE is null/empty:
    → SURVEY = "MISSED" (site was planned but not surveyed)
    
ELSE IF SUBSITE in SITES but NOT in LOGSUMMARY:
    → SURVEY = "OUTSIDE" (site not planned for this survey period)
```

#### **Step 3C: Calculate COUNTTYPE**

```
IF SURVEY = "OTTER":
    IF COUNT column is not null:
        → COUNTTYPE = 4 (visual count from aircraft)
    ELSE IF PASS column is not null:
        → COUNTTYPE = 3 (photographic count)
    ELSE:
        → COUNTTYPE = blank
ELSE:
    → COUNTTYPE = blank
```

#### **Step 3D: Calculate PHOTO**

```
IF PASS column is not null:
    → PHOTO = "Y"
ELSE:
    → PHOTO = blank
```

### 4. Data Validation & Flagging

Create FLAGS column (Column A) with quality checks:

```
FLAGS = "NEW SITE" IF:
    - SUBSITE appears in LOGSUMMARY
    - SUBSITE NOT in SITES.xlsx
    - SUBSITE does NOT contain "DO NOT USE"
    
FLAGS = "NEEDS_REVIEW" IF:
    - MML_ID in LOGSUMMARY ≠ MML_ID in SITES for same SUBSITE
    - Multiple non-"DO NOT USE" entries for same SUBSITE exist
    - SUBSITE can't be matched between files
    - Other data integrity issues
    
FLAGS = blank IF:
    - No issues detected
```

### 5. Column Mapping & Merging

Build output template with these columns in order:

| Column | Source | Logic |
|--------|--------|-------|
| FLAGS | Calculated | Quality check flags |
| SUBSITE | LOGSUMMARY (priority) or SITES | Primary identifier |
| SUBSITE_ID | SITES | From master list |
| PARENTSITE | SITES | From master list |
| PARENTSITE_ID | SITES | From master list |
| MML_ID | SITES | From master list (validate against LOGSUMMARY) |
| REGION | LOGSUMMARY (if surveyed) else SITES | - |
| REGNO | LOGSUMMARY (if surveyed) else SITES | - |
| RCA | LOGSUMMARY (if surveyed) else SITES | - |
| ROOK | LOGSUMMARY (if surveyed) else SITES | - |
| LAT | SITES | From master list |
| LON | SITES | From master list |
| PRIORITY | LOGSUMMARY | From field log (if exists) |
| DATE | LOGSUMMARY | Blank if not surveyed |
| SURVEY | Calculated | "OTTER" / "MISSED" / "OUTSIDE" |
| COUNTTYPE | Calculated | 3, 4, or blank |
| TIME | LOGSUMMARY | Copy if not null |
| PHOTO | Calculated | "Y" or blank |
| LOG_COUNT | LOGSUMMARY "COUNT" | Copy if not null |
| ADD | LOGSUMMARY | Copy if not null |
| FRAME | Manual entry | Always blank in template |
| BULL | Manual entry | Always blank in template |
| SAM | Manual entry | Always blank in template |
| FEM | Manual entry | Always blank in template |
| JUV | Manual entry | Always blank in template |
| PUP | Manual entry | Always blank in template |
| PUP_DEAD | Manual entry | Always blank in template |
| NP_DEAD | Manual entry | Always blank in template |
| NP_TOTAL | Manual entry | Always blank in template |
| ALL_COUNT | Manual entry | Always blank in template |
| COUNTER_NOTES | Manual entry | Always blank in template |
| DISTURBANCE | LOGSUMMARY | Copy if not null |
| BRANDS | Manual entry | Always blank in template |
| COUNTER | Manual entry | Always blank in template |
| SURVEY NOTES | LOGSUMMARY "PASS DESCRIPTION" | Copy if not null |

**Important Notes:**
- For sites with SURVEY="OTTER" or "MISSED": use data from LOGSUMMARY where available
- For sites with SURVEY="OUTSIDE": use only SITES.xlsx data, leave survey-specific fields blank
- Preserve all sites from SITES.xlsx (502 total)
- Add any valid NEW SITE entries from LOGSUMMARY

### 6. Output Generation

**Create file: `COUNTSHEET_TEMPLATE_{YEAR}.xlsx`**
- Sort by: REGION, then SUBSITE alphabetically
- Apply formatting:
  - Header row: Bold
  - Freeze top row
  - Auto-fit column widths
  - If possible, highlight rows with FLAGS in yellow

**Generate summary report:**
```
Count Sheet Generation Summary for {YEAR}
==========================================
Total sites in template: {count}
  - OTTER (surveyed): {count}
  - MISSED (planned but not surveyed): {count}
  - OUTSIDE (not planned): {count}
  
Flags raised:
  - NEW SITE: {count} sites
  - NEEDS_REVIEW: {count} sites
  
Output file: COUNTSHEET_TEMPLATE_{YEAR}.xlsx
```

### 7. Quality Checks

Before finalizing, verify:
- All sites from SITES.xlsx are present in output
- No duplicate SUBSITE entries (except flagged ones)
- SURVEY column only contains: "OTTER", "MISSED", "OUTSIDE", or blank
- COUNTTYPE only contains: 3, 4, or blank
- All FLAGS are documented in summary
- Column count matches expected (34 columns)

## Error Handling

**If input files are missing:**
```
Error: Required file not found
Expected files in current directory:
  - SITES.xlsx
  - {YEAR}_LOGSummary.xlsx

Please ensure files are present and try again.
```

**If required columns are missing:**
```
Error: Missing required columns in {filename}
Expected: {list of columns}
Found: {list of actual columns}
Missing: {list of missing columns}
```

**If data validation fails:**
- Continue processing
- Flag issues in FLAGS column
- Include in summary report
- Do NOT halt execution

## Usage Examples

**Example 1: Generate 2024 count sheet**
```bash
User: "Generate count sheet from 2024 log summary"
or
User: "Create the 2024 template"
```

**Example 2: Generate 2023 count sheet**
```bash
User: "Build count sheet for 2023"
or
User: "Process 2023_LOGSummary.xlsx"
```

**Example 3: Re-run with corrections**
```bash
User: "Regenerate 2024 count sheet"
```

## Output Files

**Primary output:**
- `COUNTSHEET_TEMPLATE_{YEAR}.xlsx` - Formatted Excel file ready for manual counting

**Intermediate files (optional, for debugging):**
- Can save intermediate merged data as CSV if requested

## Notes for Implementation

- Use pandas for Excel file handling
- Use openpyxl for Excel formatting features
- Handle NaN/null values appropriately (convert to blanks in output)
- Preserve data types (dates as dates, numbers as numbers)
- Case-insensitive SUBSITE matching recommended
- Trim whitespace from SUBSITE values before matching

## Success Criteria

Template generation is successful when:
1. Output file is created without errors
2. All sites from SITES.xlsx are present
3. Survey status logic correctly applied
4. Data validation flags are accurate
5. Summary report is clear and actionable
6. File can be opened in Excel without errors
