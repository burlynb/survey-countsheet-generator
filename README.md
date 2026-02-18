# Steller sea lion Survey Count Sheet Generator

Automated generation of count sheet templates from aerial survey field logs using Claude Code AI skills.

## Overview

This project automates the creation of sea otter survey count sheet templates by:
- Reading master site lists and field survey logs
- Applying survey status classification logic
- Validating data integrity
- Generating formatted Excel templates ready for manual counting

**Time Saved:** ~2-3 hours per survey year  
**Error Reduction:** Eliminates manual copy-paste errors and missing sites

---

## Project Structure

```
C:\Users\burly\OneDrive\Documents\GitHub\survey-countsheet-generator\
├── skills\
│   └── generate-countsheet.md          ← AI skill definition
├── scripts\
│   └── generate_countsheet.py          ← Python implementation
├── inputs\
│   ├── SITES.xlsx                      ← Master site list (502 sites)
│   ├── 2024_LOGSummary.xlsx           ← 2024 field data
│   ├── 2023_LOGSummary.xlsx           ← 2023 field data
│   ├── 2022_LOGSummary.xlsx           ← 2022 field data
│   └── 2021_LOGSummary.xlsx           ← 2021 field data
├── outputs\
│   └── (generated templates saved here)
├── README.md                           ← This file
└── .gitignore
```

---

## Quick Start

### Prerequisites

- **Claude Code** installed (see below)
- **Claude Pro or Max subscription** (or API key)
- **Python 3.8+** with pandas and openpyxl

### Installation

1. **Install Claude Code** (if not already installed):
   ```powershell
   # Windows PowerShell (run as Administrator)
   irm https://claude.ai/install.ps1 | iex

   # Verify installation
   claude --version
   ```

2. **Navigate to project folder**:
   ```powershell
   cd C:\Users\burly\OneDrive\Documents\GitHub\survey-countsheet-generator
   ```

3. **Create subdirectories if needed**:
   ```powershell
   mkdir skills, scripts, inputs, outputs
   ```

4. **Place your data files in `inputs\`:**
   - `SITES.xlsx`
   - `2024_LOGSummary.xlsx`
   - `2023_LOGSummary.xlsx`
   - `2022_LOGSummary.xlsx`
   - `2021_LOGSummary.xlsx`

5. **Install Python dependencies**:
   ```powershell
   pip install pandas openpyxl
   ```

---

## Usage

### Method 1: Using Claude Code (Recommended)

1. **Navigate to the inputs folder**:
   ```powershell
   cd C:\Users\burly\OneDrive\Documents\GitHub\survey-countsheet-generator\inputs
   ```

2. **Start Claude Code**:
   ```powershell
   claude
   ```

3. **Generate a count sheet using natural language**:
   ```
   Generate count sheet from 2024 log summary using the skill in ..\skills\generate-countsheet.md
   ```

4. **Claude will**:
   - Read the skill file
   - Load SITES.xlsx and the specified LOGSummary
   - Execute the transformation logic
   - Save output to `..\outputs\`
   - Print a summary report

### Method 2: Python Script Directly

```powershell
cd C:\Users\burly\OneDrive\Documents\GitHub\survey-countsheet-generator\inputs
python ..\scripts\generate_countsheet.py 2024
```

---

## Input File Requirements

### SITES.xlsx (Master Site List)

Required columns: SUBSITE, SUBSITE_ID, PARENTSITE, PARENTSITE_ID, MML_ID, REGION, REGNO, RCA, ROOK, LAT, LON

### YYYY_LOGSummary.xlsx (Field Log)

Required columns: DATE, MML_ID, SUBSITE, PARENTSITE, TIME, COUNT, PASS, PASS DESCRIPTION, ADD, DISTURBANCE, Priority, REGION, REGNO, RCA, ROOK

Naming convention: Must follow pattern `YYYY_LOGSummary.xlsx` (e.g., `2024_LOGSummary.xlsx`)

---

## Output

### COUNTSHEET_TEMPLATE_YYYY.xlsx

Saved to `outputs\` and contains:
- All sites from SITES.xlsx (502+ sites)
- Merged survey data from LOGSummary
- Calculated fields (SURVEY, COUNTTYPE, PHOTO)
- Quality check flags in Column A (NEW SITE, NEEDS_REVIEW)
- Blank columns for manual counting (BULL, SAM, FEM, JUV, PUP, etc.)
- Bold headers, frozen top row, yellow highlighting for flagged rows

---

## Business Logic

### Survey Status (SURVEY column)

| Status | Condition |
|--------|-----------|
| **OTTER** | Site in LOGSummary with a DATE - successfully surveyed |
| **MISSED** | Site in LOGSummary but DATE is empty - planned but not surveyed |
| **OUTSIDE** | Site in SITES.xlsx but not in LOGSummary - not planned this year |

### Count Type (COUNTTYPE column)

| Value | Condition |
|-------|-----------|
| **4** | COUNT column has a value (visual count from aircraft) |
| **3** | PASS column has a value (photographic count) |
| *blank* | Site not surveyed |

### Photo Indicator (PHOTO column)

| Value | Condition |
|-------|-----------|
| **Y** | PASS column is not null |
| *blank* | No photos taken |

### Quality Flags (FLAGS column)

| Flag | Trigger | Action Required |
|------|---------|-----------------|
| **NEW SITE** | SUBSITE in LOGSummary but not in SITES.xlsx | Add to master SITES list |
| **NEEDS_REVIEW** | MML_ID mismatch or duplicate entries found | Investigate discrepancy |

---

## Benchmark Testing

| Test | Year | Purpose |
|------|------|---------|
| Development | 2024 | Build and iterate on current year |
| Validation 1 | 2023 | Blind test - previous year |
| Validation 2 | 2022 | Blind test - older data |
| Validation 3 | 2021 | Blind test - oldest available |

### Results

*(Fill in after testing)*

| Year | Total Sites | OTTER | MISSED | OUTSIDE | Flags | Time (sec) | Accuracy |
|------|-------------|-------|--------|---------|-------|------------|----------|
| 2024 | | | | | | | |
| 2023 | | | | | | | |
| 2022 | | | | | | | |
| 2021 | | | | | | | |

---

## Troubleshooting

**"File not found"** - Run from `inputs\` directory and check file names match exactly

**"Missing columns"** - Column names are case-sensitive; verify against required list above

**"claude not recognized"** - Restart PowerShell; verify `C:\Users\burly\.local\bin` is in PATH

**Wrong year processed** - Specify year directly: `python ..\scripts\generate_countsheet.py 2023`

---

## Author

Created for MSIS 549 - AI and GenAI for Business Applications  
University of Washington Foster School of Business

MIT License
