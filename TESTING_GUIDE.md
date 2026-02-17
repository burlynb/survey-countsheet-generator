# Testing Guide for Count Sheet Generator

## Setup (One-Time)

Your project lives at:
```
C:\Users\burly\OneDrive\Documents\GitHub\survey-countsheet-generator\
```

### Verify Folder Structure

Make sure you have:
```
survey-countsheet-generator\
├── skills\
│   └── generate-countsheet.md
├── scripts\
│   └── generate_countsheet.py
├── inputs\
│   ├── SITES.xlsx
│   ├── 2024_LOGSummary.xlsx
│   ├── 2023_LOGSummary.xlsx   ← upload when ready
│   ├── 2022_LOGSummary.xlsx   ← upload when ready
│   └── 2021_LOGSummary.xlsx   ← upload when ready
├── outputs\
│   └── (empty for now - generated files go here)
├── README.md
├── TESTING_GUIDE.md
└── .gitignore
```

### Install Python Dependencies (if not done yet)

```powershell
pip install pandas openpyxl
```

---

## Test 1: Generate 2024 Count Sheet (Development)

**Goal:** Verify the skill works correctly on your current year's data.

### Steps

1. Open PowerShell

2. Navigate to inputs folder:
   ```powershell
   cd C:\Users\burly\OneDrive\Documents\GitHub\survey-countsheet-generator\inputs
   ```

3. Start Claude Code:
   ```powershell
   claude
   ```

4. Type this prompt:
   ```
   Generate count sheet from 2024 log summary using the skill in ..\skills\generate-countsheet.md
   ```

5. Wait for Claude to finish. It will show a summary report.

6. Check the output folder:
   ```
   C:\Users\burly\OneDrive\Documents\GitHub\survey-countsheet-generator\outputs\
   ```
   You should see `COUNTSHEET_TEMPLATE_2024.xlsx`

### Verification Checklist

- [ ] Output file exists in `outputs\` folder
- [ ] File opens in Excel without errors
- [ ] Header row is bold
- [ ] Top row is frozen when scrolling
- [ ] FLAGS column (Column A) exists
- [ ] Flagged rows are highlighted in yellow
- [ ] SURVEY column contains only: OTTER, MISSED, OUTSIDE, or blank
- [ ] COUNTTYPE column contains only: 3, 4, or blank
- [ ] Spot check 10 random sites - are dates, times, counts correct?

### Record Results

| Metric | Value |
|--------|-------|
| Total sites in output | |
| OTTER (surveyed) | |
| MISSED (planned, not surveyed) | |
| OUTSIDE (not planned) | |
| NEW SITE flags | |
| NEEDS_REVIEW flags | |
| Time to generate (seconds) | |
| Errors encountered | |
| Spot check accuracy (out of 10) | |

---

## Test 2: Generate 2023 Count Sheet (Validation 1)

**Goal:** Confirm the skill works on a different year without any modifications.

### Steps

1. Make sure `2023_LOGSummary.xlsx` is in the `inputs\` folder

2. In Claude Code (still open) type:
   ```
   Generate count sheet from 2023 log summary
   ```

3. Check `outputs\` for `COUNTSHEET_TEMPLATE_2023.xlsx`

4. Complete the same verification checklist as Test 1

### Record Results (same table as above for 2023)

---

## Test 3: Generate 2022 Count Sheet (Validation 2)

Repeat Test 2 steps using 2022 data.

---

## Test 4: Generate 2021 Count Sheet (Validation 3)

Repeat Test 2 steps using 2021 data.

---

## Benchmark Analysis

After all four tests, fill in this comparison table for your write-up:

| Metric | 2024 | 2023 | 2022 | 2021 |
|--------|------|------|------|------|
| Total sites | | | | |
| OTTER | | | | |
| MISSED | | | | |
| OUTSIDE | | | | |
| NEW SITE flags | | | | |
| NEEDS_REVIEW flags | | | | |
| Time (seconds) | | | | |
| Spot check accuracy | /10 | /10 | /10 | /10 |

### Accuracy Check Methodology

For each year, manually verify 10 randomly selected sites:
1. Pick 10 sites from the output template
2. Look them up in the original LOGSummary
3. Confirm: SURVEY status, DATE, TIME, COUNT, COUNTTYPE, PHOTO, SURVEY NOTES
4. Record: (correct / 10) = ____% accuracy

---

## Manual vs. Automated Comparison

Include this in your write-up to demonstrate value:

**Manual process (before automation):**
1. Open SITES.xlsx, LOGSummary, and blank template simultaneously
2. For each of 500+ sites: look up in both files, copy fields one by one
3. Manually calculate SURVEY status, COUNTTYPE, PHOTO
4. Check for missing sites, duplicate entries, MML_ID mismatches
5. Apply formatting
6. **Estimated time: 2-3 hours**

**Automated process (with skill):**
1. Navigate to folder (30 seconds)
2. Type one sentence prompt (10 seconds)
3. Wait for output (~30 seconds)
4. Review flags and spot check (~5 minutes)
5. **Estimated time: under 10 minutes total**

**Time savings: ~2+ hours per survey year**

---

## Common Issues & Fixes

### "Skill not found"
Make sure to reference the path in your prompt:
```
...using the skill in ..\skills\generate-countsheet.md
```

### "SITES.xlsx not found"
- Confirm you are running from the `inputs\` directory
- File name must be exactly `SITES.xlsx`

### Output file already exists
Claude Code will overwrite it. If you want to keep previous outputs, rename them first:
```powershell
rename COUNTSHEET_TEMPLATE_2024.xlsx COUNTSHEET_TEMPLATE_2024_v1.xlsx
```

### Python not found
Install from python.org (choose "Add to PATH" during install), then restart PowerShell

### pandas/openpyxl not installed
```powershell
pip install pandas openpyxl
```

---

## Screenshots to Capture for Tutorial Write-Up

1. PowerShell window showing `claude --version`
2. Navigating to the inputs folder
3. Claude Code starting up
4. Typing the generate prompt
5. Claude's processing output / summary report
6. `outputs\` folder showing the generated file
7. Excel file open showing formatted data
8. Flagged rows highlighted in yellow
9. Spot check: same site in LOGSummary vs. output template side by side
10. Benchmark results table

---

## Iteration Log

Document any changes you make to the skill here (required for assignment):

| Change # | What Changed | Why | Did It Improve Output? |
|----------|-------------|-----|----------------------|
| 1 | | | |
| 2 | | | |

---

## Success Criteria

Your skill is successful when:
- [x] Generates output file without errors on first run
- [x] All 502+ sites from SITES.xlsx are present in output
- [x] Survey status correctly assigned (OTTER/MISSED/OUTSIDE)
- [x] Data accurately transferred from both source files
- [x] Flags correctly identify NEW SITE and NEEDS_REVIEW cases
- [x] Works consistently across 2021, 2022, 2023, and 2024 data
- [x] Output file is immediately usable for manual counting work
- [x] Significant time savings vs. manual process demonstrated
