# Rent Roll Cleaner & Validator

A Python automation tool that transforms messy, multi-row rent roll exports from property management software into clean, analysis-ready Excel workbooks — and validates every number against the source.

## The Problem

Analysts received rent roll exports from property management software (Yardi, RealPage, etc.) where each unit spans multiple rows — one row per charge type (base rent, parking, pet fees, etc.). Before analysis, someone had to manually:
- Pivot all charge rows per unit into a single wide row
- Verify totals matched the source
- Apply consistent date formatting, freeze panes, and filters

This took **4–6 hours per report**.

## What This Does

**`clean_rent_roll.py`** — parses and restructures the export:
- Auto-detects the header row (handles blank rows, report titles, and two-row merged headers)
- Forward-fills unit metadata across charge rows
- Pivots charge codes into columns (base rent, parking, pet rent, waste, etc.)
- Normalizes unit IDs so `1001` and `1001.0` match correctly
- Outputs a clean, analyst-ready Excel file with frozen headers, auto-filter, and correct date formatting

**`validate_rent_roll.py`** — verifies the output against the source:
- Confirms every unit from the original appears in the cleaned file
- Checks each charge amount per unit matches the source (tolerance ≤ $0.01)
- Validates that the `Total Charges` column matches the sum of individual charge columns
- Checks all metadata fields (move-in dates, lease expiration, deposits, balance)
- Writes a detailed `validation_report.txt` and a validated Excel workbook
- Exits with a non-zero code on any failure so it can be wired into CI

## Impact

- Reduced per-report processing time from **4–6 hours to under 5 minutes**
- Eliminated manual transcription errors in charge aggregation
- Validation script catches any regression if the source export format changes

## Try It

```bash
# Install dependencies
python3 -m venv venv
source venv/bin/activate   # Windows: venv\Scripts\activate
pip install -r requirements.txt

# Generate the included sample rent roll
python generate_sample_data.py

# Clean it
python clean_rent_roll.py

# Validate the output
python validate_rent_roll.py

# Results:
#   Rent_Roll_Organized.xlsx   — cleaned workbook
#   Rent_Roll_Validated.xlsx   — formatted validated copy
#   validation_report.txt      — full pass/fail report
```

## Using Your Own File

1. Place your rent roll export in this folder.
2. In `clean_rent_roll.py`, set `INPUT_FILE` to your filename.
3. Run `python clean_rent_roll.py` then `python validate_rent_roll.py`.

The header detection handles common export layouts automatically. If it fails, see Troubleshooting below.

## Project Structure

```
clean_rent_roll.py        # Parsing and restructuring logic
validate_rent_roll.py     # Validation and reporting
generate_sample_data.py   # Generates sample_rent_roll.xlsx for demo
sample_rent_roll.xlsx     # Fake data — safe to commit
requirements.txt
```

## Troubleshooting

| Error | Fix |
|---|---|
| `ModuleNotFoundError: No module named pandas` | Run `source venv/bin/activate && pip install -r requirements.txt` |
| `Could not detect header row` | The export layout changed; update the `find_header_row` detection in `clean_rent_roll.py` |
| Validation shows `FAIL` | Open `validation_report.txt` — each mismatch is listed with unit, field, original value, and cleaned value |

## Stack

- Python 3.10+
- pandas — data reshaping and aggregation
- openpyxl — Excel read/write and formatting
