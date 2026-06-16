# ATLAS Report Generator

Use `tools/atlas_report_generator.py` to turn RISE monthly source reports into an ATLAS-ready upload package.

## HTML Workflow

Start the local HTML front end:

```bash
python3 tools/atlas_generator_server.py
```

Then open:

```text
http://127.0.0.1:8765
```

Choose one or more reports, click **Generate ATLAS Upload Package**, then download the combined Excel workbook.
The CSV ZIP and individual CSVs remain available for audit or table-by-table loading.
The HTML app runs locally on your machine.

## Folder Workflow

Create a monthly drop folder and place the reports inside it:

```bash
mkdir -p atlas_reports/2026-04
```

Then run:

```bash
python3 tools/atlas_report_generator.py \
  --input atlas_reports/2026-04 \
  --output outputs/atlas_uploads/2026-04
```

You can also point the generator at specific files:

```bash
python3 tools/atlas_report_generator.py \
  --file "/path/to/02 - RISE - Box Score.xlsx" \
  --file "/path/to/RISE_Baymeadows_Market.xlsx" \
  --output outputs/atlas_uploads/2026-04
```

## Outputs

The generator writes `atlas_combined_upload.xlsx` with one tab per target table.
It also writes one CSV per target table:

- `monthly_upload.csv` - the main ATLAS monthly community-level upload.
- `boxscore_floor_plan_monthly.csv` - floor-plan detail from Box Score reports.
- `boxscore_pulse_monthly.csv` - move-ins, notices, move-outs, renewals, leased activity.
- `boxscore_lead_monthly.csv` - lead-source and tour activity.
- `boxscore_application_monthly.csv` - application/conversion activity.
- `boxscore_make_ready_monthly.csv` - ready/not-ready vacancy status.
- `market_property_snapshot.csv` - subject and comp property market snapshots.
- `market_floor_plan_snapshot.csv` - comp floor-plan rent and availability detail.
- `market_unit_snapshot.csv` - comp unit-level listing detail.
- `rent_roll_unit_monthly.csv` - rent-roll unit-level detail.
- `renewal_tracker_monthly.csv` - multi-month renewal tracker resident/offer detail.
- `trending_occupancy_monthly.csv` - multi-month occupancy trend rows by community from CSV or workbook trend exports.
- `source_file_log.csv` - what was detected and parsed from each file.
- `exceptions.csv` - items to review before production load.
- `manifest.csv` - row counts by output table.

## Report Month Rules

The generator tries to infer the month from each source:

- Box Score: report period such as `04/01/2026 - 04/30/2026`.
- Market files: workbook updated date or filename timestamp.
- Rent Roll PDF: report month such as `Mar 2026`.
- Renewal Tracker: month tabs such as `March 2026`, `April 2026`.
- Trending Occupancy: each row's date range.

Use a fallback month when files do not include a clean date:

```bash
python3 tools/atlas_report_generator.py \
  --input atlas_reports/2026-04 \
  --report-month 2026-04 \
  --output outputs/atlas_uploads/2026-04
```

Use `--force-report-month` only when you want every file assigned to that month regardless of source dates.

## PII

Rent-roll resident names, balances, deposits, and renewal tracker resident names/phone
numbers are masked by default. To include them:

```bash
python3 tools/atlas_report_generator.py \
  --input atlas_reports/2026-04 \
  --include-pii \
  --output outputs/atlas_uploads/2026-04-with-pii
```

Only use `--include-pii` if the ATLAS destination is approved for sensitive resident-level data.

## HTML Checkbox Guide

- **Force fallback month:** usually leave unchecked. Check it only when source report dates are missing, wrong, or when every uploaded file should be intentionally reclassified into the fallback month.
- **Include resident PII fields:** usually leave unchecked. Check it only for an approved restricted resident-level workflow that needs resident names, phone numbers, balances, or deposits.
- **Skip PDF rent-roll parsing:** check it when PDFs are attached but the run does not need unit-level rent-roll detail, or when you want a faster Excel/CSV-only package.

## Loading Guidance

Use `atlas_combined_upload.xlsx` when you want one reviewed package with all tabs.
Load the `monthly_upload` tab into the current ATLAS monthly upload flow.

The other CSVs are detail tables. They should be loaded only if ATLAS or Power BI has matching tables/tabs for those grains.

Trending Occupancy files populate occupancy, move-ins, and move-outs. The workbook
export is parsed sheet by sheet, and it can backfill past months when you re-import
an updated trend file.

Always check `exceptions.csv` before production upload. A row marked `REVIEW` usually means one of three things:

- A community appears in the source but is missing from the current ATLAS `Community_Master`.
- A PDF rent roll was parsed heuristically and should be reviewed.
- A source file shape was not recognized.

The generator also writes `WARN` exceptions when a run is only a partial package, such as when Renewal Tracker or Rent Roll inputs were not included. That makes mixed uploads visible in the HTML results page and in `exceptions.csv` before anything is loaded downstream.
