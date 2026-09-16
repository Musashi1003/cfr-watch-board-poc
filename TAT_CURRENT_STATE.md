# TAT Current State

## Source

- Authoritative local source: `C:\=Codex study==\20260528- Field Escalation TAT\dashboard_app.py`
- Local portal: `http://localhost:8500/tat`
- Migrated module in this repository: `tat_dashboard.py`

## Current workflow

- User uploads an ASUS customer complaint workbook.
- User selects a feedback month in `YYYY-MM` format.
- The renderer reads the workbook sheet whose name starts with `New Template`.
- Output is a presentation-sized dashboard PNG and a JSON summary.

## Reused logic

- Workbook row loading and month selection.
- Feedback month parsing.
- KPI metrics and previous-month comparison.
- YoY monthly case trend.
- QoQ confirmed-case bar chart.
- Problem/category/region/family/weekly/parts charts.
- PNG and JSON export behavior.

## Cloud compatibility notes

- Local absolute workbook paths are not used by the Streamlit Cloud page.
- The portal uses Streamlit file upload and temporary files.
- PNG generation requires `pillow`.
- Excel parsing requires `openpyxl`.
- No database is required for the current migration.

## Validation basis

- Local TAT output remains the golden version.
- For the same workbook and feedback month, Streamlit Cloud output should match the local dashboard metrics and exported PNG/JSON.
