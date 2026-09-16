# IEC Quality Portal Integration Plan

## Objective

Integrate the local Customer Complaint TAT dashboard into the existing CFR Watch Board Streamlit app with a single IEC Quality Portal entry point. Preserve existing CFR Watch Board behavior.

## Existing CFR structure

- Entry point: `streamlit_app.py`
- CFR parsing and calculations: `cfr_watch_analyzer.py`
- ACT persistence workbook: `ACT table.xlsx`
- Streamlit Cloud target: `https://iec-cfr-watch-board.streamlit.app/`
- Authentication already exists through Streamlit secrets: `APP_PASSWORD` and optional `APP_USERNAME`.

## Existing TAT structure

- Local runnable source: `C:\=Codex study==\20260528- Field Escalation TAT\dashboard_app.py`
- Local output renderer: Pillow PNG plus JSON summary
- Workbook ingestion: uploaded Excel workbook, `New Template` worksheet
- Current migration approach: copy the renderer into `tat_dashboard.py` and call it from the Streamlit page.

## Required changes

- Add IEC Quality Portal home after login.
- Add navigation for Home, Customer Complaint TAT, CFR Watch Board, and Logout.
- Wrap existing CFR page flow in `render_cfr_watch_board()` without changing CFR calculations.
- Add Streamlit TAT page with workbook upload, feedback month input, PNG preview, PNG download, and JSON download.
- Add `pillow` to `requirements.txt`.

## Authentication design

- Reuse existing `password_gate()`.
- Do not hard-code account or password.
- Continue using Streamlit secrets for credentials.
- Keep TAT and CFR protected behind the same login.

## Cloud incompatible dependencies

- Local file paths from the local TAT portal are not used in the cloud flow.
- TAT data source is Streamlit file upload.
- Temporary files are used only during a request.

## Deployment impact

- No second production website is required.
- Existing Streamlit Cloud app remains the deployment target.
- Existing CFR upload, filters, charts, ACT persistence, and group compare functions are preserved.

## Validation plan

- Compile `streamlit_app.py`, `tat_dashboard.py`, and CFR analyzer modules.
- Run the Streamlit app locally.
- Verify Login, Portal, TAT upload/export, CFR upload path, and navigation.
- Compare a TAT output against the local golden dashboard for the same workbook and month.
