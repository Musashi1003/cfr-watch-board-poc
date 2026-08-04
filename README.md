# Sites Prototype

This folder contains an isolated static prototype for previewing a new sites entry experience without changing any existing LOCAL UI.

It now also includes a connected pilot view using sample output from:

- `C:\=Codex study==\20260430-Monthly Shipment Analyze Local`

## Files

- `index.html`: main prototype page
- `styles.css`: layout and visual styling
- `script.js`: simple interaction for swapping mock views
- connected sample data rendering for Monthly Shipment Analyze Local

## Preview

Open `index.html` in a browser to review the simulation locally.

## Backend Prototype

For a server-side employee ID login prototype, run:

- `start_server.bat`

Then open:

- `http://127.0.0.1:8088`

This version validates the employee ID on the server and creates a session cookie before showing the dashboard.

Current demo allowlist:

- `IEC950458`

To let coworkers on the same LAN or VPN upload files and use the CFR watch board, run:

- `start_server_lan.bat`

Then share the coworker URL printed in the command window, for example:

- `http://172.20.10.2:8088/cfr-watch`

LAN mode settings:

- `SITES_HOST=0.0.0.0` listens on the machine's network interfaces.
- `SITES_ALLOWED_EMPLOYEE_IDS=*` accepts any non-empty employee ID for the current internal prototype.
- For a restricted pilot, replace `*` with comma-separated IDs such as `IEC950458,IEC123456`.
- Coworkers must be on the same reachable network or VPN.
- If coworkers cannot open the page, Windows Defender Firewall or company network policy may need to allow inbound TCP `8088`.

## Excel Mail Summary Generator

After signing in, open:

- `http://127.0.0.1:8088/mail-summary`

Upload the monthly shipment workbook, such as `ASUS & HP AIO shipment summary_2026_May_20260605_Ver.01 (1).xlsx`, to generate mail-style summary content.

Current data rules:

- HP-AIO BPS / CPS values are read from the `HP-AIO` sheet.
- ASUS Gaming / PC and OEM / ODM mix are derived from the `ASUS` sheet.
- Customer-level total comparison is aligned to the workbook `Total` sheet because those figures match the approved mail summary numbers.
- Territory mix is not generated yet because the needed region fields are not present in the `HP-AIO` and `ASUS` source sheets.

## Weekly CFR Watch Board

After signing in, open:

- `http://127.0.0.1:8088/cfr-watch`

Upload the weekly Gaming NB and PC NB CFR workbooks together. The page reads the `raw data` sheet from each workbook and builds an interactive watch board.

Current views:

- Main view: `ORG_MODEL(PRODUCT_DESC)`, `Segment`, or `ODM_OEM`
- Breakdown: `MUC_MODULE` or `PROBLEM_Mapping`
- Filters: model, segment, and ODM/OEM
- Filters: model, segment, ODM/OEM, MUC module, and problem mapping
- Visuals: weekly trend, source mix, share view, Pareto table, cross-matrix heatmap, and top watch combinations
- CFR card rule: `derived ACT = IW Failure Q'ty / 2026 CFR(A) for model` from `SUMMARY_IEC`; filtered CFR is then calculated as `filtered raw-data failure count / derived ACT` for the selected model scope.

## Streamlit Cloud POC

The cloud POC uses the same CFR analysis engine in `cfr_watch_analyzer.py`, but runs through Streamlit instead of the local WSGI server.

Cloud entry files:

- `streamlit_app.py`: Streamlit dashboard UI for upload, filters, KPI cards, trend, ratio, Pareto, and matrix.
- `requirements.txt`: Python packages for Streamlit Cloud.
- `.streamlit/config.toml`: Streamlit runtime settings.

Recommended POC flow:

1. Push this folder to a GitHub repository.
2. In Streamlit Community Cloud, create a new app from that repository.
3. Set the app entry point to `streamlit_app.py`.
4. Deploy the app and share the generated Streamlit URL with pilot users.
5. Ask users to upload the weekly Gaming NB and PC NB CFR workbooks together.

Important POC notes:

- Uploaded Excel files are processed at runtime and are not written into the repository.
- Do not upload customer-confidential data to a public app unless your internal policy allows it.
- For internal production use, prefer company-approved hosting such as a company server or Azure Web App with access control.
- Streamlit Cloud is best for a quick POC and user feedback, not final governance.

### Streamlit POC access control

The Streamlit POC has an app-level password gate. Credentials must be configured in Streamlit Cloud secrets, not committed to GitHub.

In Streamlit Cloud:

1. Open the app settings.
2. Go to `Secrets`.
3. Add:

```toml
APP_USERNAME = "iec"
APP_PASSWORD = "replace-with-a-strong-shared-password"
```

`APP_USERNAME` is optional. `APP_PASSWORD` is required. If `APP_PASSWORD` is missing, the app will not show the upload dashboard.

To let the app permanently save newly uploaded ACT history back to `data/activation_history.csv`, also add a GitHub fine-grained token with contents read/write permission for this repository:

```toml
ACT_HISTORY_GITHUB_TOKEN = "github_pat_..."
ACT_HISTORY_GITHUB_REPO = "Musashi1003/cfr-watch-board-poc"
ACT_HISTORY_GITHUB_BRANCH = "main"
```

If `ACT_HISTORY_GITHUB_TOKEN` is not configured, uploaded ACT values are still available in the current session and the app will offer an updated `activation_history.csv` download, but Streamlit restarts will not preserve those new rows automatically.

### ACT table maintenance

`ACT table.xlsx` is the preferred ACT database when it is available. The app reads this workbook before falling back to `data/activation_history.csv`, so manually maintained ACT values remain the source of truth.

- `2025 ACT` uses `MODEL_GROUP` as the model key.
- `2026 ACT` uses `ORG_MODEL(PRODUCT_DESC)` as the model key.
- The app decides which sheet to use from the raw-data `開賣年度` column.
- Existing week/model values in `ACT table.xlsx` are preserved and not overwritten by CFR-derived estimates.
- Missing latest-week values are calculated from the uploaded workbook and written into the ACT table update.

To let Streamlit Cloud save `ACT table.xlsx` back to GitHub, use the same token settings above and optionally set:

```toml
ACT_TABLE_GITHUB_PATH = "ACT table.xlsx"
```

If GitHub write access is not configured, the page offers an updated `ACT table.xlsx` download after upload.

This is a POC control layer only. For formal production use, use Streamlit private viewer settings, company SSO, Azure Web App authentication, or an internal server approved by IT.
