from __future__ import annotations

import hmac
import os
from datetime import datetime
from pathlib import Path
from tempfile import NamedTemporaryFile

import altair as alt
import pandas as pd
import streamlit as st

from cfr_watch_analyzer import (
  FILTER_FIELDS,
  WorkbookUpload,
  analyze_dataset,
  dataset_options,
  parse_workbooks,
)


st.set_page_config(
  page_title="CFR Watch Board",
  layout="wide",
)


FILTER_LABELS = {
  "model": "Model",
  "segment": "Segment",
  "odm_oem": "ODM/OEM",
  "muc_module": "Module",
  "problem_mapping": "Problem",
}

BAR_COLOR = "#1a9295"
BAR_LIGHT_COLOR = "#8fd6d5"
LINE_COLOR = "#075f73"
ACCENT_GREEN = "#12805c"
ACCENT_RED = "#d83b35"
TEXT_DARK = "#071316"


def read_secret(name: str) -> str:
  value = os.environ.get(name, "")
  if value:
    return value
  try:
    return str(st.secrets.get(name, "") or "")
  except Exception:
    return ""


def password_gate() -> bool:
  expected_username = read_secret("APP_USERNAME").strip()
  expected_password = read_secret("APP_PASSWORD")

  if not expected_password:
    st.error("Access control is not configured. Ask the app owner to set APP_PASSWORD in Streamlit secrets.")
    return False

  if st.session_state.get("authenticated"):
    return True

  st.markdown("### CFR Watch Board Login")
  if expected_username:
    st.info(f"Account: `{expected_username}`. Please ask the dashboard owner for the password.")
  else:
    st.info("Please enter the shared password provided by the dashboard owner.")
  with st.form("login_form"):
    username = st.text_input("Account", value="", disabled=not bool(expected_username))
    password = st.text_input("Password", type="password")
    submitted = st.form_submit_button("Sign in")

  if submitted:
    username_ok = True
    if expected_username:
      username_ok = hmac.compare_digest(username.strip(), expected_username)
    password_ok = hmac.compare_digest(password, expected_password)
    if username_ok and password_ok:
      st.session_state["authenticated"] = True
      st.rerun()
    st.error("Account or password is incorrect.")

  return False


def apply_page_style():
  st.markdown(
    """
    <style>
      .block-container {
        padding-top: 1.8rem;
        padding-bottom: 2.5rem;
        max-width: 1480px;
      }
      .stApp {
        background: #f6f9fb;
      }
      [data-testid="stSidebar"] {
        background: #f8fbfc;
        border-right: 1px solid #dce8ea;
      }
      [data-testid="stSidebar"] h1,
      [data-testid="stSidebar"] h2,
      [data-testid="stSidebar"] h3,
      [data-testid="stSidebar"] p,
      [data-testid="stSidebar"] label {
        color: #122426;
      }
      [data-testid="stSidebar"] [data-baseweb="select"] > div {
        background: #ffffff;
        border-color: #cddbdd;
      }
      [data-testid="stSidebar"] [data-baseweb="select"] span,
      [data-testid="stSidebar"] [data-baseweb="select"] input {
        color: #122426 !important;
      }
      [data-testid="stSidebar"] [data-baseweb="tag"] {
        background-color: #dff3f3 !important;
        border-color: #b6dfdf !important;
      }
      [data-testid="stSidebar"] [data-baseweb="tag"] span {
        color: #075f73 !important;
        font-weight: 650;
      }
      .cfr-title {
        color: #122426;
        font-size: 2.1rem;
        font-weight: 780;
        margin-bottom: 0.15rem;
      }
      .cfr-subtitle {
        color: #5f6f72;
        font-size: 0.98rem;
        margin-bottom: 1.1rem;
      }
      .dashboard-head {
        display: flex;
        align-items: flex-end;
        justify-content: space-between;
        gap: 1rem;
        margin: 1rem 0 0.8rem 0;
      }
      .dashboard-head h2 {
        font-size: 1.45rem;
        margin: 0 0 0.2rem 0;
      }
      .dashboard-head p {
        color: #62767b;
        margin: 0;
      }
      .refresh-note {
        color: #557179;
        font-size: 0.85rem;
        white-space: nowrap;
      }
      .metric-card {
        background: linear-gradient(135deg, #ffffff 0%, #f9fffd 100%);
        border: 1px solid #e5e0d8;
        border-radius: 8px;
        padding: 1rem 1.1rem;
        min-height: 150px;
        box-shadow: 0 10px 24px rgba(16, 35, 42, 0.06);
        border-top: 4px solid var(--accent, #138a8e);
      }
      .metric-label {
        color: #4d6568;
        font-size: 0.82rem;
        margin-bottom: 0.45rem;
      }
      .metric-value {
        color: #071316;
        font-size: 1.75rem;
        font-weight: 800;
        line-height: 1.1;
      }
      .metric-note {
        color: #607478;
        font-size: 0.82rem;
        margin-top: 0.65rem;
      }
      .metric-delta {
        display: inline-block;
        color: var(--accent, #138a8e);
        font-size: 0.78rem;
        font-weight: 700;
        margin-top: 0.75rem;
      }
      .sparkline {
        width: 100%;
        height: 32px;
        margin-top: 0.75rem;
      }
      h2, h3 {
        color: #0c1c1f;
      }
      div[data-testid="stDataFrame"] {
        border: 1px solid #e5e0d8;
        border-radius: 8px;
      }
    </style>
    """,
    unsafe_allow_html=True,
  )


def pct(value: float | None) -> str:
  if value is None:
    return "N/A"
  return f"{value * 100:.2f}%"


def whole(value: float | int | None) -> str:
  if value is None:
    return "N/A"
  return f"{value:,.0f}"


def sparkline_svg(values: list[int], color: str) -> str:
  if len(values) < 2:
    return ""

  width = 220
  height = 34
  padding = 3
  min_value = min(values)
  max_value = max(values)
  span = max(max_value - min_value, 1)
  points = []
  for index, value in enumerate(values):
    x = padding + index * ((width - padding * 2) / (len(values) - 1))
    y = height - padding - ((value - min_value) / span) * (height - padding * 2)
    points.append(f"{x:.1f},{y:.1f}")
  return (
    f'<svg class="sparkline" viewBox="0 0 {width} {height}" preserveAspectRatio="none">'
    f'<polyline points="{" ".join(points)}" fill="none" stroke="{color}" stroke-width="3" '
    'stroke-linecap="round" stroke-linejoin="round" />'
    f'<circle cx="{points[-1].split(",")[0]}" cy="{points[-1].split(",")[1]}" r="3.2" fill="{color}" />'
    '</svg>'
  )


def dashboard_header(result: dict):
  latest_week = result["kpis"]["latest_week"]
  st.markdown(
    f"""
    <div class="dashboard-head">
      <div>
        <h2>Overview Dashboard</h2>
        <p>Carry risk and CTX-style watch signals from weekly raw-data uploads.</p>
      </div>
      <div class="refresh-note">Data updated: {datetime.now().strftime("%Y-%m-%d %H:%M")} 繚 Latest week: {latest_week}</div>
    </div>
    """,
    unsafe_allow_html=True,
  )


def write_uploads_to_temp_files(uploaded_files) -> tuple[list[WorkbookUpload], list[Path]]:
  workbooks: list[WorkbookUpload] = []
  temp_paths: list[Path] = []
  for uploaded_file in uploaded_files:
    suffix = Path(uploaded_file.name).suffix
    with NamedTemporaryFile(delete=False, suffix=suffix) as temp_file:
      temp_file.write(uploaded_file.getbuffer())
      temp_path = Path(temp_file.name)
    temp_paths.append(temp_path)
    workbooks.append(WorkbookUpload(path=temp_path, filename=uploaded_file.name))
  return workbooks, temp_paths


def parse_uploaded_files(uploaded_files) -> dict:
  workbooks, temp_paths = write_uploads_to_temp_files(uploaded_files)
  try:
    return parse_workbooks(workbooks)
  finally:
    for temp_path in temp_paths:
      try:
        temp_path.unlink(missing_ok=True)
      except PermissionError:
        pass


def selected_filters(records: list[dict]) -> dict[str, list[str]]:
  selections: dict[str, list[str]] = {}
  with st.sidebar:
    st.header("Filters")
    st.caption("Leave a filter empty to include all values.")
    for key in FILTER_FIELDS:
      options = dataset_options(records, selections).get(key, [])
      with st.expander(FILTER_LABELS[key], expanded=False):
        selections[key] = st.multiselect(
          "Select values",
          options,
          default=[],
          key=f"filter_{key}",
          label_visibility="collapsed",
        )
  return selections


def metric_row(result: dict):
  kpis = result["kpis"]
  cols = st.columns(4)
  trend_values = [row["count"] for row in result.get("trend", [])]
  week_delta = kpis["week_delta"]
  delta_color = ACCENT_RED if week_delta > 0 else ACCENT_GREEN if week_delta < 0 else "#557179"
  cards = [
    {
      "label": "Filtered CFR",
      "value": pct(kpis["filtered_cfr"]),
      "note": f"{whole(kpis['filtered_failure_qty'])} failures / {whole(kpis['derived_act'])} derived ACT",
      "accent": BAR_COLOR,
      "delta": f"Target CFR: {pct(kpis['target_cfr'])}",
      "spark": "",
    },
    {
      "label": "Failure Qty",
      "value": whole(kpis["filtered_failure_qty"]),
      "note": "Filtered raw-data failures",
      "accent": ACCENT_RED,
      "delta": f"Latest week: {whole(kpis['latest_count'])}",
      "spark": sparkline_svg(trend_values, ACCENT_RED),
    },
    {
      "label": "Derived ACT",
      "value": whole(kpis["derived_act"]),
      "note": f"From SUMMARY_IEC CFR rule; {whole(kpis['act_model_count'])} ACT models",
      "accent": LINE_COLOR,
      "delta": "Current-upload estimate only",
      "spark": "",
    },
    {
      "label": "Latest WoW",
      "value": f"{week_delta:+,}",
      "note": f"{kpis['previous_week']} to {kpis['latest_week']}",
      "accent": delta_color,
      "delta": "Failure count change",
      "spark": sparkline_svg(trend_values[-6:], delta_color),
    },
  ]
  for column, card in zip(cols, cards):
    column.markdown(
      f"""
      <div class="metric-card" style="--accent: {card['accent']};">
        <div class="metric-label">{card['label']}</div>
        <div class="metric-value">{card['value']}</div>
        <div class="metric-note">{card['note']}</div>
        <div class="metric-delta">{card['delta']}</div>
        {card['spark']}
      </div>
      """,
      unsafe_allow_html=True,
    )


def bar_chart(frame: pd.DataFrame, x_column: str, y_column: str, height: int = 260):
  chart = (
    alt.Chart(frame)
    .mark_bar(color=BAR_COLOR, cornerRadiusTopLeft=3, cornerRadiusTopRight=3)
    .encode(
      y=alt.Y(f"{x_column}:N", sort="-x", axis=alt.Axis(title=None, labelLimit=180)),
      x=alt.X(f"{y_column}:Q", axis=alt.Axis(title=None)),
      tooltip=[
        alt.Tooltip(f"{x_column}:N", title="Item"),
        alt.Tooltip(f"{y_column}:Q", title="Count"),
      ],
    )
    .properties(height=height)
  )
  st.altair_chart(chart, width="stretch")


def rows_to_pareto_frame(rows: list[dict], limit: int = 12) -> pd.DataFrame:
  frame = pd.DataFrame(rows).head(limit).copy()
  if frame.empty:
    return frame
  frame["cumulative_count"] = frame["count"].cumsum()
  total = frame["count"].sum() or 1
  if "share" not in frame:
    frame["share"] = frame["count"] / total * 100
  frame["cumulative"] = frame["cumulative_count"] / total * 100
  frame["short_label"] = frame["label"].str.slice(0, 26)
  return frame


def render_pareto_frame(title: str, rows: list[dict], empty_message: str):
  st.subheader(title)
  pareto = rows_to_pareto_frame(rows)
  if pareto.empty:
    st.info(empty_message)
    return

  bars = (
    alt.Chart(pareto)
    .mark_bar(color=BAR_LIGHT_COLOR, cornerRadiusTopLeft=3, cornerRadiusTopRight=3)
    .encode(
      x=alt.X("short_label:N", sort=None, axis=alt.Axis(title=None, labelAngle=-35, labelLimit=90)),
      y=alt.Y("count:Q", axis=alt.Axis(title="Failure Qty")),
      tooltip=[
        alt.Tooltip("label:N", title="Item"),
        alt.Tooltip("count:Q", title="Failure Qty"),
        alt.Tooltip("share:Q", title="Share", format=".1f"),
        alt.Tooltip("cumulative:Q", title="Cumulative", format=".1f"),
      ],
    )
  )
  line = (
    alt.Chart(pareto)
    .mark_line(color=LINE_COLOR, strokeWidth=2.8)
    .encode(
      x=alt.X("short_label:N", sort=None),
      y=alt.Y("cumulative:Q", axis=alt.Axis(title="Cumulative %", orient="right")),
    )
  )
  points = (
    alt.Chart(pareto)
    .mark_point(color=LINE_COLOR, filled=True, size=62)
    .encode(
      x=alt.X("short_label:N", sort=None),
      y=alt.Y("cumulative:Q"),
      tooltip=[
        alt.Tooltip("label:N", title="Item"),
        alt.Tooltip("cumulative:Q", title="Cumulative", format=".1f"),
      ],
    )
  )
  st.altair_chart(
    alt.layer(bars, line, points).resolve_scale(y="independent").properties(height=320),
    width="stretch",
  )

  table = pareto.copy()
  table.insert(0, "rank", range(1, len(table) + 1))
  table["share"] = table["share"].map(lambda value: f"{value:.1f}%")
  table["cumulative"] = table["cumulative"].map(lambda value: f"{value:.1f}%")
  st.dataframe(
    table[["rank", "label", "count", "share", "cumulative"]],
    width="stretch",
    hide_index=True,
  )


def render_ratio(title: str, rows: list[dict]):
  st.subheader(title)
  if not rows:
    st.info("No records in the current filter.")
    return
  frame = pd.DataFrame(rows).head(10)
  bar_chart(frame, "label", "count", height=max(240, len(frame) * 28))
  table = frame.copy()
  table.insert(0, "rank", range(1, len(table) + 1))
  table["share"] = table["share"].map(lambda value: f"{value:.1f}%")
  table = table.rename(columns={"label": "item", "count": "failure_qty", "share": "share"})
  st.dataframe(
    table[["rank", "item", "failure_qty", "share"]],
    width="stretch",
    hide_index=True,
  )


def render_trend(result: dict):
  st.subheader("Weekly Failure Trend")
  trend = pd.DataFrame(result["trend"])
  if trend.empty:
    st.info("No trend data in the current filter.")
    return
  chart = (
    alt.Chart(trend)
    .mark_line(point=True, color="#138a8e", strokeWidth=2.5)
    .encode(
      x=alt.X("week:N", axis=alt.Axis(labelAngle=-45, title=None)),
      y=alt.Y("count:Q", axis=alt.Axis(title=None)),
      tooltip=[
        alt.Tooltip("week:N", title="Week"),
        alt.Tooltip("count:Q", title="Failure Qty"),
      ],
    )
    .properties(height=300)
  )
  st.altair_chart(chart, width="stretch")


def render_pareto(result: dict):
  render_pareto_frame(
    "PROBLEM_Mapping Pareto",
    result["pareto"],
    "No problem data in the current filter.",
  )


def render_module_pareto(result: dict):
  render_pareto_frame(
    "MUC_MODULE Pareto",
    result["module_ratio"],
    "No module data in the current filter.",
  )


def render_heatmap(result: dict):
  st.subheader("Watch Matrix Heatmap")
  st.caption("Each cell shows failure count with CFR in parentheses when ACT is available.")
  columns = result["heatmap_columns"]
  rows = result["heatmap"]
  if not columns or not rows:
    st.info("No matrix data in the current filter.")
    return

  count_frame = pd.DataFrame(
    [[cell["count"] for cell in row["values"]] for row in rows],
    index=[row["label"] for row in rows],
    columns=columns,
  )
  cfr_frame = pd.DataFrame(
    [[cell["cfr"] for cell in row["values"]] for row in rows],
    index=[row["label"] for row in rows],
    columns=columns,
  )

  heatmap_rows = []
  has_cfr = cfr_frame.notna().any().any()
  for row_label in count_frame.index:
    for column_label in count_frame.columns:
      count = int(count_frame.loc[row_label, column_label])
      cfr_value = cfr_frame.loc[row_label, column_label]
      color_value = cfr_value * 100 if has_cfr and pd.notna(cfr_value) else count
      heatmap_rows.append(
        {
          "model": row_label,
          "problem": column_label,
          "count": count,
          "cfr_pct": cfr_value * 100 if pd.notna(cfr_value) else None,
          "color_value": color_value,
          "label": f"{count}\n{cfr_value * 100:.2f}%" if count and pd.notna(cfr_value) else str(count),
        }
      )

  frame = pd.DataFrame(heatmap_rows)
  rect = (
    alt.Chart(frame)
    .mark_rect(stroke="#ffffff", strokeWidth=1.5)
    .encode(
      x=alt.X("problem:N", sort=columns, axis=alt.Axis(title=None, labelAngle=-30, labelLimit=120)),
      y=alt.Y("model:N", sort=[row["label"] for row in rows], axis=alt.Axis(title=None, labelLimit=120)),
      color=alt.Color(
        "color_value:Q",
        scale=alt.Scale(scheme="tealblues"),
        legend=alt.Legend(title="CFR %" if has_cfr else "Failure Qty"),
      ),
      tooltip=[
        alt.Tooltip("model:N", title="Model"),
        alt.Tooltip("problem:N", title="Problem"),
        alt.Tooltip("count:Q", title="Failure Qty"),
        alt.Tooltip("cfr_pct:Q", title="CFR %", format=".2f"),
      ],
    )
  )
  text = (
    alt.Chart(frame)
    .mark_text(color="#071316", fontSize=11, fontWeight="bold")
    .encode(
      x=alt.X("problem:N", sort=columns),
      y=alt.Y("model:N", sort=[row["label"] for row in rows]),
      text="label:N",
    )
  )
  st.altair_chart(
    alt.layer(rect, text).properties(height=max(320, len(rows) * 42)),
    width="stretch",
  )


def render_file_summary(parsed: dict):
  files = parsed.get("files", [])
  if not files:
    return
  with st.expander("Loaded weekly raw data", expanded=False):
    for file_info in files:
      st.markdown(
        f"- **{file_info['source_type']}** - `{file_info['filename']}` - {file_info['rows']} rows"
      )


def main():
  apply_page_style()
  st.markdown('<div class="cfr-title">CFR Watch Board</div>', unsafe_allow_html=True)
  st.markdown(
    '<div class="cfr-subtitle">Upload weekly Gaming NB and PC NB CFR workbooks, then filter by model, segment, ODM/OEM, module, and problem.</div>',
    unsafe_allow_html=True,
  )

  if not password_gate():
    return

  with st.sidebar:
    if st.button("Sign out"):
      st.session_state.pop("authenticated", None)
      st.rerun()

  uploaded_files = st.file_uploader(
    "Weekly CFR workbooks",
    type=["xlsx", "xlsm", "xls"],
    accept_multiple_files=True,
    help="Upload the weekly Gaming NB and PC NB files together.",
  )

  if not uploaded_files:
    st.info("Upload weekly CFR raw-data workbooks to start.")
    return

  try:
    parsed = parse_uploaded_files(uploaded_files)
  except Exception as exc:
    st.error(f"CFR analysis failed: {exc}")
    return

  records = parsed.get("records", [])
  if not records:
    st.warning("No raw-data records were found. Please confirm the workbook includes a 'raw data' sheet.")
    return

  render_file_summary(parsed)
  selections = selected_filters(records)
  result = analyze_dataset(
    records,
    summary_by_model=parsed.get("summary_by_model", {}),
    primary_dimension="org_model",
    breakdown_dimension="problem_mapping",
    filters=selections,
  )

  dashboard_header(result)
  metric_row(result)
  st.divider()

  render_trend(result)

  left, right = st.columns(2)
  with left:
    render_ratio("ORG_MODEL(PRODUCT_DESC) Share", result["share"])
  with right:
    render_module_pareto(result)

  left, right = st.columns([1, 1.2])
  with left:
    render_pareto(result)
  with right:
    render_heatmap(result)


if __name__ == "__main__":
  main()

