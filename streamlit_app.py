from __future__ import annotations

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


def apply_page_style():
  st.markdown(
    """
    <style>
      .block-container {
        padding-top: 1.8rem;
        padding-bottom: 2.5rem;
        max-width: 1480px;
      }
      [data-testid="stSidebar"] {
        background: #f3f6f6;
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
      .metric-card {
        background: #fffdf8;
        border: 1px solid #e5e0d8;
        border-radius: 8px;
        padding: 1rem 1.1rem;
        min-height: 132px;
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
  cards = [
    ("Filtered CFR", pct(kpis["filtered_cfr"]), f"{whole(kpis['filtered_failure_qty'])} failures / {whole(kpis['derived_act'])} derived ACT"),
    ("Failure Qty", whole(kpis["filtered_failure_qty"]), "Filtered raw-data failures"),
    ("Derived ACT", whole(kpis["derived_act"]), f"From SUMMARY_IEC CFR rule; {whole(kpis['act_model_count'])} ACT models"),
    ("Latest WoW", f"{kpis['week_delta']:+,}", f"{kpis['previous_week']} to {kpis['latest_week']}"),
  ]
  for column, (label, value, note) in zip(cols, cards):
    column.markdown(
      f"""
      <div class="metric-card">
        <div class="metric-label">{label}</div>
        <div class="metric-value">{value}</div>
        <div class="metric-note">{note}</div>
      </div>
      """,
      unsafe_allow_html=True,
    )


def bar_chart(frame: pd.DataFrame, x_column: str, y_column: str, height: int = 260):
  chart = (
    alt.Chart(frame)
    .mark_bar(color="#138a8e", cornerRadiusTopLeft=3, cornerRadiusTopRight=3)
    .encode(
      x=alt.X(f"{x_column}:N", sort="-y", axis=alt.Axis(labelAngle=-45, title=None)),
      y=alt.Y(f"{y_column}:Q", axis=alt.Axis(title=None)),
      tooltip=[
        alt.Tooltip(f"{x_column}:N", title="Item"),
        alt.Tooltip(f"{y_column}:Q", title="Count"),
      ],
    )
    .properties(height=height)
  )
  st.altair_chart(chart, use_container_width=True)


def render_ratio(title: str, rows: list[dict]):
  st.subheader(title)
  if not rows:
    st.info("No records in the current filter.")
    return
  frame = pd.DataFrame(rows)
  bar_chart(frame, "label", "count")
  st.dataframe(
    frame.assign(Share=frame["share"].map(lambda value: f"{value:.1f}%"))[
      ["label", "count", "Share"]
    ],
    use_container_width=True,
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
  st.altair_chart(chart, use_container_width=True)


def render_pareto(result: dict):
  st.subheader("PROBLEM_Mapping Pareto")
  pareto = pd.DataFrame(result["pareto"])
  if pareto.empty:
    st.info("No problem data in the current filter.")
    return
  table = pareto.copy()
  table["share"] = table["share"].map(lambda value: f"{value:.1f}%")
  table["cumulative"] = table["cumulative"].map(lambda value: f"{value:.1f}%")
  st.dataframe(table, use_container_width=True, hide_index=True)


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

  display_frame = count_frame.astype(str)
  for row_label in display_frame.index:
    for column_label in display_frame.columns:
      cfr_value = cfr_frame.loc[row_label, column_label]
      if count_frame.loc[row_label, column_label] and pd.notna(cfr_value):
        display_frame.loc[row_label, column_label] = (
          f"{count_frame.loc[row_label, column_label]} ({pct(cfr_value)})"
        )

  max_cfr = cfr_frame.max(numeric_only=True).max()
  max_count = count_frame.max(numeric_only=True).max()

  def cell_styles(_: pd.DataFrame) -> pd.DataFrame:
    styles = pd.DataFrame("", index=display_frame.index, columns=display_frame.columns)
    for row_label in display_frame.index:
      for column_label in display_frame.columns:
        count = count_frame.loc[row_label, column_label]
        cfr_value = cfr_frame.loc[row_label, column_label]
        if not count:
          styles.loc[row_label, column_label] = "background-color: #f6fbfb; color: #5f6f72;"
          continue
        denominator = max_cfr if pd.notna(cfr_value) and max_cfr else max_count
        value = cfr_value if pd.notna(cfr_value) and max_cfr else count
        intensity = min(float(value or 0) / float(denominator or 1), 1)
        lightness = int(94 - 42 * intensity)
        styles.loc[row_label, column_label] = (
          f"background-color: hsl(183, 36%, {lightness}%); "
          "color: #071316; font-weight: 650; text-align: center;"
        )
    return styles

  st.dataframe(
    display_frame.style.apply(cell_styles, axis=None),
    use_container_width=True,
    height=360,
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

  metric_row(result)
  st.divider()

  render_trend(result)

  left, right = st.columns(2)
  with left:
    render_ratio("ORG_MODEL(PRODUCT_DESC) Share", result["share"])
  with right:
    render_ratio("MUC_MODULE Ratio", result["module_ratio"])

  left, right = st.columns([1, 1.2])
  with left:
    render_pareto(result)
  with right:
    render_heatmap(result)


if __name__ == "__main__":
  main()
