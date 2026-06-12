from __future__ import annotations

from pathlib import Path
from tempfile import NamedTemporaryFile

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
  cols = st.columns(5)
  cols[0].metric("Filtered CFR", pct(kpis["filtered_cfr"]))
  cols[0].caption(f"{whole(kpis['filtered_failure_qty'])} failures / {whole(kpis['derived_act'])} derived ACT")
  cols[1].metric("Failure Qty", whole(kpis["filtered_failure_qty"]))
  cols[1].caption("Filtered raw-data failures")
  cols[2].metric("Derived ACT", whole(kpis["derived_act"]))
  cols[2].caption("From SUMMARY_IEC CFR rule")
  cols[3].metric("ACT Model Coverage", whole(kpis["act_model_count"]))
  cols[3].caption(f"Target CFR: {pct(kpis['target_cfr'])}")
  cols[4].metric("Latest WoW", f"{kpis['week_delta']:+,}")
  cols[4].caption(f"{kpis['previous_week']} to {kpis['latest_week']}")


def render_ratio(title: str, rows: list[dict]):
  st.subheader(title)
  if not rows:
    st.info("No records in the current filter.")
    return
  frame = pd.DataFrame(rows)
  st.bar_chart(frame.set_index("label")["count"])
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
  st.line_chart(trend.set_index("week")["count"])


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
  st.subheader("ORG_MODEL(PRODUCT_DESC) by PROBLEM_Mapping")
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

  st.dataframe(
    display_frame,
    use_container_width=True,
  )


def render_file_summary(parsed: dict):
  files = parsed.get("files", [])
  if not files:
    return
  with st.expander("Loaded weekly raw data", expanded=True):
    st.dataframe(pd.DataFrame(files), use_container_width=True, hide_index=True)


def main():
  st.title("CFR Watch Board")
  st.caption("Upload weekly Gaming NB and PC NB CFR workbooks, then filter by model, segment, ODM/OEM, module, and problem.")

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

  left, right = st.columns([1.1, 1])
  with left:
    render_trend(result)
  with right:
    render_ratio("Gaming / PC Mix", result["source_mix"])

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
