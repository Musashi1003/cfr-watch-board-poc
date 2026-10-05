from __future__ import annotations

import argparse
import os
from pathlib import Path
import re
import sys

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from act_postgres import persist_act_updates


def normalize_source_type(value) -> str:
  text = str(value or "").upper().replace("_", "-").replace(" ", "-")
  if "GAMING" in text:
    return "Gaming NB"
  if "PC" in text:
    return "PC NB"
  return str(value or "").strip()


def workbook_updates(path: Path) -> list[dict]:
  updates: list[dict] = []
  workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
  try:
    for worksheet in workbook.worksheets:
      rows = worksheet.iter_rows(values_only=True)
      headers = [str(value or "").strip() for value in next(rows)]
      normalized = {re.sub(r"[^A-Z0-9]", "", value.upper()): index for index, value in enumerate(headers)}
      source_index = normalized.get("GAMINGPC")
      model_index = normalized.get("ORGMODELPRODUCTDESC", normalized.get("MODELGROUP"))
      if source_index is None or model_index is None:
        continue
      week_columns = [
        (index, value.upper())
        for index, value in enumerate(headers)
        if re.fullmatch(r"W\d{4}", value, flags=re.IGNORECASE)
      ]
      title_match = re.search(r"(20\d{2})", worksheet.title)
      sheet_year = title_match.group(1) if title_match else ""
      for row in rows:
        launch_year = str(row[0] or sheet_year).strip()
        if launch_year.endswith(".0"):
          launch_year = launch_year[:-2]
        source_type = normalize_source_type(row[source_index])
        model = str(row[model_index] or "").strip()
        if not launch_year or not source_type or not model:
          continue
        for column_index, week in week_columns:
          value = row[column_index] if column_index < len(row) else None
          if value in (None, ""):
            continue
          updates.append(
            {
              "source_type": source_type,
              "launch_year": launch_year,
              "model": model,
              "week": week,
              "cumulative_activation": value,
              "source_files": [path.name],
            }
          )
  finally:
    workbook.close()
  return updates


def main() -> int:
  parser = argparse.ArgumentParser(description="Import ACT table.xlsx into PostgreSQL.")
  parser.add_argument("workbook", type=Path)
  parser.add_argument("--database-url-env", default="ACT_DATABASE_URL")
  args = parser.parse_args()

  database_url = os.environ.get(args.database_url_env, "").strip()
  if not database_url:
    parser.error(f"Environment variable {args.database_url_env} is not configured.")
  if not args.workbook.exists():
    parser.error(f"Workbook not found: {args.workbook}")

  updates = workbook_updates(args.workbook)
  result = persist_act_updates(
    database_url,
    updates,
    [(args.workbook.name, args.workbook.read_bytes())],
    actor="ACT table bootstrap",
    allow_initial_snapshot=True,
  )
  print(
    f"{result['status']}: {result['week_label']}; "
    f"updates={result['generated_count']}; verified={result['verified_count']}; "
    f"batch={result['batch_id']}"
  )
  return 0


if __name__ == "__main__":
  raise SystemExit(main())
