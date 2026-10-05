from __future__ import annotations

from collections import defaultdict
import hashlib
import json
import math
import re
from typing import Iterable
from uuid import uuid4


ACT_TABLE = "cfr_act_qty"
BATCH_TABLE = "cfr_act_upload_batches"
WEEK_PATTERN = re.compile(r"W(\d{2})(\d{2})", flags=re.IGNORECASE)
ADVISORY_LOCK_ID = 2026100501


class ActDatabaseError(RuntimeError):
  """Raised when ACT persistence cannot be safely completed."""


class ActConflictError(ActDatabaseError):
  """Raised when an existing ACT value conflicts with an uploaded value."""


class ActContinuityError(ActDatabaseError):
  """Raised when a weekly ACT snapshot would leave a gap."""


def week_code(week: str) -> int:
  match = WEEK_PATTERN.fullmatch(str(week or "").strip())
  if not match:
    raise ValueError(f"Invalid ACT week: {week!r}")
  return int(match.group(1)) * 100 + int(match.group(2))


def required_previous_week_code(target_week_code: int) -> int | None:
  week_number = target_week_code % 100
  if week_number <= 1:
    return None
  return target_week_code - 1


def _database_driver():
  try:
    import psycopg
  except ImportError as exc:
    raise ActDatabaseError(
      "PostgreSQL driver is unavailable. Install psycopg[binary] from requirements.txt."
    ) from exc
  return psycopg


def _connect(database_url: str):
  if not str(database_url or "").strip():
    raise ActDatabaseError("ACT_DATABASE_URL is not configured.")
  psycopg = _database_driver()
  try:
    return psycopg.connect(database_url, connect_timeout=10)
  except Exception as exc:
    raise ActDatabaseError(f"PostgreSQL connection failed: {exc}") from exc


def check_database(database_url: str) -> tuple[bool, str]:
  try:
    with _connect(database_url) as connection:
      with connection.cursor() as cursor:
        cursor.execute(
          "SELECT current_database(), to_regclass(%s), to_regclass(%s)",
          (f"public.{ACT_TABLE}", f"public.{BATCH_TABLE}"),
        )
        database_name, act_table, batch_table = cursor.fetchone()
        if act_table is None or batch_table is None:
          return False, "Connected, but the CFR ACT schema has not been installed."
        cursor.execute(f"SELECT COUNT(*) FROM {ACT_TABLE}")
        row_count = int(cursor.fetchone()[0])
        if row_count == 0:
          return False, "Connected and schema is ready, but the ACT baseline has not been imported."
        return True, f"Connected to PostgreSQL database {database_name}; ACT schema is ready."
  except ActDatabaseError as exc:
    return False, str(exc)
  except Exception as exc:
    return False, f"PostgreSQL health check failed: {exc}"


def load_act_store(database_url: str) -> dict[tuple[str, str, str, str], float]:
  try:
    with _connect(database_url) as connection:
      with connection.cursor() as cursor:
        cursor.execute(
          f"""
          SELECT source_type, launch_year, model, week, act_qty
          FROM {ACT_TABLE}
          ORDER BY source_type, launch_year, model, week_code
          """
        )
        return {
          (str(source_type), str(launch_year), str(model), str(week)): float(act_qty)
          for source_type, launch_year, model, week, act_qty in cursor.fetchall()
        }
  except ActDatabaseError:
    raise
  except Exception as exc:
    raise ActDatabaseError(f"PostgreSQL ACT read failed: {exc}") from exc


def upload_manifest(upload_payloads: Iterable[tuple[str, bytes]]) -> list[dict[str, object]]:
  return [
    {
      "filename": str(filename),
      "size_bytes": len(content),
      "sha256": hashlib.sha256(content).hexdigest(),
    }
    for filename, content in upload_payloads
  ]


def _normalized_updates(updates: Iterable[dict]) -> list[dict]:
  normalized_by_key: dict[tuple[str, str, str, str], dict] = {}
  for raw in updates:
    source_type = str(raw.get("source_type", "")).strip()
    launch_year = str(raw.get("launch_year", "")).strip()
    model = str(raw.get("model", "")).strip()
    week = str(raw.get("week", "")).strip().upper()
    if not source_type or not re.fullmatch(r"\d{4}", launch_year) or not model:
      raise ActDatabaseError("ACT update is missing source type, launch year, or model.")
    try:
      code = week_code(week)
      raw_qty = float(raw.get("cumulative_activation"))
    except (TypeError, ValueError) as exc:
      raise ActDatabaseError(f"ACT update has an invalid week or quantity for {model}.") from exc
    if not math.isfinite(raw_qty) or raw_qty < 0:
      raise ActDatabaseError(f"ACT update has a negative or non-finite quantity for {model}.")

    source_files_value = raw.get("source_files") or []
    if isinstance(source_files_value, str):
      source_files = [source_files_value]
    else:
      source_files = [str(value).strip() for value in source_files_value if str(value).strip()]

    key = (source_type, launch_year, model, week)
    normalized = {
      "source_type": source_type,
      "launch_year": int(launch_year),
      "model": model,
      "week": week,
      "week_code": code,
      "act_qty": int(round(raw_qty)),
      "source_files": sorted(set(source_files)),
    }
    existing = normalized_by_key.get(key)
    if existing and existing["act_qty"] != normalized["act_qty"]:
      raise ActConflictError(f"Upload contains conflicting ACT values for {key}.")
    normalized_by_key[key] = normalized

  return sorted(
    normalized_by_key.values(),
    key=lambda row: (row["week_code"], row["source_type"], row["launch_year"], row["model"]),
  )


def rows_checksum(rows: Iterable[dict]) -> str:
  canonical_rows = [
    {
      "source_type": str(row["source_type"]),
      "launch_year": int(row["launch_year"]),
      "model": str(row["model"]),
      "week": str(row["week"]),
      "act_qty": int(row["act_qty"]),
    }
    for row in rows
  ]
  canonical_rows.sort(
    key=lambda row: (
      row["source_type"],
      row["launch_year"],
      week_code(row["week"]),
      row["model"],
    )
  )
  payload = json.dumps(canonical_rows, ensure_ascii=False, separators=(",", ":"))
  return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def _source_metadata(update: dict, manifest_by_name: dict[str, dict]) -> tuple[str, str]:
  names = update.get("source_files") or []
  matching = [manifest_by_name[name] for name in names if name in manifest_by_name]
  if not matching:
    return "", ""
  filename = "; ".join(str(item["filename"]) for item in matching)
  digest_source = "|".join(str(item["sha256"]) for item in matching)
  return filename, hashlib.sha256(digest_source.encode("ascii")).hexdigest()


def persist_act_updates(
  database_url: str,
  updates: Iterable[dict],
  upload_payloads: Iterable[tuple[str, bytes]] = (),
  actor: str = "IEC-CFR",
  allow_initial_snapshot: bool = False,
) -> dict:
  normalized = _normalized_updates(updates)
  if not normalized:
    return {
      "status": "skipped",
      "message": "No ACT values were found to record.",
      "persistence_ok": True,
      "backend": "postgresql",
    }

  manifest = upload_manifest(upload_payloads)
  manifest_by_name = {str(item["filename"]): item for item in manifest}
  payload_checksum = rows_checksum(normalized)
  batch_id = str(uuid4())
  actor = str(actor or "IEC-CFR").strip()[:128] or "IEC-CFR"
  grouped: dict[tuple[str, int, str, int], list[dict]] = defaultdict(list)
  for update in normalized:
    grouped[
      (
        update["source_type"],
        update["launch_year"],
        update["week"],
        update["week_code"],
      )
    ].append(update)

  created_scope_count = 0
  applied_update_count = 0
  carried_count = 0
  verified_rows: list[dict] = []
  try:
    with _connect(database_url) as connection:
      with connection.transaction():
        with connection.cursor() as cursor:
          cursor.execute("SET LOCAL lock_timeout = '10s'")
          cursor.execute("SET LOCAL statement_timeout = '60s'")
          cursor.execute("SELECT pg_advisory_xact_lock(%s)", (ADVISORY_LOCK_ID,))
          cursor.execute(
            f"""
            INSERT INTO {BATCH_TABLE} (
              batch_id, status, expected_update_count, payload_checksum,
              source_files, created_by
            ) VALUES (%s, 'pending', %s, %s, %s::jsonb, %s)
            """,
            (batch_id, len(normalized), payload_checksum, json.dumps(manifest), actor),
          )

          for (source_type, launch_year, week, target_code), scope_updates in sorted(
            grouped.items(), key=lambda item: (item[0][3], item[0][0], item[0][1])
          ):
            cursor.execute(
              f"""
              SELECT model, act_qty
              FROM {ACT_TABLE}
              WHERE source_type = %s AND launch_year = %s AND week = %s
              ORDER BY model
              """,
              (source_type, launch_year, week),
            )
            existing_target = {str(model): int(qty) for model, qty in cursor.fetchall()}

            cursor.execute(
              f"""
              SELECT week, week_code
              FROM {ACT_TABLE}
              WHERE source_type = %s AND launch_year = %s AND week_code < %s
              ORDER BY week_code DESC
              LIMIT 1
              """,
              (source_type, launch_year, target_code),
            )
            prior_week_row = cursor.fetchone()
            previous_snapshot: dict[str, int] = {}
            prior_code_value: int | None = None
            if prior_week_row:
              prior_week, prior_code = str(prior_week_row[0]), int(prior_week_row[1])
              prior_code_value = prior_code
              required_code = required_previous_week_code(target_code)
              if required_code is not None and prior_code != required_code:
                raise ActContinuityError(
                  f"Cannot save {week}: the latest PostgreSQL ACT week for "
                  f"{source_type} {launch_year} is {prior_week}."
                )
              cursor.execute(
                f"""
                SELECT model, act_qty
                FROM {ACT_TABLE}
                WHERE source_type = %s AND launch_year = %s AND week = %s
                ORDER BY model
                """,
                (source_type, launch_year, prior_week),
              )
              previous_snapshot = {str(model): int(qty) for model, qty in cursor.fetchall()}

            if not existing_target and not previous_snapshot and not allow_initial_snapshot:
              raise ActContinuityError(
                f"Cannot save {source_type} {week}: PostgreSQL has no prior ACT baseline for "
                f"launch year {launch_year}. Run the ACT table bootstrap import first."
              )

            uploaded_values = {str(row["model"]): int(row["act_qty"]) for row in scope_updates}
            if existing_target:
              if previous_snapshot and not set(previous_snapshot).issubset(existing_target):
                raise ActDatabaseError(
                  f"PostgreSQL snapshot {source_type} {week} is incomplete; manual review is required."
                )
              conflicts = {
                model: (existing_target.get(model), qty)
                for model, qty in uploaded_values.items()
                if existing_target.get(model) != qty
              }
              if conflicts:
                model, values = next(iter(conflicts.items()))
                raise ActConflictError(
                  f"PostgreSQL already has a different {week} ACT value for {model}: "
                  f"stored {values[0]}, uploaded {values[1]}."
                )
              actual_snapshot = existing_target
            else:
              created_scope_count += 1
              applied_update_count += len(scope_updates)
              if previous_snapshot:
                cursor.execute(
                  f"""
                  INSERT INTO {ACT_TABLE} (
                    source_type, launch_year, model, week, week_code, act_qty,
                    upload_batch_id, carried_forward, updated_by
                  )
                  SELECT source_type, launch_year, model, %s, %s, act_qty,
                         %s, TRUE, %s
                  FROM {ACT_TABLE}
                  WHERE source_type = %s AND launch_year = %s AND week_code = %s
                  """,
                  (
                    week,
                    target_code,
                    batch_id,
                    actor,
                    source_type,
                    launch_year,
                    prior_code_value,
                  ),
                )
                carried_count += len(set(previous_snapshot) - set(uploaded_values))

              for update in scope_updates:
                source_file, source_sha256 = _source_metadata(update, manifest_by_name)
                cursor.execute(
                  f"""
                  INSERT INTO {ACT_TABLE} (
                    source_type, launch_year, model, week, week_code, act_qty,
                    source_file, source_sha256, upload_batch_id,
                    carried_forward, updated_by, updated_at
                  ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, FALSE, %s, NOW())
                  ON CONFLICT (source_type, launch_year, model, week)
                  DO UPDATE SET
                    act_qty = EXCLUDED.act_qty,
                    source_file = EXCLUDED.source_file,
                    source_sha256 = EXCLUDED.source_sha256,
                    upload_batch_id = EXCLUDED.upload_batch_id,
                    carried_forward = FALSE,
                    updated_by = EXCLUDED.updated_by,
                    updated_at = NOW()
                  """,
                  (
                    source_type,
                    launch_year,
                    update["model"],
                    week,
                    target_code,
                    update["act_qty"],
                    source_file,
                    source_sha256,
                    batch_id,
                    actor,
                  ),
                )

              cursor.execute(
                f"""
                SELECT model, act_qty
                FROM {ACT_TABLE}
                WHERE source_type = %s AND launch_year = %s AND week = %s
                ORDER BY model
                """,
                (source_type, launch_year, week),
              )
              actual_snapshot = {str(model): int(qty) for model, qty in cursor.fetchall()}

              expected_snapshot = dict(previous_snapshot)
              expected_snapshot.update(uploaded_values)
              if actual_snapshot != expected_snapshot:
                raise ActDatabaseError(
                  f"PostgreSQL read-back verification failed for {source_type} {week}; "
                  "the transaction was rolled back."
                )

            verified_rows.extend(
              {
                "source_type": source_type,
                "launch_year": launch_year,
                "model": model,
                "week": week,
                "act_qty": qty,
              }
              for model, qty in actual_snapshot.items()
            )

          verified_checksum = rows_checksum(verified_rows)
          cursor.execute(
            f"""
            UPDATE {BATCH_TABLE}
            SET status = 'committed',
                verified_snapshot_count = %s,
                verified_checksum = %s,
                committed_at = NOW()
            WHERE batch_id = %s
            """,
            (len(verified_rows), verified_checksum, batch_id),
          )
  except (ActDatabaseError, ActConflictError, ActContinuityError):
    raise
  except Exception as exc:
    raise ActDatabaseError(f"PostgreSQL ACT transaction failed and was rolled back: {exc}") from exc

  weeks = sorted({row["week"] for row in normalized}, key=week_code)
  week_label = weeks[0] if len(weeks) == 1 else f"{weeks[0]}-{weeks[-1]}"
  status = "saved" if created_scope_count else "unchanged"
  return {
    "status": status,
    "message": (
      f"PostgreSQL transaction committed and read-back verified for batch {batch_id}."
      if status == "saved"
      else f"PostgreSQL already contained matching values; read-back verified for batch {batch_id}."
    ),
    "backend": "postgresql",
    "batch_id": batch_id,
    "added_count": applied_update_count,
    "kept_count": carried_count,
    "generated_count": len(normalized),
    "verified_count": len(verified_rows),
    "verified_checksum": rows_checksum(verified_rows),
    "week_label": week_label,
    "persistence_ok": True,
    "requires_manual_save": False,
  }
