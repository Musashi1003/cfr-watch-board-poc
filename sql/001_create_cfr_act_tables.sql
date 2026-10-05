BEGIN;

CREATE TABLE IF NOT EXISTS cfr_act_upload_batches (
  batch_id uuid PRIMARY KEY,
  status text NOT NULL CHECK (status IN ('pending', 'committed', 'failed')),
  expected_update_count integer NOT NULL CHECK (expected_update_count >= 0),
  verified_snapshot_count integer,
  payload_checksum text NOT NULL,
  verified_checksum text,
  source_files jsonb NOT NULL DEFAULT '[]'::jsonb,
  created_by text NOT NULL,
  created_at timestamptz NOT NULL DEFAULT now(),
  committed_at timestamptz,
  error_message text
);

CREATE TABLE IF NOT EXISTS cfr_act_qty (
  source_type text NOT NULL,
  launch_year smallint NOT NULL CHECK (launch_year BETWEEN 2000 AND 2100),
  model text NOT NULL,
  week text NOT NULL CHECK (week ~ '^W[0-9]{4}$'),
  week_code integer NOT NULL CHECK (week_code BETWEEN 0 AND 9999),
  act_qty bigint NOT NULL CHECK (act_qty >= 0),
  source_file text,
  source_sha256 text,
  upload_batch_id uuid NOT NULL REFERENCES cfr_act_upload_batches(batch_id),
  carried_forward boolean NOT NULL DEFAULT false,
  updated_by text NOT NULL,
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now(),
  PRIMARY KEY (source_type, launch_year, model, week),
  CHECK (week_code = substring(week from 2)::integer)
);

CREATE INDEX IF NOT EXISTS idx_cfr_act_qty_scope_week
  ON cfr_act_qty (source_type, launch_year, week_code DESC);

CREATE INDEX IF NOT EXISTS idx_cfr_act_qty_model_week
  ON cfr_act_qty (launch_year, model, week_code DESC);

COMMIT;
