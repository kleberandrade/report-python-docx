CREATE TABLE IF NOT EXISTS class_evaluations (
  id uuid PRIMARY KEY,
  stage text NOT NULL CHECK (stage IN ('alpha', 'beta')),
  payload jsonb NOT NULL CHECK (jsonb_typeof(payload) = 'object'),
  created_at timestamptz NOT NULL DEFAULT now()
);
CREATE INDEX IF NOT EXISTS class_evaluations_stage_created
  ON class_evaluations(stage, created_at, id);
