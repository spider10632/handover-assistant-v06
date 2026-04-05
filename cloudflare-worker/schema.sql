CREATE TABLE IF NOT EXISTS handover_state (
  server_id TEXT PRIMARY KEY,
  payload TEXT NOT NULL,
  updated_at TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_handover_state_updated_at ON handover_state(updated_at);
