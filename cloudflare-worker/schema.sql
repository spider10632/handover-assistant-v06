CREATE TABLE IF NOT EXISTS handover_state (
  server_id TEXT PRIMARY KEY,
  payload TEXT NOT NULL,
  updated_at TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_handover_state_updated_at ON handover_state(updated_at);

CREATE TABLE IF NOT EXISTS auth_accounts (
  server_id TEXT NOT NULL,
  username TEXT NOT NULL,
  role TEXT NOT NULL,
  password_hash TEXT NOT NULL,
  enabled INTEGER NOT NULL DEFAULT 1,
  created_at TEXT NOT NULL,
  updated_at TEXT NOT NULL,
  created_by TEXT,
  updated_by TEXT,
  PRIMARY KEY (server_id, username)
);

CREATE INDEX IF NOT EXISTS idx_auth_accounts_server ON auth_accounts(server_id);
CREATE INDEX IF NOT EXISTS idx_auth_accounts_enabled ON auth_accounts(server_id, enabled);

CREATE TABLE IF NOT EXISTS auth_sessions (
  token_hash TEXT PRIMARY KEY,
  server_id TEXT NOT NULL,
  username TEXT NOT NULL,
  role TEXT NOT NULL,
  created_at TEXT NOT NULL,
  expires_at TEXT NOT NULL,
  revoked_at TEXT
);

CREATE INDEX IF NOT EXISTS idx_auth_sessions_server ON auth_sessions(server_id);
CREATE INDEX IF NOT EXISTS idx_auth_sessions_user ON auth_sessions(server_id, username);
CREATE INDEX IF NOT EXISTS idx_auth_sessions_expires ON auth_sessions(expires_at);

CREATE TABLE IF NOT EXISTS audit_logs (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  server_id TEXT NOT NULL,
  actor_username TEXT NOT NULL,
  actor_role TEXT NOT NULL,
  action TEXT NOT NULL,
  target_type TEXT NOT NULL,
  target_id TEXT,
  summary TEXT,
  details TEXT,
  created_at TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_audit_logs_server_created ON audit_logs(server_id, created_at DESC);
CREATE INDEX IF NOT EXISTS idx_audit_logs_actor ON audit_logs(server_id, actor_username, created_at DESC);

CREATE TABLE IF NOT EXISTS push_subscriptions (
  endpoint TEXT PRIMARY KEY,
  server_id TEXT NOT NULL,
  username TEXT NOT NULL,
  p256dh TEXT,
  auth TEXT,
  user_agent TEXT,
  active INTEGER NOT NULL DEFAULT 1,
  created_at TEXT NOT NULL,
  updated_at TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_push_subscriptions_server_user ON push_subscriptions(server_id, username, active);

CREATE TABLE IF NOT EXISTS push_events (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  endpoint TEXT NOT NULL,
  title TEXT NOT NULL,
  body TEXT,
  click_url TEXT,
  created_at TEXT NOT NULL,
  delivered_at TEXT
);

CREATE INDEX IF NOT EXISTS idx_push_events_endpoint_pending ON push_events(endpoint, delivered_at, id);
