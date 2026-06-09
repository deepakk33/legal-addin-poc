-- Append-only legal audit trail. SQLite for the POC; schema ports straight to Postgres.
-- This records the who/what/which-model trail that Word's own version history cannot.
CREATE TABLE IF NOT EXISTS edit_log (
  id            INTEGER PRIMARY KEY AUTOINCREMENT,
  created_at    TEXT    NOT NULL,   -- ISO timestamp
  doc_name      TEXT,
  instruction   TEXT    NOT NULL,
  model_name    TEXT    NOT NULL,
  model_version TEXT    NOT NULL,
  original_text TEXT    NOT NULL,
  edited_text   TEXT    NOT NULL,
  status        TEXT    NOT NULL DEFAULT 'pending'  -- pending | accepted | rejected
);
