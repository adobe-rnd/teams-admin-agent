CREATE TABLE IF NOT EXISTS requests (
  id                INTEGER PRIMARY KEY AUTOINCREMENT,
  requester_name    TEXT NOT NULL,
  requester_aad_id  TEXT NOT NULL,
  team_id           TEXT NOT NULL,
  team_name         TEXT NOT NULL,
  member_email      TEXT NOT NULL,
  original_message  TEXT,
  conversation_id   TEXT,
  service_url       TEXT,
  status            TEXT NOT NULL DEFAULT 'pending',
  reviewer_id       TEXT,
  reviewer_name     TEXT,
  review_note       TEXT,
  slack_message_ts  TEXT,
  created_at        TEXT NOT NULL DEFAULT (datetime('now')),
  reviewed_at       TEXT
);
