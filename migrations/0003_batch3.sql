-- Uploads, editing, sharing, consent, pause, retention and cost tracking.
ALTER TABLE meetings ADD COLUMN source TEXT NOT NULL DEFAULT 'bot';   -- bot | upload
ALTER TABLE meetings ADD COLUMN attendees_json TEXT;                  -- for uploads / ad-hoc bots
ALTER TABLE meetings ADD COLUMN meeting_type_override TEXT;
ALTER TABLE meetings ADD COLUMN edited_at TEXT;
ALTER TABLE meetings ADD COLUMN shared_at TEXT;
ALTER TABLE meetings ADD COLUMN shared_to TEXT;
ALTER TABLE meetings ADD COLUMN consent_declined INTEGER NOT NULL DEFAULT 0;
ALTER TABLE meetings ADD COLUMN paused INTEGER NOT NULL DEFAULT 0;
ALTER TABLE meetings ADD COLUMN bot_joined_at TEXT;
ALTER TABLE meetings ADD COLUMN bot_left_at TEXT;
ALTER TABLE meetings ADD COLUMN lobby_alert_sent INTEGER NOT NULL DEFAULT 0;
ALTER TABLE meetings ADD COLUMN audio_type TEXT;
ALTER TABLE meetings ADD COLUMN audio_purged_at TEXT;
ALTER TABLE meetings ADD COLUMN recall_purged_at TEXT;
ALTER TABLE meetings ADD COLUMN deals_matched_at TEXT;

ALTER TABLE action_items ADD COLUMN dex_reminder_id TEXT;
ALTER TABLE action_items ADD COLUMN nudged_at TEXT;

-- Remembered speaker names ("iPhone" -> "Jane Doe").
CREATE TABLE speaker_aliases (
  raw TEXT PRIMARY KEY,
  name TEXT NOT NULL
);

-- Meetings linked to deals in the deal-pipeline app (read-only there).
CREATE TABLE meeting_deals (
  meeting_id TEXT NOT NULL,
  deal_id TEXT NOT NULL,
  deal_name TEXT,
  source TEXT NOT NULL DEFAULT 'auto',   -- auto | manual
  PRIMARY KEY (meeting_id, deal_id)
);
CREATE INDEX meeting_deals_deal ON meeting_deals (deal_id);

-- Estimated costs (OpenAI tokens/minutes; Recall hours are derived from meetings).
CREATE TABLE usage (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  at TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ', 'now')),
  purpose TEXT NOT NULL,
  model TEXT,
  input_tokens INTEGER,
  output_tokens INTEGER,
  minutes REAL,
  cost REAL NOT NULL DEFAULT 0
);
CREATE INDEX usage_at ON usage (at);
