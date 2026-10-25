-- Microsoft Graph token (our own copy, encrypted) for email + To Do.
CREATE TABLE ms_tokens (
  id INTEGER PRIMARY KEY CHECK (id = 1),
  blob TEXT NOT NULL,                -- AES-GCM encrypted JSON {rt, at, exp}
  scope TEXT,
  email TEXT,
  updated_at TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ', 'now'))
);

-- Action items extracted from summaries, with completion + To Do sync state.
CREATE TABLE action_items (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  meeting_id TEXT NOT NULL,
  position INTEGER NOT NULL,
  task TEXT NOT NULL,
  owner TEXT,
  due_text TEXT,
  due_date TEXT,                     -- YYYY-MM-DD or NULL
  is_mine INTEGER NOT NULL DEFAULT 0,
  done INTEGER NOT NULL DEFAULT 0,
  done_at TEXT,
  todo_task_id TEXT,
  created_at TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ', 'now'))
);
CREATE INDEX action_items_meeting ON action_items (meeting_id);
CREATE INDEX action_items_open ON action_items (done, due_date);

-- Transcript chunks for "Ask your meetings" (vectors live in Vectorize).
CREATE TABLE chunks (
  id TEXT PRIMARY KEY,               -- {meeting_id}:{n}
  meeting_id TEXT NOT NULL,
  n INTEGER NOT NULL,
  start_seconds REAL,
  text TEXT NOT NULL
);
CREATE INDEX chunks_meeting ON chunks (meeting_id);
CREATE VIRTUAL TABLE chunks_fts USING fts5 (
  chunk_id UNINDEXED,
  text,
  tokenize = 'porter unicode61'
);

ALTER TABLE meetings ADD COLUMN recap_sent_at TEXT;
ALTER TABLE meetings ADD COLUMN indexed_at TEXT;
ALTER TABLE meetings ADD COLUMN dex_synced_at TEXT;
ALTER TABLE meetings ADD COLUMN dex_note TEXT;
ALTER TABLE meetings ADD COLUMN audio_key TEXT;
ALTER TABLE meetings ADD COLUMN audio_bytes INTEGER;
ALTER TABLE meetings ADD COLUMN delivery_error TEXT;

ALTER TABLE events ADD COLUMN is_all_day INTEGER NOT NULL DEFAULT 0;
ALTER TABLE events ADD COLUMN bot_config_version INTEGER;
ALTER TABLE events ADD COLUMN brief_sent_at TEXT;
