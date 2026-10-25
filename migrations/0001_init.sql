-- Connected calendars (Recall Calendar V2). Single user, but keyed by Recall id.
CREATE TABLE calendars (
  id TEXT PRIMARY KEY,               -- Recall calendar id
  platform TEXT NOT NULL,            -- microsoft_outlook
  email TEXT,
  status TEXT,                       -- connected | disconnected | connecting
  created_at TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ', 'now')),
  updated_at TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ', 'now'))
);

-- Calendar events mirrored from Recall, plus our recording decision.
CREATE TABLE events (
  id TEXT PRIMARY KEY,               -- Recall calendar event id
  calendar_id TEXT NOT NULL,
  title TEXT,
  start_time TEXT NOT NULL,
  end_time TEXT NOT NULL,
  meeting_url TEXT,
  meeting_platform TEXT,
  organizer TEXT,
  attendees_json TEXT,
  response TEXT,                     -- accepted | declined | organizer | tentativelyAccepted | none
  is_cancelled INTEGER NOT NULL DEFAULT 0,
  is_deleted INTEGER NOT NULL DEFAULT 0,
  record_override INTEGER,           -- NULL = follow rules, 1 = always, 0 = never
  bot_id TEXT,
  will_record INTEGER NOT NULL DEFAULT 0,
  skip_reason TEXT,
  schedule_error TEXT,
  updated_at TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ', 'now'))
);
CREATE INDEX events_start ON events (start_time);

-- One row per bot (a recorded or attempted meeting).
CREATE TABLE meetings (
  id TEXT PRIMARY KEY,               -- Recall bot id
  event_id TEXT,
  title TEXT,
  start_time TEXT,
  end_time TEXT,
  meeting_url TEXT,
  platform TEXT,
  status TEXT NOT NULL DEFAULT 'scheduled',
  status_detail TEXT,
  recording_id TEXT,
  transcript_id TEXT,
  participants_json TEXT,
  summary_json TEXT,
  duration_seconds INTEGER,
  error TEXT,
  created_at TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ', 'now')),
  updated_at TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ', 'now'))
);
CREATE INDEX meetings_start ON meetings (start_time);

-- Full-text search over titles, summaries and transcripts.
CREATE VIRTUAL TABLE meetings_fts USING fts5 (
  meeting_id UNINDEXED,
  title,
  summary,
  transcript,
  tokenize = 'porter unicode61'
);

CREATE TABLE settings (
  key TEXT PRIMARY KEY,
  value TEXT NOT NULL
);

CREATE TABLE oauth_states (
  state TEXT PRIMARY KEY,
  created_at TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ', 'now'))
);

-- Recent webhook deliveries, for debugging from the Settings page.
CREATE TABLE webhook_log (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  received_at TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ', 'now')),
  event TEXT,
  verified INTEGER NOT NULL,
  note TEXT
);
