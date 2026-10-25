-- Which business a meeting belongs to: main | second | personal (kept strictly separate).
-- entity is computed from attendees and deal links; entity_override is your choice on the meeting page.
ALTER TABLE meetings ADD COLUMN entity TEXT;
ALTER TABLE meetings ADD COLUMN entity_override TEXT;

-- Call activity posted to the deal-pipeline app (one per meeting and deal).
CREATE TABLE pipeline_posts (
  meeting_id TEXT NOT NULL,
  deal_id TEXT NOT NULL,
  activity_id INTEGER,           -- row in the pipeline's activity table
  note_id INTEGER,               -- row in its notes table (summary mode only)
  hash TEXT,
  posted_at TEXT NOT NULL,
  PRIMARY KEY (meeting_id, deal_id)
);

-- "Looks done" suggestions: a sent email that seems to fulfil one of your action items.
CREATE TABLE item_suggestions (
  item_id INTEGER NOT NULL,
  message_id TEXT NOT NULL,
  subject TEXT,
  to_text TEXT,
  web_link TEXT,
  sent_at TEXT,
  reason TEXT,
  status TEXT NOT NULL DEFAULT 'pending', -- pending | accepted | dismissed
  created_at TEXT NOT NULL,
  PRIMARY KEY (item_id, message_id)
);

-- Questions you ask often; answers are refreshed for the Monday review.
CREATE TABLE saved_questions (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  question TEXT NOT NULL,
  answer TEXT,
  sources_json TEXT,
  answered_at TEXT,
  created_at TEXT NOT NULL
);
