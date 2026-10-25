-- Deal Snapshots: a short extract of "<deal> — Deal Snapshot.md" (broker, asking price, size)
-- that helps link calls to the right deal. Re-read only when the file's revision changes.
ALTER TABLE dropbox_folders ADD COLUMN snapshot TEXT;
ALTER TABLE dropbox_folders ADD COLUMN snapshot_rev TEXT;

-- Market comps quoted on calls (stored in the summary), with your review decision.
CREATE TABLE comp_reviews (
  meeting_id TEXT NOT NULL,
  idx INTEGER NOT NULL,          -- position in summary.comps
  status TEXT NOT NULL,          -- approved | rejected
  reviewed_at TEXT NOT NULL,
  PRIMARY KEY (meeting_id, idx)
);

-- Nightly Markdown copy of each meeting in Dropbox (/_Automation/Meeting Archive/Main).
CREATE TABLE meeting_archive (
  meeting_id TEXT PRIMARY KEY,
  path TEXT NOT NULL,
  hash TEXT NOT NULL,
  saved_at TEXT NOT NULL
);
