-- Dropbox: deal folders (owned, active, passed, sold) to link meetings to.
CREATE TABLE dropbox_tokens (
  id INTEGER PRIMARY KEY CHECK (id = 1),
  blob TEXT NOT NULL,          -- AES-GCM encrypted refresh token
  account TEXT,
  scope TEXT,
  updated_at TEXT NOT NULL
);

-- id is Dropbox's file id, which survives moves (Active -> Passed keeps links).
CREATE TABLE dropbox_folders (
  id TEXT PRIMARY KEY,
  name TEXT NOT NULL,
  status TEXT NOT NULL,        -- owned | active | passed | sold
  path TEXT NOT NULL,          -- display path, e.g. /10 - Active Deals/Desert Shops (AZ)
  grp TEXT,                    -- e.g. "02 - Arizona" for passed deals
  shared INTEGER NOT NULL DEFAULT 0, -- a shared folder: others can see what's saved there
  modified_at TEXT,
  seen_at TEXT NOT NULL
);

-- Meeting notes saved into deal folders (re-saved when the notes change).
CREATE TABLE dropbox_saves (
  meeting_id TEXT NOT NULL,
  deal_id TEXT NOT NULL,
  path TEXT NOT NULL,
  hash TEXT NOT NULL,
  explicit INTEGER NOT NULL DEFAULT 0, -- you chose this folder with "Save notes to Dropbox"
  saved_at TEXT NOT NULL,
  PRIMARY KEY (meeting_id, deal_id)
);
