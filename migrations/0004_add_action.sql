-- Distinguish add (invite) from remove (destructive) requests.
-- Existing rows predate removal support, so they default to 'add'.
ALTER TABLE requests ADD COLUMN action TEXT NOT NULL DEFAULT 'add';
