-- The Teams message an @admin-bot request came from (the activity id), so an edit of that
-- message does not file the same request again. NULL for service-binding requests.
ALTER TABLE requests ADD COLUMN teams_message_id TEXT;
