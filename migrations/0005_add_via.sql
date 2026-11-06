-- Who submitted a request on the requester's behalf, when it did not come from an @admin-bot
-- mention: another Worker's service binding (src/rpc.js), such as "opspile". NULL for mentions.
ALTER TABLE requests ADD COLUMN via TEXT;
