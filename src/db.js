/**
 * D1 (SQLite) persistence for member-addition requests.
 * Every function takes the D1 binding as its first argument.
 */

export async function createRequest(db, {
  requesterName, requesterAadId, requesterEmail, teamId, teamName, teamsChannelId,
  memberEmail, originalMessage, conversationId, serviceUrl,
}) {
  const { meta } = await db
    .prepare(
      `INSERT INTO requests
        (requester_name, requester_aad_id, requester_email, team_id, team_name, teams_channel_id,
         member_email, original_message, conversation_id, service_url)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
    )
    .bind(requesterName, requesterAadId, requesterEmail ?? null, teamId, teamName, teamsChannelId ?? null,
          memberEmail, originalMessage ?? null,
          conversationId ?? null, serviceUrl ?? null)
    .run();

  return db.prepare('SELECT * FROM requests WHERE id = ?')
    .bind(meta.last_row_id).first();
}

export async function setSlackMessageTs(db, requestId, ts) {
  await db.prepare('UPDATE requests SET slack_message_ts = ? WHERE id = ?')
    .bind(ts, requestId).run();
}

export async function reviewRequest(db, requestId, {
  status, reviewerId, reviewerName, reviewNote,
}) {
  await db.prepare(
    `UPDATE requests
       SET status = ?, reviewer_id = ?, reviewer_name = ?,
           review_note = ?, reviewed_at = datetime('now')
     WHERE id = ?`
  ).bind(status, reviewerId, reviewerName, reviewNote ?? null, requestId).run();

  return db.prepare('SELECT * FROM requests WHERE id = ?')
    .bind(requestId).first();
}

export async function getRequest(db, requestId) {
  return db.prepare('SELECT * FROM requests WHERE id = ?')
    .bind(requestId).first();
}
