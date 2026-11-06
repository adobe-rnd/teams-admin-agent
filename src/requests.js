/**
 * Submitting member requests: the one path both triggers take.
 *
 * - An @admin-bot mention in Teams (src/teams.js), delivered by Azure Bot Service.
 * - A call from another Worker of the account through a service binding (src/rpc.js), such as
 *   opspile's thread agents. Teams never delivers a bot's message to another bot, so a bot
 *   cannot ask by mentioning @admin-bot; it calls `MembershipRequests.submit` instead.
 *
 * Either way, every request is a D1 row and an approval card in SLACK_ADMIN_CHANNEL_ID, and
 * nothing changes in Teams until an admin approves it there.
 */
import { createRequest } from './db.js';
import { postApprovalCard } from './slack.js';
import { getTeamMemberEmails } from './graph.js';

/**
 * Creates a request per email that needs one, and posts its approval card. People already in
 * the team (for an add) or not in it (for a remove) are skipped. Returns
 * `{ created, skipped, failed }`: request rows, emails, and `{ email, error }`.
 */
export async function submitRequests(env, {
  action, emails, teamId, teamName, teamsChannelId, requesterName, requesterAadId,
  requesterEmail, originalMessage, conversationId, serviceUrl, via,
}, { memberEmails = getTeamMemberEmails, approvalCard = postApprovalCard } = {}) {
  let members;
  try {
    members = await memberEmails(env, teamId);
  } catch (err) {
    console.error('getTeamMemberEmails failed:', err);
    members = new Set();
  }
  const isRemove = action === 'remove';
  const toProcess = [];
  const skipped = [];
  for (const email of emails) {
    const isMember = members.has(email.toLowerCase().trim());
    // remove → process members, skip non-members; add → the reverse.
    if (isMember === isRemove) {
      toProcess.push(email);
    } else {
      skipped.push(email);
    }
  }

  const created = [];
  const failed = [];
  for (const email of toProcess) {
    try {
      const req = await createRequest(env.DB, {
        requesterName, requesterAadId, requesterEmail, teamId, teamName, teamsChannelId,
        memberEmail: email, originalMessage, conversationId, serviceUrl, action, via,
      });
      await approvalCard(env, req);
      created.push(req);
    } catch (err) {
      console.error(`Request failed for ${email}:`, err);
      failed.push({ email, error: err.message });
    }
  }
  return { created, skipped, failed };
}

/** The reply in Teams to submitted requests, or '' when there is nothing to say. */
export function replyText(action, { created, skipped, failed }) {
  const isRemove = action === 'remove';
  const lines = [];
  if (skipped.length > 0) {
    const one = skipped.length === 1;
    const skippedText = isRemove
      ? (one ? 'The following user is not a member of this team:' : 'The following users are not members of this team:')
      : (one ? 'The following user is already a member of this team:' : 'The following users are already members of this team:');
    lines.push(skippedText, '');
    skipped.forEach((e, i) => {
      lines.push(`• ${e}`);
      if (i < skipped.length - 1) lines.push('');
    });
    if (created.length > 0 || failed.length > 0) lines.push('', '');
  }
  if (created.length) {
    const n = created.length;
    const requestWord = n === 1 ? 'request' : 'requests';
    lines.push(`${isRemove ? 'Removal' : 'Invitation'} ${requestWord} submitted for approval:`, '');
    created.forEach((r, i) => {
      lines.push(`• ${r.member_email}`);
      if (i < created.length - 1) lines.push('');
    });
    lines.push('', '', `You'll be notified here once the ${requestWord} ${n === 1 ? 'has' : 'have'} been reviewed.`);
  }
  if (failed.length) {
    lines.push('', `**${failed.length} failed to submit:**`);
    failed.forEach((f) => lines.push(`- \`${f.email}\`: ${f.error}`));
  }
  return lines.join('\r\n');
}
