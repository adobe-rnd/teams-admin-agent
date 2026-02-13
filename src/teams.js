/**
 * Handles incoming Bot Framework activities from Microsoft Teams.
 *
 * Validates the JWT, parses @admin messages for emails, creates DB
 * records, posts Slack approval cards, and replies in Teams.
 */
import { createRemoteJWKSet, jwtVerify } from 'jose';
import { extractEmails } from './parser.js';
import { createRequest } from './db.js';
import { postApprovalCard } from './slack.js';
import { getTeamName } from './graph.js';

const BF_JWKS = createRemoteJWKSet(
  new URL('https://login.botframework.com/v1/.well-known/keys'),
);

// ── Entry point ─────────────────────────────────────────────────

export async function handleTeamsActivity(request, env, ctx) {
  // Validate Bot Framework bearer token
  const auth = request.headers.get('Authorization') ?? '';
  if (!auth.startsWith('Bearer ')) {
    return new Response('Unauthorized', { status: 401 });
  }
  try {
    await jwtVerify(auth.slice(7), BF_JWKS, {
      audience: env.BOT_ID,
      issuer: 'https://api.botframework.com',
      clockTolerance: 300,
    });
  } catch (err) {
    console.error('JWT validation failed:', err.message);
    return new Response('Unauthorized', { status: 401 });
  }

  const activity = await request.json();

  // We only care about messages; ack everything else with 200
  if (activity.type !== 'message') {
    return new Response('', { status: 200 });
  }

  // Do the heavy lifting after returning 200 so Teams doesn't time out
  ctx.waitUntil(processMessage(activity, env));
  return new Response('', { status: 200 });
}

// ── Process an @admin message ───────────────────────────────────

async function processMessage(activity, env) {
  try {
    const teamId = activity.channelData?.team?.id;
    if (!teamId) {
      await replyToTeams(activity, env,
        'I can only process requests inside a Teams channel. Please @mention me in the team you want to add members to.');
      return;
    }

    const emails = extractEmails(activity.text ?? '');
    if (emails.length === 0) {
      await replyToTeams(activity, env,
        "I didn't find any email addresses in your message.\n\n" +
        'Try: **@admin** please add alice@company.com and bob@company.com');
      return;
    }

    let teamName;
    try { teamName = await getTeamName(env, teamId); }
    catch { teamName = teamId; }

    const requesterName = activity.from?.name ?? 'Unknown';
    const requesterAadId = activity.from?.aadObjectId ?? activity.from?.id;
    const cleanMsg = stripMentions(activity.text ?? '');

    const created = [];
    const failed = [];

    for (const email of emails) {
      try {
        const req = await createRequest(env.DB, {
          requesterName, requesterAadId, teamId, teamName,
          memberEmail: email, originalMessage: cleanMsg,
          conversationId: activity.conversation?.id,
          serviceUrl: activity.serviceUrl,
        });
        await postApprovalCard(env, req);
        created.push(req);
      } catch (err) {
        console.error(`Request failed for ${email}:`, err);
        failed.push({ email, error: err.message });
      }
    }

    const lines = [];
    if (created.length) {
      lines.push(`**Submitted ${created.length} request(s)** for admin approval:`);
      created.forEach(r => lines.push(`- \`${r.member_email}\` → request #${r.id}`));
      lines.push('', "You'll be notified here once each is approved or rejected.");
    }
    if (failed.length) {
      lines.push('', `**${failed.length} failed to submit:**`);
      failed.forEach(f => lines.push(`- \`${f.email}\`: ${f.error}`));
    }
    await replyToTeams(activity, env, lines.join('\n'));
  } catch (err) {
    console.error('processMessage error:', err);
  }
}

// ── Send a reply back to the Teams conversation ────────────────

let _botToken = { value: null, expiresAt: 0 };

async function getBotToken(env) {
  if (_botToken.value && Date.now() < _botToken.expiresAt - 60_000) return _botToken.value;
  const res = await fetch(
    'https://login.microsoftonline.com/botframework.com/oauth2/v2.0/token',
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        grant_type: 'client_credentials',
        client_id: env.BOT_ID,
        client_secret: env.BOT_PASSWORD,
        scope: 'https://api.botframework.com/.default',
      }),
    },
  );
  if (!res.ok) throw new Error(`Bot token error ${res.status}: ${await res.text()}`);
  const data = await res.json();
  _botToken = { value: data.access_token, expiresAt: Date.now() + data.expires_in * 1000 };
  return _botToken.value;
}

export async function replyToTeams(activity, env, text) {
  const token = await getBotToken(env);
  const base = (activity.serviceUrl ?? '').replace(/\/+$/, '');
  const convId = activity.conversation?.id;
  if (!base || !convId) return;

  await fetch(`${base}/v3/conversations/${convId}/activities`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ type: 'message', text }),
  });
}

// ── Helpers ─────────────────────────────────────────────────────

function stripMentions(text) {
  return text.replace(/<at[^>]*>.*?<\/at>\s*/gi, '').trim();
}
