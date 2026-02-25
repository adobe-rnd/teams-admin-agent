/**
 * Handles incoming Bot Framework activities from Microsoft Teams.
 *
 * Validates the JWT, parses @admin-bot messages for emails, creates DB
 * records, posts Slack approval cards, and replies in Teams.
 */
import { createRemoteJWKSet, jwtVerify } from 'jose';
import { extractEmails, hasAddIntent } from './parser.js';
import { createRequest } from './db.js';
import { postApprovalCard } from './slack.js';
import { getTeamName, getRequesterEmail, getTeamMemberEmails } from './graph.js';

// JWKS endpoints — Bot Framework (multi-tenant) and Entra ID (single-tenant)
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

  const token = auth.slice(7);

  // Single-tenant bots receive tokens from the tenant's Entra ID endpoint;
  // multi-tenant bots receive tokens from api.botframework.com.
  // Try both JWKS sources so either configuration works.
  const tenantJWKS = createRemoteJWKSet(
    new URL(`https://login.microsoftonline.com/${env.MS_TENANT_ID}/discovery/v2.0/keys`),
  );

  let verified = false;
  for (const { jwks, issuer } of [
    { jwks: tenantJWKS, issuer: `https://sts.windows.net/${env.MS_TENANT_ID}/` },
    { jwks: BF_JWKS, issuer: 'https://api.botframework.com' },
  ]) {
    try {
      await jwtVerify(token, jwks, {
        audience: env.BOT_ID,
        issuer,
        clockTolerance: 300,
      });
      verified = true;
      break;
    } catch { /* try next */ }
  }

  if (!verified) {
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

// ── Process an @admin-bot message ──────────────────────────────────

function looksLikeThreadId(s) {
  if (!s || typeof s !== 'string') return false;
  return s.includes('@thread') || (s.includes(':') && s.length > 36);
}

/** Fetch team display name and AAD group ID from Bot Framework Connector. */
async function getTeamDetailsFromConnector(activity, env, teamsTeamId) {
  const base = (activity.serviceUrl ?? '').replace(/\/+$/, '');
  if (!base) return null;
  try {
    const token = await getBotToken(env);
    const res = await fetch(`${base}/v3/teams/${encodeURIComponent(teamsTeamId)}`, {
      method: 'GET',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    });
    if (!res.ok) {
      console.error('Connector GET /v3/teams failed:', res.status, await res.text());
      return null;
    }
    const data = await res.json();
    return { name: data.name ?? null, aadGroupId: data.aadGroupId ?? null };
  } catch (err) {
    console.error('getTeamDetailsFromConnector error:', err.message);
    return null;
  }
}

async function processMessage(activity, env) {
  try {
    const teamsTeamId = activity.channelData?.team?.id;
    if (!teamsTeamId) return;

    const text = activity.text ?? '';
    if (!hasAddIntent(text)) return;

    const emails = extractEmails(text);
    if (emails.length === 0) return;

    // Resolve display name and Graph-compatible ID via Connector (team.id is thread-style, not Graph GUID)
    let teamName = activity.channelData?.team?.name ?? null;
    let teamId = teamsTeamId;
    const details = await getTeamDetailsFromConnector(activity, env, teamsTeamId);
    if (details?.name) teamName = details.name;
    if (details?.aadGroupId) teamId = details.aadGroupId;
    const isGraphGuid = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(teamId);
    if (isGraphGuid && (!teamName || teamName === teamId)) {
      try { teamName = await getTeamName(env, teamId); } catch { teamName = null; }
    }
    if (!teamName || teamName === teamId || looksLikeThreadId(teamName)) {
      teamName = 'This team';
    }

    if (!isGraphGuid) {
      await replyToTeams(activity, env, "This team couldn't be resolved (missing Graph ID). The bot may need to be re-added to the team, or check the Connector endpoint.");
      return;
    }

    const requesterName = activity.from?.name ?? 'Unknown';
    const requesterAadId = activity.from?.aadObjectId ?? activity.from?.id;
    const requesterEmail = await getRequesterEmail(env, requesterAadId);
    const cleanMsg = stripMentions(activity.text ?? '');

    // Exclude emails that are already team members (check before creating requests)
    let memberEmails;
    try {
      memberEmails = await getTeamMemberEmails(env, teamId);
    } catch (err) {
      console.error('getTeamMemberEmails failed:', err);
      memberEmails = new Set();
    }
    const toAdd = [];
    const alreadyMembers = [];
    for (const email of emails) {
      const normalized = email.toLowerCase().trim();
      if (memberEmails.has(normalized)) {
        alreadyMembers.push(email);
      } else {
        toAdd.push(email);
      }
    }

    const created = [];
    const failed = [];

    for (const email of toAdd) {
      try {
        const req = await createRequest(env.DB, {
          requesterName, requesterAadId, requesterEmail, teamId, teamName, teamsChannelId: teamsTeamId,
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
    if (alreadyMembers.length > 0) {
      const alreadyText =
        alreadyMembers.length === 1
          ? 'The following user is already a member of this team:'
          : 'The following users are already a member of this team:';
      lines.push(alreadyText, '');
      alreadyMembers.forEach((e, i) => {
        lines.push(`• ${e}`);
        if (i < alreadyMembers.length - 1) lines.push('');
      });
      if (created.length > 0 || failed.length > 0) lines.push('', '');
    }
    if (created.length) {
      const n = created.length;
      const requestWord = n === 1 ? 'request' : 'requests';
      lines.push(`Invitation ${requestWord} submitted for approval:`, '');
      created.forEach((r, i) => {
        lines.push(`• ${r.member_email}`);
        if (i < created.length - 1) lines.push('');
      });
      lines.push('', '', `You'll be notified here once the ${requestWord} ${n === 1 ? 'has' : 'have'} been reviewed.`);
    }
    if (failed.length) {
      lines.push('', `**${failed.length} failed to submit:**`);
      failed.forEach(f => lines.push(`- \`${f.email}\`: ${f.error}`));
    }
    if (lines.length > 0) {
      await replyToTeams(activity, env, lines.join('\r\n'));
    }
  } catch (err) {
    console.error('processMessage error:', err);
  }
}

// ── Send a reply back to the Teams conversation ────────────────

let _botToken = { value: null, expiresAt: 0 };

async function getBotToken(env) {
  if (_botToken.value && Date.now() < _botToken.expiresAt - 60_000) return _botToken.value;

  // Single-tenant bots use the tenant's token endpoint;
  // multi-tenant bots use botframework.com. Try tenant first.
  const endpoints = [
    `https://login.microsoftonline.com/${env.MS_TENANT_ID}/oauth2/v2.0/token`,
    'https://login.microsoftonline.com/botframework.com/oauth2/v2.0/token',
  ];

  for (const url of endpoints) {
    const res = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        grant_type: 'client_credentials',
        client_id: env.BOT_ID,
        client_secret: env.BOT_PASSWORD,
        scope: 'https://api.botframework.com/.default',
      }),
    });
    if (res.ok) {
      const data = await res.json();
      _botToken = { value: data.access_token, expiresAt: Date.now() + data.expires_in * 1000 };
      return _botToken.value;
    }
  }

  throw new Error('Failed to acquire bot token from any endpoint');
}

export async function replyToTeams(activity, env, text) {
  const token = await getBotToken(env);
  const base = (activity.serviceUrl ?? '').replace(/\/+$/, '');
  const convId = activity.conversation?.id;
  if (!base || !convId) return;

  const res = await fetch(`${base}/v3/conversations/${convId}/activities`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({
      type: 'message',
      text,
      textFormat: 'markdown',
      from: { id: env.BOT_ID, name: 'admin-bot' },
      conversation: activity.conversation,
      recipient: activity.from,
      replyToId: activity.id,
    }),
  });
  if (!res.ok) {
    const errBody = await res.text();
    console.error(`replyToTeams: ${res.status} ${errBody}`);
  }
}

// ── Helpers ─────────────────────────────────────────────────────

function stripMentions(text) {
  return text.replace(/<at[^>]*>.*?<\/at>\s*/gi, '').trim();
}
