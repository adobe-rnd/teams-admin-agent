/**
 * Slack integration — posts approval cards and handles button / modal interactions.
 * Uses raw Slack Web API via fetch (no SDK).
 */
import { setSlackMessageTs, getRequest, reviewRequest } from './db.js';
import { addTeamMember } from './graph.js';
import { replyToTeams } from './teams.js';

// ── Slack Web API helper ────────────────────────────────────────

async function slack(env, method, body) {
  const res = await fetch(`https://slack.com/api/${method}`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${env.SLACK_BOT_TOKEN}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(body),
  });
  const data = await res.json();
  if (!data.ok) throw new Error(`Slack ${method}: ${data.error}`);
  return data;
}

// ── Post one approval card per email ────────────────────────────

export async function postApprovalCard(env, request) {
  const result = await slack(env, 'chat.postMessage', {
    channel: env.SLACK_ADMIN_CHANNEL_ID,
    text: `Request #${request.id}: add ${request.member_email} to ${request.team_name} (from ${request.requester_name})`,
    blocks: [
      {
        type: 'header',
        text: { type: 'plain_text', text: `📋  Request #${request.id}` },
      },
      {
        type: 'section',
        fields: [
          { type: 'mrkdwn', text: `*Requested by:*\n${request.requester_name}` },
          { type: 'mrkdwn', text: `*Date:*\n${request.created_at}` },
        ],
      },
      {
        type: 'section',
        fields: [
          { type: 'mrkdwn', text: `*Microsoft Team:*\n${request.team_name}` },
          { type: 'mrkdwn', text: `*Email to add:*\n${request.member_email}` },
        ],
      },
      ...(request.original_message
        ? [{
            type: 'section',
            text: { type: 'mrkdwn', text: `*Original message:*\n> ${request.original_message.replace(/\n/g, '\n> ')}` },
          }]
        : []),
      { type: 'divider' },
      {
        type: 'actions',
        block_id: 'approval_actions',
        elements: [
          {
            type: 'button',
            text: { type: 'plain_text', text: '✅ Approve' },
            style: 'primary',
            action_id: 'approve_request',
            value: String(request.id),
            confirm: {
              title: { type: 'plain_text', text: 'Confirm Approval' },
              text: { type: 'mrkdwn', text: `Add *${request.member_email}* to *${request.team_name}*?` },
              confirm: { type: 'plain_text', text: 'Approve' },
              deny: { type: 'plain_text', text: 'Cancel' },
            },
          },
          {
            type: 'button',
            text: { type: 'plain_text', text: '❌ Reject' },
            style: 'danger',
            action_id: 'reject_request',
            value: String(request.id),
          },
        ],
      },
    ],
  });

  await setSlackMessageTs(env.DB, request.id, result.ts);
}

// ── Incoming Slack interaction webhook ──────────────────────────

export async function handleSlackInteraction(request, env, ctx) {
  const rawBody = await request.text();

  // Verify Slack request signature + reject replay attacks
  const ts = request.headers.get('x-slack-request-timestamp') ?? '';
  const sig = request.headers.get('x-slack-signature') ?? '';

  // Reject requests older than 5 minutes to prevent replay attacks
  const age = Math.abs(Math.floor(Date.now() / 1000) - Number(ts));
  if (isNaN(age) || age > 300) {
    return new Response('Request too old', { status: 403 });
  }

  if (!(await verifySignature(env.SLACK_SIGNING_SECRET, ts, rawBody, sig))) {
    return new Response('Invalid signature', { status: 401 });
  }

  const params = new URLSearchParams(rawBody);
  const payload = JSON.parse(params.get('payload'));

  if (payload.type === 'block_actions') {
    // Return 200 immediately; do work in the background
    ctx.waitUntil(handleBlockAction(payload, env));
    return new Response('', { status: 200 });
  }

  if (payload.type === 'view_submission' && payload.view?.callback_id === 'reject_reason_modal') {
    ctx.waitUntil(handleRejectSubmission(payload, env));
    return Response.json({ response_action: 'clear' });
  }

  return new Response('', { status: 200 });
}

// ── Button clicks ───────────────────────────────────────────────

async function handleBlockAction(payload, env) {
  const action = payload.actions?.[0];
  if (!action) return;

  if (action.action_id === 'approve_request') {
    await handleApprove(payload, action, env);
  } else if (action.action_id === 'reject_request') {
    await openRejectModal(payload, action, env);
  }
}

async function handleApprove(payload, action, env) {
  const id = parseInt(action.value, 10);
  const request = await getRequest(env.DB, id);
  if (!request || request.status !== 'pending') {
    await slack(env, 'chat.postEphemeral', {
      channel: payload.channel.id,
      user: payload.user.id,
      text: 'This request has already been processed.',
    });
    return;
  }

  const reviewerName = payload.user.name ?? payload.user.username ?? payload.user.id;

  try {
    await addTeamMember(env, request.team_id, request.member_email);

    const updated = await reviewRequest(env.DB, id, {
      status: 'approved',
      reviewerId: payload.user.id,
      reviewerName,
    });

    await slack(env, 'chat.update', {
      channel: env.SLACK_ADMIN_CHANNEL_ID,
      ts: request.slack_message_ts,
      text: `Request #${id} approved by ${reviewerName}`,
      blocks: reviewedBlocks(updated, 'approved'),
    });

    // Notify requester in Teams
    if (request.service_url && request.conversation_id) {
      await replyToTeams(
        { serviceUrl: request.service_url, conversation: { id: request.conversation_id } },
        env,
        `✅ Request #${id} approved — **${request.member_email}** has been added to **${request.team_name}**.`,
      );
    }
  } catch (err) {
    console.error('Approve failed:', err);
    await slack(env, 'chat.postEphemeral', {
      channel: payload.channel.id,
      user: payload.user.id,
      text: `Failed to add member: ${err.message}`,
    });
  }
}

async function openRejectModal(payload, action, env) {
  const id = parseInt(action.value, 10);
  const request = await getRequest(env.DB, id);
  if (!request || request.status !== 'pending') return;

  await slack(env, 'views.open', {
    trigger_id: payload.trigger_id,
    view: {
      type: 'modal',
      callback_id: 'reject_reason_modal',
      private_metadata: JSON.stringify({
        requestId: id,
        reviewerId: payload.user.id,
        reviewerName: payload.user.name ?? payload.user.username ?? payload.user.id,
      }),
      title: { type: 'plain_text', text: 'Reject Request' },
      submit: { type: 'plain_text', text: 'Reject' },
      close: { type: 'plain_text', text: 'Cancel' },
      blocks: [
        {
          type: 'section',
          text: { type: 'mrkdwn', text: `Rejecting request *#${id}* — add *${request.member_email}* to *${request.team_name}*.` },
        },
        {
          type: 'input',
          block_id: 'reject_reason',
          optional: true,
          label: { type: 'plain_text', text: 'Reason' },
          element: {
            type: 'plain_text_input',
            action_id: 'reason',
            multiline: true,
            placeholder: { type: 'plain_text', text: 'Optional reason…' },
          },
        },
      ],
    },
  });
}

// ── Reject modal submission ─────────────────────────────────────

async function handleRejectSubmission(payload, env) {
  const { requestId, reviewerId, reviewerName } = JSON.parse(payload.view.private_metadata);
  const reviewNote = payload.view.state.values.reject_reason?.reason?.value ?? null;

  const request = await getRequest(env.DB, requestId);
  if (!request || request.status !== 'pending') return;

  try {
    const updated = await reviewRequest(env.DB, requestId, {
      status: 'rejected', reviewerId, reviewerName, reviewNote,
    });

    await slack(env, 'chat.update', {
      channel: env.SLACK_ADMIN_CHANNEL_ID,
      ts: request.slack_message_ts,
      text: `Request #${requestId} rejected by ${reviewerName}`,
      blocks: reviewedBlocks(updated, 'rejected'),
    });

    if (request.service_url && request.conversation_id) {
      const note = reviewNote ? `\n> ${reviewNote}` : '';
      await replyToTeams(
        { serviceUrl: request.service_url, conversation: { id: request.conversation_id } },
        env,
        `🚫 Request #${requestId} rejected — **${request.member_email}** was *not* added to **${request.team_name}**.${note}`,
      );
    }
  } catch (err) {
    console.error('Reject failed:', err);
  }
}

// ── Card after review ───────────────────────────────────────────

function reviewedBlocks(req, outcome) {
  const ok = outcome === 'approved';
  return [
    { type: 'header', text: { type: 'plain_text', text: `${ok ? '✅' : '🚫'}  Request #${req.id} — ${ok ? 'Approved' : 'Rejected'}` } },
    { type: 'section', fields: [
      { type: 'mrkdwn', text: `*Requested by:*\n${req.requester_name}` },
      { type: 'mrkdwn', text: `*Reviewed by:*\n${req.reviewer_name}` },
    ]},
    { type: 'section', fields: [
      { type: 'mrkdwn', text: `*Team:*\n${req.team_name}` },
      { type: 'mrkdwn', text: `*Member:*\n${req.member_email}` },
    ]},
    ...(req.review_note ? [{ type: 'section', text: { type: 'mrkdwn', text: `*Note:*\n${req.review_note}` } }] : []),
    { type: 'context', elements: [{ type: 'mrkdwn', text: `${ok ? 'Approved' : 'Rejected'} on ${req.reviewed_at}` }] },
  ];
}

// ── Slack signature verification ────────────────────────────────

async function verifySignature(secret, timestamp, body, expected) {
  const enc = new TextEncoder();
  const key = await crypto.subtle.importKey(
    'raw', enc.encode(secret), { name: 'HMAC', hash: 'SHA-256' }, false, ['sign'],
  );
  const sig = await crypto.subtle.sign('HMAC', key, enc.encode(`v0:${timestamp}:${body}`));
  const computed = 'v0=' + [...new Uint8Array(sig)].map(b => b.toString(16).padStart(2, '0')).join('');
  // constant-time comparison
  if (computed.length !== expected.length) return false;
  let result = 0;
  for (let i = 0; i < computed.length; i++) result |= computed.charCodeAt(i) ^ expected.charCodeAt(i);
  return result === 0;
}
