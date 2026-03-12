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
  if (!data.ok) {
    const detail = data.response_metadata?.messages?.join('; ') ?? data.response_metadata?.errors?.map((e) => e?.message ?? e).join('; ') ?? '';
    const msg = detail ? `Slack ${method}: ${data.error} (${detail})` : `Slack ${method}: ${data.error}`;
    throw new Error(msg);
  }
  return data;
}

// ── Card body (reused when replacing only the buttons) ───────────

/** Build Teams deep link to open the team (groupId + tenantId + channel id for path). */
function teamDeepLink(request, env) {
  const channelId = request.teams_channel_id;
  const groupId = request.team_id;
  const tenantId = env.MS_TENANT_ID;
  if (!channelId || !groupId || !tenantId) return null;
  return `https://teams.microsoft.com/l/team/${encodeURIComponent(channelId)}/conversations?groupId=${encodeURIComponent(groupId)}&tenantId=${encodeURIComponent(tenantId)}`;
}

function cardBodyBlocks(request, env) {
  const requester = request.requester_email ?? request.requester_name ?? 'Someone';
  const teamLink = teamDeepLink(request, env);
  const teamDisplay = teamLink
    ? `<${teamLink}|${request.team_name}>`
    : request.team_name;
  // Blockquote lines (>) render as gray vertical bars in Slack
  const intro = `${requester} requested to invite one person to Adobe Enterprise Support`;
  const fields = `> *Email*: ${request.member_email}\n> *Team*: ${teamDisplay}`;
  return [
    {
      type: 'section',
      text: {
        type: 'mrkdwn',
        text: `${intro}\n\n${fields}`,
      },
    },
    { type: 'divider' },
  ];
}

/** Blocks for the card with a "Processing…" line instead of buttons (spinner state). */
function spinnerBlocks(request, env) {
  return [
    ...cardBodyBlocks(request, env),
    { type: 'section', text: { type: 'mrkdwn', text: ':hourglass_flowing_sand: Processing…' } },
  ];
}

// ── Post one approval card per email ────────────────────────────

export async function postApprovalCard(env, request) {
  const channel = env.SLACK_ADMIN_CHANNEL_ID;
  if (!channel) {
    throw new Error('SLACK_ADMIN_CHANNEL_ID is not set. Set it in wrangler.toml [vars] or in the Cloudflare dashboard so it is not removed on deploy.');
  }
  const requester = request.requester_email ?? request.requester_name ?? 'Someone';
  const result = await slack(env, 'chat.postMessage', {
    channel,
    text: `${requester} requested to invite ${request.member_email} to ${request.team_name}`,
    blocks: [
      ...cardBodyBlocks(request, env),
      {
        type: 'actions',
        block_id: 'approval_actions',
        elements: [
          {
            type: 'button',
            text: { type: 'plain_text', text: 'Approve' },
            style: 'primary',
            action_id: 'approve_request',
            value: String(request.id),
          },
          {
            type: 'button',
            text: { type: 'plain_text', text: 'Reject' },
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
  console.log('Slack interaction:', payload.type, payload.actions?.[0]?.action_id);

  if (payload.type === 'block_actions') {
    const action = payload.actions?.[0];
    const responseUrl = payload.response_url;
    if (
      responseUrl &&
      (action?.action_id === 'approve_request' || action?.action_id === 'reject_request')
    ) {
      const id = parseInt(action.value, 10);
      const request = await getRequest(env.DB, id);
      const blocks = request
        ? spinnerBlocks(request, env)
        : [{ type: 'section', text: { type: 'mrkdwn', text: ':hourglass_flowing_sand: Processing…' } }];
      const fallbackText = request
        ? `${request.requester_email ?? request.requester_name} requested to invite ${request.member_email} to ${request.team_name}\nProcessing…`
        : 'Processing…';
      try {
        await fetch(responseUrl, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            replace_original: true,
            text: fallbackText,
            blocks,
          }),
        });
      } catch (e) {
        console.error('response_url spinner failed:', e);
      }
    }
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
  console.log('Approve started for request', id);
  const request = await getRequest(env.DB, id);
  if (!request || request.status !== 'pending') {
    const statusText = !request
      ? 'This request could not be found.'
      : request.status === 'approved'
        ? 'This request has already been approved.'
        : 'This request has already been rejected.';
    console.log('Request already processed:', request?.status ?? 'not_found', '- updating card');
    const channelId = payload.channel?.id ?? payload.container?.channel_id ?? env.SLACK_ADMIN_CHANNEL_ID;
    const messageTs = payload.message?.ts ?? payload.container?.message_ts ?? request?.slack_message_ts;
    if (messageTs) {
      try {
        const blocks = request
          ? [...cardBodyBlocks(request, env), { type: 'section', text: { type: 'mrkdwn', text: statusText } }]
          : [{ type: 'section', text: { type: 'mrkdwn', text: statusText } }];
        const fallbackText = request
          ? `${request.requester_email ?? request.requester_name} requested to invite ${request.member_email} to ${request.team_name}\n${statusText}`
          : statusText;
        console.log('chat.update (already processed) channel=', channelId, 'ts=', messageTs);
        const updateRes = await slack(env, 'chat.update', {
          channel: channelId,
          ts: messageTs,
          text: fallbackText,
          blocks,
        });
        console.log('chat.update (already processed) result ok=', updateRes?.ok);
      } catch (updateErr) {
        console.error('chat.update (already processed) failed:', updateErr.message, updateErr);
        await slack(env, 'chat.postEphemeral', {
          channel: channelId,
          user: payload.user.id,
          text: statusText,
        }).catch(() => {});
      }
    } else {
      console.log('No messageTs in payload, sending ephemeral only');
      await slack(env, 'chat.postEphemeral', {
        channel: channelId,
        user: payload.user.id,
        text: statusText,
      }).catch(() => {});
    }
    return;
  }

  const reviewerName = payload.user.name ?? payload.user.username ?? payload.user.id;

  try {
    const result = await addTeamMember(env, request.team_id, request.member_email);
    const invited = result.invited === true;
    console.log(invited ? 'Invitation sent and member added for request' : 'Add member succeeded for request', id);

    await reviewRequest(env.DB, id, {
      status: 'approved',
      reviewerId: payload.user.id,
      reviewerName,
    });

    const approveText = invited
      ? `:email: <@${payload.user.id}> approved this request. An invitation was sent to ${request.member_email} and they've been added to the team. They'll have access once they accept the invite.`
      : `:white_check_mark: <@${payload.user.id}> approved this request. ${request.member_email} has been added to the team.`;
    const channelId = payload.channel?.id ?? env.SLACK_ADMIN_CHANNEL_ID;
    const messageTs = payload.message?.ts ?? request.slack_message_ts;
    if (messageTs) {
      try {
        await slack(env, 'chat.update', {
          channel: channelId,
          ts: messageTs,
          text: `${request.requester_email ?? request.requester_name} requested to invite ${request.member_email} to ${request.team_name}\n${approveText}`,
          blocks: [
            ...cardBodyBlocks(request, env),
            { type: 'section', text: { type: 'mrkdwn', text: approveText } },
          ],
        });
      } catch (updateErr) {
        console.error('Slack chat.update failed:', updateErr);
        await slack(env, 'chat.postEphemeral', {
          channel: channelId,
          user: payload.user.id,
          text: invited
            ? `✅ Invitation was sent to ${request.member_email} and they've been added to the team, but the card could not be updated: ${updateErr.message}`
            : `✅ Member was added to the team, but the card could not be updated: ${updateErr.message}`,
        }).catch(() => {});
      }
    } else {
      console.error('No message ts for chat.update', { requestId: id });
      await slack(env, 'chat.postEphemeral', {
        channel: channelId,
        user: payload.user.id,
        text: invited
          ? `✅ An invitation was sent to ${request.member_email} and they've been added to the team. The approval card could not be updated (missing message reference).`
          : '✅ Member was added to the team. The approval card could not be updated (missing message reference).',
      }).catch(() => {});
    }

    // Notify requester in Teams
    if (request.service_url && request.conversation_id) {
      try {
        await replyToTeams(
          { serviceUrl: request.service_url, conversation: { id: request.conversation_id } },
          env,
          `✅ ${request.member_email} has been added to this team.`,
        );
      } catch (teamsErr) {
        console.error('replyToTeams failed:', teamsErr);
      }
    }
  } catch (err) {
    console.error('Approve failed:', err);
    const channelId = payload.channel?.id ?? payload.container?.channel_id ?? env.SLACK_ADMIN_CHANNEL_ID;
    const messageTs = payload.message?.ts ?? payload.container?.message_ts ?? request.slack_message_ts;
    const errorCardText = ':warning: An error occurred…';
    if (messageTs) {
      try {
        const fallbackText = `${request.requester_email ?? request.requester_name} requested to invite ${request.member_email} to ${request.team_name}\n${errorCardText}`;
        await slack(env, 'chat.update', {
          channel: channelId,
          ts: messageTs,
          text: fallbackText,
          blocks: [
            ...cardBodyBlocks(request, env),
            { type: 'section', text: { type: 'mrkdwn', text: errorCardText } },
          ],
        });
        const errorSnippet = String(err.message ?? err).slice(0, 2900);
        await slack(env, 'chat.postMessage', {
          channel: channelId,
          thread_ts: messageTs,
          text: 'Error response',
          blocks: [
            {
              type: 'section',
              text: {
                type: 'mrkdwn',
                text: '```\n' + errorSnippet.replace(/```/g, '`\u200b``') + '\n```',
              },
            },
          ],
        });
      } catch (updateErr) {
        console.error('chat.update (approve error) failed:', updateErr);
        const errorText = `Failed to add member: ${err.message}`;
        await slack(env, 'chat.postEphemeral', {
          channel: channelId,
          user: payload.user.id,
          text: errorText,
        }).catch(() => {});
      }
    } else {
      await slack(env, 'chat.postEphemeral', {
        channel: channelId,
        user: payload.user.id,
        text: `Failed to add member: ${err.message}`,
      }).catch(() => {});
    }
    if (err.message?.toLowerCase().includes('not found') && request.service_url && request.conversation_id) {
      try {
        await replyToTeams(
          { serviceUrl: request.service_url, conversation: { id: request.conversation_id } },
          env,
          `The following user was not found: ${request.member_email}`,
        );
      } catch (teamsErr) {
        console.error('replyToTeams (not found) failed:', teamsErr);
      }
    }
  }
}

async function openRejectModal(payload, action, env) {
  const id = parseInt(action.value, 10);
  const request = await getRequest(env.DB, id);
  if (!request || request.status !== 'pending') {
    const statusText = !request
      ? 'This request could not be found.'
      : request.status === 'approved'
        ? 'This request has already been approved.'
        : 'This request has already been rejected.';
    console.log('Reject: request already processed:', request?.status ?? 'not_found', '- updating card');
    const channelId = payload.channel?.id ?? payload.container?.channel_id ?? env.SLACK_ADMIN_CHANNEL_ID;
    const messageTs = payload.message?.ts ?? payload.container?.message_ts ?? request?.slack_message_ts;
    if (messageTs) {
      try {
        const blocks = request
          ? [...cardBodyBlocks(request, env), { type: 'section', text: { type: 'mrkdwn', text: statusText } }]
          : [{ type: 'section', text: { type: 'mrkdwn', text: statusText } }];
        const fallbackText = request
          ? `${request.requester_email ?? request.requester_name} requested to invite ${request.member_email} to ${request.team_name}\n${statusText}`
          : statusText;
        console.log('chat.update (reject already processed) channel=%s ts=%s', channelId, messageTs);
        await slack(env, 'chat.update', {
          channel: channelId,
          ts: messageTs,
          text: fallbackText,
          blocks,
        });
      } catch (updateErr) {
        console.error('chat.update (reject already processed) failed:', updateErr.message);
        await slack(env, 'chat.postEphemeral', {
          channel: channelId,
          user: payload.user.id,
          text: statusText,
        }).catch(() => {});
      }
    } else {
      await slack(env, 'chat.postEphemeral', {
        channel: channelId,
        user: payload.user.id,
        text: statusText,
      }).catch(() => {});
    }
    return;
  }

  const channelId = payload.channel?.id ?? payload.container?.channel_id ?? env.SLACK_ADMIN_CHANNEL_ID;
  const messageTs = payload.message?.ts ?? payload.container?.message_ts ?? request.slack_message_ts;
  await slack(env, 'views.open', {
    trigger_id: payload.trigger_id,
    view: {
      type: 'modal',
      callback_id: 'reject_reason_modal',
      private_metadata: JSON.stringify({
        requestId: id,
        reviewerId: payload.user.id,
        reviewerName: payload.user.name ?? payload.user.username ?? payload.user.id,
        channelId,
        messageTs,
      }),
      title: { type: 'plain_text', text: 'Reject Request' },
      submit: { type: 'plain_text', text: 'Reject' },
      close: { type: 'plain_text', text: 'Cancel' },
      blocks: [
        {
          type: 'section',
          text: {
            type: 'mrkdwn',
            text: `Rejecting request to add ${request.member_email} to ${request.team_name}.`,
          },
        },
        {
          type: 'input',
          block_id: 'reject_reason',
          optional: true,
          label: { type: 'plain_text', text: 'Reason (optional)' },
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
  const meta = JSON.parse(payload.view.private_metadata);
  const { requestId, reviewerId, reviewerName, channelId, messageTs } = meta;
  const reviewNote = payload.view.state.values.reject_reason?.reason?.value ?? null;

  const request = await getRequest(env.DB, requestId);
  if (!request || request.status !== 'pending') return;

  const channel = channelId ?? env.SLACK_ADMIN_CHANNEL_ID;
  const ts = messageTs ?? request.slack_message_ts;

  try {
    if (ts) {
      try {
        await slack(env, 'chat.update', {
          channel,
          ts,
          text: `${request.requester_email ?? request.requester_name} requested to invite ${request.member_email} to ${request.team_name}\nProcessing…`,
          blocks: spinnerBlocks(request, env),
        });
      } catch (e) {
        console.error('Reject spinner update failed:', e);
      }
    }

    await reviewRequest(env.DB, requestId, {
      status: 'rejected', reviewerId, reviewerName, reviewNote,
    });

    const reason = reviewNote || '—';
    const rejectText = `:no_entry_sign: <@${reviewerId}> rejected this request. Reason: ${reason}`;
    const requester = request.requester_email ?? request.requester_name ?? 'Someone';
    await slack(env, 'chat.update', {
      channel,
      ts,
      text: `${requester} requested to invite ${request.member_email} to ${request.team_name}\n${rejectText}`,
      blocks: [
        ...cardBodyBlocks(request, env),
        { type: 'section', text: { type: 'mrkdwn', text: rejectText } },
      ],
    });

    if (request.service_url && request.conversation_id) {
      const reason = reviewNote || '—';
      await replyToTeams(
        { serviceUrl: request.service_url, conversation: { id: request.conversation_id } },
        env,
        `🚫 ${request.member_email} was not added to this team.\n\nReason: ${reason}`,
      );
    }
  } catch (err) {
    console.error('Reject failed:', err);
  }
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
