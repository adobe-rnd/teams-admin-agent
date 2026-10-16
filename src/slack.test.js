import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createHmac } from 'node:crypto';
import { handleSlackInteraction } from './slack.js';

const graphError = {
  error: {
    code: 'TooManyRequests',
    message: 'Too many requests to this resource.',
    innerError: { code: 'ActivityLimitReached' },
  },
};

const scenarios = [
  { name: 'existing member approval', target: 'members' },
  { name: 'guest invitation', guest: true, target: 'invitations' },
  { name: 'guest membership after invitation', guest: true, target: 'members' },
  { name: 'display-name modal invitation', guest: true, modal: true, target: 'invitations' },
  { name: 'display-name modal membership', guest: true, modal: true, target: 'members' },
  { name: 'member removal', remove: true, target: 'members' },
];

for (const scenario of scenarios) {
  test(`429 restores controls and allows retry: ${scenario.name}`, async (t) => {
    const { request, calls, submit } = approvalFixture(t, scenario);

    await submit();

    assert.equal(request.status, 'pending');
    const card = calls.filter(c => c.method === 'chat.update').at(-1).body;
    assert.equal(card.channel, 'admin-channel');
    assert.equal(card.ts, '123.456');
    const actions = card.blocks.find(b => b.type === 'actions');
    assert.ok(actions, 'failed approval must restore the action buttons');
    assert.deepEqual(actions.elements.map(b => [b.action_id, b.value]), [
      ['approve_request', '42'],
      ['reject_request', '42'],
    ]);
    if (scenario.remove) assert.ok(actions.elements[0].confirm);
    if (scenario.modal) assert.match(card.blocks[0].text.text, /Test Person/);
    const thread = calls.find(c => c.method === 'chat.postMessage').body;
    assert.equal(thread.thread_ts, '123.456');
    assert.match(thread.text, /wait.*retry/i);
    assert.match(thread.text, /TooManyRequests/);
    assert.match(thread.text, /ActivityLimitReached/);
    assert.match(thread.text, /\(429\)/);

    await submit();

    assert.equal(request.status, 'approved');
    const approvedCard = calls.filter(c => c.method === 'chat.update').at(-1).body;
    assert.equal(approvedCard.blocks.some(b => b.type === 'actions'), false);
    assert.match(approvedCard.text, scenario.remove ? /removed/ : /approved/);
  });
}

test('non-recoverable Graph errors keep the existing error card behavior', async (t) => {
  const { request, calls, submit } = approvalFixture(t, { target: 'members', status: 400 });
  await submit();
  assert.equal(request.status, 'pending');
  const card = calls.filter(c => c.method === 'chat.update').at(-1).body;
  assert.equal(card.blocks.some(b => b.type === 'actions'), false);
  assert.match(card.text, /An error occurred/);
});

test('team ownership errors still restore approval controls', async (t) => {
  const { request, calls, submit } = approvalFixture(t, { target: 'members', status: 403 });
  await submit();
  assert.equal(request.status, 'pending');
  const card = calls.filter(c => c.method === 'chat.update').at(-1).body;
  assert.ok(card.blocks.find(b => b.type === 'actions'));
  assert.match(calls.find(c => c.method === 'chat.postMessage').body.text, /not an owner/);
});

function approvalFixture(t, { guest = false, modal = false, remove = false, target, status = 429 }) {
  const request = {
    id: 42,
    status: 'pending',
    action: remove ? 'remove' : 'add',
    member_email: modal ? 'person@adobe.com' : 'person@example.com',
    requester_name: 'Requester',
    team_id: 'team-id',
    team_name: 'Test Team',
    slack_message_ts: '123.456',
  };
  const env = {
    SLACK_SIGNING_SECRET: 'test-signing-secret',
    SLACK_ADMIN_CHANNEL_ID: 'admin-channel',
    DELEGATED_REFRESH_TOKEN: 'test-refresh-token',
    MS_TENANT_ID: 'test-tenant',
    DB: {
      prepare(sql) {
        return {
          bind(...values) {
            return {
              async first() { return { ...request }; },
              async run() {
                assert.match(sql, /UPDATE requests/);
                assert.equal(values.at(-1), request.id);
                request.status = values[0];
              },
            };
          },
        };
      },
    },
  };
  const calls = [];
  let userExists = !guest;
  let failed = false;
  t.mock.method(globalThis, 'fetch', async (url, options = {}) => {
    if (url === 'https://slack.test/response') return Response.json({ ok: true });
    if (url.startsWith('https://slack.com/api/')) {
      calls.push({ method: url.split('/').at(-1), body: JSON.parse(options.body) });
      return Response.json({ ok: true });
    }
    if (url.startsWith('https://login.microsoftonline.com/')) {
      return Response.json({ access_token: 'test-access-token', expires_in: 3600 });
    }
    const path = new URL(url).pathname;
    if (path.startsWith('/v1.0/users')) {
      if (userExists) return Response.json({ id: 'user-id', displayName: 'Test Person' });
      return path === '/v1.0/users'
        ? Response.json({ value: [] })
        : Response.json({ error: 'not found' }, { status: 404 });
    }
    if (remove && options.method === 'GET' && path.endsWith('/members')) {
      return Response.json({ value: [{ id: 'membership-id', email: request.member_email }] });
    }
    assert.ok(path.includes('/members') || path === '/v1.0/invitations', `Unexpected fetch: ${url}`);
    if (!failed && path.includes(`/${target}`)) {
      failed = true;
      return Response.json(graphError, { status });
    }
    if (path === '/v1.0/invitations') {
      userExists = true;
      return Response.json({ invitedUser: { id: 'user-id' } });
    }
    return Response.json({});
  });

  async function submit() {
    const payload = modal ? {
      type: 'view_submission',
      view: {
        callback_id: 'approve_display_name_modal',
        private_metadata: JSON.stringify({
          requestId: request.id,
          reviewerId: 'admin-id',
          reviewerName: 'Admin',
          channelId: 'admin-channel',
          messageTs: request.slack_message_ts,
        }),
        state: { values: { approve_display_name: { display_name: { value: 'Test Person' } } } },
      },
    } : {
      type: 'block_actions',
      actions: [{ action_id: 'approve_request', value: String(request.id) }],
      channel: { id: 'admin-channel' },
      message: { ts: request.slack_message_ts },
      user: { id: 'admin-id', name: 'Admin' },
      response_url: 'https://slack.test/response',
    };
    const body = new URLSearchParams({ payload: JSON.stringify(payload) }).toString();
    const timestamp = String(Math.floor(Date.now() / 1000));
    const signature = 'v0=' + createHmac('sha256', env.SLACK_SIGNING_SECRET)
      .update(`v0:${timestamp}:${body}`).digest('hex');
    const tasks = [];
    const response = await handleSlackInteraction(new Request('https://worker.test/api/slack/interactions', {
      method: 'POST',
      headers: { 'x-slack-request-timestamp': timestamp, 'x-slack-signature': signature },
      body,
    }), env, { waitUntil(task) { tasks.push(task); } });
    assert.equal(response.status, 200);
    await Promise.all(tasks);
  }

  return { request, calls, submit };
}
