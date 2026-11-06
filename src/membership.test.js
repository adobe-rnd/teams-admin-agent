import { test } from 'node:test';
import assert from 'node:assert/strict';
import { MembershipRequests } from './membership.js';
import { submitRequests, replyText } from './requests.js';

const TEAM = '16e06978-0000-4000-8000-000000000000';
const THREAD = '19:abc@thread.tacv2;messageid=1793904314799';

function setup(over = {}) {
  const calls = { submit: [], reply: [] };
  const deps = {
    teamName: async () => 'aem-example',
    submit: async (env, input) => {
      calls.submit.push(input);
      return {
        created: input.emails.map((e, i) => ({ id: i + 1, member_email: e })),
        skipped: [],
        failed: [],
      };
    },
    reply: async (activity, env, text) => {
      calls.reply.push({ activity, text });
      return true;
    },
    ...over,
  };
  return { rpc: new MembershipRequests({}, deps), calls };
}

const valid = {
  action: 'add',
  emails: ['Someone@Example.com'],
  teamId: TEAM,
  conversationId: THREAD,
  requester: { name: 'A Person', email: 'person@example.com', aadObjectId: 'aad-1' },
  via: 'opspile',
  message: 'please add someone@example.com',
};

test('submit: files a request as a mention would, and answers in the thread', async () => {
  const { rpc, calls } = setup();
  const out = await rpc.submit(valid);
  assert.deepEqual(out, {
    ok: true, created: [{ id: 1, email: 'someone@example.com' }], skipped: [], failed: [], replied: true,
  });
  const [input] = calls.submit;
  assert.equal(input.action, 'add');
  assert.deepEqual(input.emails, ['someone@example.com']);
  assert.equal(input.teamName, 'aem-example');
  assert.equal(input.teamsChannelId, '19:abc@thread.tacv2');
  assert.equal(input.conversationId, THREAD);
  assert.equal(input.requesterEmail, 'person@example.com');
  assert.equal(input.via, 'opspile');
  assert.equal(calls.reply[0].activity.conversation.id, THREAD);
  assert.match(calls.reply[0].text, /Invitation request submitted for approval/);
});

test('submit: refuses wrong input without filing anything', async () => {
  const { rpc, calls } = setup();
  const cases = [
    [{ action: 'promote' }, /action/],
    [{ emails: ['not an address'] }, /no email/],
    [{ emails: Array.from({ length: 21 }, (_, i) => `p${i}@example.com`) }, /at most 20/],
    [{ teamId: '19:abc@thread.tacv2' }, /group id/],
    [{ conversationId: 'https://evil.example/' }, /Teams channel or thread/],
    [{ requester: {}, via: '' }, /who asks/],
  ];
  for (const [change, error] of cases) {
    // eslint-disable-next-line no-await-in-loop
    const out = await rpc.submit({ ...valid, ...change });
    assert.equal(out.ok, false);
    assert.match(out.error, error);
  }
  assert.equal(calls.submit.length, 0);
  assert.equal(calls.reply.length, 0);
});

test('submit: says when admin-bot could not answer in the thread', async () => {
  const { rpc } = setup({ reply: async () => { throw new Error('403'); } });
  const out = await rpc.submit(valid);
  assert.equal(out.ok, true);
  assert.equal(out.replied, false);
});

test('submitRequests: skips people who are already members, and cards the rest', async () => {
  const rows = [];
  const env = {
    DB: {
      prepare: (sql) => ({
        bind: (...args) => ({
          run: async () => {
            rows.push({ sql, args });
            return { meta: { last_row_id: rows.length } };
          },
          first: async () => ({ id: args[0], member_email: rows.at(-1).args[6], via: rows.at(-1).args[11] }),
        }),
      }),
    },
  };
  const cards = [];
  const result = await submitRequests(env, {
    action: 'add', emails: ['in@example.com', 'new@example.com'], teamId: TEAM, via: 'opspile',
  }, {
    memberEmails: async () => new Set(['in@example.com']),
    approvalCard: async (e, req) => cards.push(req),
  });
  assert.deepEqual(result.skipped, ['in@example.com']);
  assert.equal(result.created.length, 1);
  assert.equal(cards[0].member_email, 'new@example.com');
  assert.equal(cards[0].via, 'opspile');
  assert.match(replyText('add', result), /already a member[\s\S]*submitted for approval/);
});
