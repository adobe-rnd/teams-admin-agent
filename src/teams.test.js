import { test } from 'node:test';
import assert from 'node:assert/strict';
import { emailsToRequest, isCommandActivity } from './teams.js';

const CONV = '19:abc@thread.tacv2;messageid=1795168746480';

test('isCommandActivity: handles new messages and message edits only', () => {
  assert.equal(isCommandActivity({ type: 'message' }), true);
  assert.equal(isCommandActivity({ type: 'messageUpdate', channelData: { eventType: 'editMessage' } }), true);
  for (const activity of [
    { type: 'messageUpdate', channelData: { eventType: 'undeleteMessage' } },
    { type: 'messageUpdate' },
    { type: 'messageDelete', channelData: { eventType: 'softDeleteMessage' } },
    { type: 'conversationUpdate' },
    { type: 'messageReaction' },
    null,
  ]) {
    assert.equal(isCommandActivity(activity), false, JSON.stringify(activity));
  }
});

test('emailsToRequest: a new message files every email without a lookup', async () => {
  const requested = async () => { throw new Error('should not look up'); };
  const emails = await emailsToRequest({}, { type: 'message', id: '1', conversation: { id: CONV } }, 'add', ['a@example.com'], { requested });
  assert.deepEqual(emails, ['a@example.com']);
});

test('emailsToRequest: an edit skips emails this message already requested', async () => {
  const calls = [];
  const requested = async (db, query) => {
    calls.push(query);
    return new Set(['a@example.com']);
  };
  const activity = { type: 'messageUpdate', id: '1795439226449', conversation: { id: CONV } };
  const emails = await emailsToRequest({ DB: {} }, activity, 'add', ['A@Example.com', 'b@example.com'], { requested });
  assert.deepEqual(emails, ['b@example.com']);
  assert.deepEqual(calls, [{ conversationId: CONV, teamsMessageId: '1795439226449', action: 'add' }]);
});

test('emailsToRequest: an edit of a message the bot never handled files everything', async () => {
  const requested = async () => new Set();
  const activity = { type: 'messageUpdate', id: '1795439226449', conversation: { id: CONV } };
  assert.deepEqual(
    await emailsToRequest({ DB: {} }, activity, 'add', ['kasingh@example.com'], { requested }),
    ['kasingh@example.com'],
  );
});
