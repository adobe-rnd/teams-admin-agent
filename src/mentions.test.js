import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mentionedPeople, mentionEmails, noAddressText } from './mentions.js';
import { hasAddIntent, getCommandText } from './parser.js';

const bot = { type: 'mention', text: '<at>admin-bot</at>', mentioned: { id: 'bot-id', name: 'admin-bot' } };
const jane = { type: 'mention', text: '<at>Jane Doe</at>', mentioned: { id: '29:jane', name: 'Jane Doe' } };

test('mentionedPeople: the people mentioned after the bot, not the bot', () => {
  const text = '<at>admin-bot</at> add <at>Jane Doe</at>';
  assert.deepEqual(mentionedPeople(text, [bot, jane], 'bot-id'), [
    { id: '29:jane', name: 'Jane Doe', aadObjectId: null },
  ]);
  // The command still reads as an add.
  assert.equal(hasAddIntent(getCommandText(text, [bot, jane], 'bot-id')), true);
});

test('mentionedPeople: nobody mentioned before the command, each person once', () => {
  const chamu = { type: 'mention', text: '<at>Chamu</at>', mentioned: { id: '29:chamu', name: 'Chamu' } };
  const text = '<at>Chamu</at> see below. <at>admin-bot</at> add <at>Jane Doe</at> and <at>Jane Doe</at>';
  const people = mentionedPeople(text, [chamu, bot, jane, jane], 'bot-id');
  assert.deepEqual(people.map((p) => p.id), ['29:jane']);
});

test('mentionedPeople: keeps an Entra object id when the mention has one', () => {
  const withId = { ...jane, mentioned: { ...jane.mentioned, aadObjectId: 'aad-1' } };
  const [p] = mentionedPeople('<at>admin-bot</at> add <at>Jane Doe</at>', [bot, withId], 'bot-id');
  assert.equal(p.aadObjectId, 'aad-1');
});

test('mentionedPeople: nothing without a command', () => {
  assert.deepEqual(mentionedPeople('<at>Jane Doe</at> hello', [jane], 'bot-id'), []);
  assert.deepEqual(mentionedPeople('', [], 'bot-id'), []);
});

test('mentionEmails: found addresses, and the names of whom none was found', async () => {
  const people = [
    { id: '1', name: 'Jane Doe' }, { id: '2', name: 'John Roe' }, { id: '3', name: 'Fails' },
  ];
  const find = async (env, p) => {
    if (p.id === '3') throw new Error('graph down');
    return p.id === '1' ? 'Jane@Example.com' : null;
  };
  assert.deepEqual(await mentionEmails({}, people, find), {
    emails: ['jane@example.com'], unresolved: ['John Roe', 'Fails'],
  });
});

test('noAddressText: names whom no address was found for, or asks for addresses', () => {
  assert.match(noAddressText('add', ['Jane Doe']), /\*\*Jane Doe\*\*.*`@admin-bot add name@example.com`/);
  assert.match(noAddressText('remove', []), /email address of each person to remove/);
});
