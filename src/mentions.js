/**
 * People named by an @mention in an @admin-bot command, instead of by their email address.
 *
 * Teams completes an @mention from the whole tenant's directory, so "@admin-bot add @Jane Doe"
 * reads as a complete request to the person who writes it. The message has no address for
 * extractEmails, though: the mention is an `<at>Jane Doe</at>` tag and an entity with Jane's
 * Teams id. Each such person's address is looked up in Graph here; whom none is found for, the
 * bot says so in its answer instead of staying silent.
 */
import { getCommandText } from './parser.js';

const MAX_MENTIONS = 20;

/**
 * The people `entities` mention in the command of `text` (what follows the bot's own mention),
 * each once, as `{ id, name, aadObjectId }`: not the bot (`botId`), and nobody mentioned before
 * the command, whom the message talks to rather than about.
 */
export function mentionedPeople(text, entities = [], botId) {
  const command = getCommandText(text, entities, botId);
  if (!command) return [];
  const seen = new Set();
  const out = [];
  for (const e of entities || []) {
    const who = e && e.type === 'mention' ? e.mentioned : null;
    if (!who || !who.id || who.id === botId || seen.has(who.id)) continue;
    if (!e.text || !command.includes(e.text)) continue;
    seen.add(who.id);
    out.push({
      id: who.id,
      name: String(who.name || e.text.replace(/<[^>]+>/g, '')).trim(),
      aadObjectId: who.aadObjectId || null,
    });
    if (out.length >= MAX_MENTIONS) break;
  }
  return out;
}

/**
 * The mail addresses of `people` (from mentionedPeople), through `findEmail(env, person)`,
 * which gives an address or null: `{ emails, unresolved }`, `unresolved` the names of those
 * without one. A lookup that fails counts as none found.
 */
export async function mentionEmails(env, people, findEmail) {
  const emails = [];
  const unresolved = [];
  for (const p of people) {
    let email = null;
    try {
      // eslint-disable-next-line no-await-in-loop
      email = await findEmail(env, p);
    } catch (err) {
      console.error(`Mention lookup failed for ${p.id}:`, err.message);
    }
    if (email) {
      emails.push(email.toLowerCase());
    } else {
      unresolved.push(p.name || 'someone');
    }
  }
  return { emails, unresolved };
}

/**
 * What admin-bot answers about people it has no address for (`unresolved`, their names), or
 * when the command named nobody at all (`unresolved` empty), for the action asked for.
 */
export function noAddressText(action, unresolved = []) {
  const verb = action === 'remove' ? 'remove' : 'add';
  const example = `\`@admin-bot ${verb} name@example.com\``;
  if (!unresolved.length) {
    return `I need the email address of each person to ${verb}, for example ${example}.`;
  }
  const names = unresolved.map((n) => `**${n}**`).join(', ');
  return `I couldn't find the email address of ${names}, so I requested nothing for them. Please write it out, for example ${example}.`;
}
