import { test } from 'node:test';
import assert from 'node:assert/strict';
import { extractEmails, hasAddIntent } from './parser.js';

// ── hasAddIntent ──────────────────────────────────────────────────

test('hasAddIntent: matches add/invite/include verbs', () => {
  assert.equal(hasAddIntent('please add alice@example.com'), true);
  assert.equal(hasAddIntent('invite Bob to the team'), true);
  assert.equal(hasAddIntent('include this user'), true);
  assert.equal(hasAddIntent('onboard the new hire'), true);
});

test('hasAddIntent: ignores text inside <at>...</at> mention tags', () => {
  // Bot name might contain "add" — must not count as intent on its own
  assert.equal(hasAddIntent('<at>add-bot</at> hello there'), false);
});

test('hasAddIntent: returns false for empty/missing text', () => {
  assert.equal(hasAddIntent(''), false);
  assert.equal(hasAddIntent(null), false);
  assert.equal(hasAddIntent(undefined), false);
});

// ── extractEmails: mailto chips (commit 22f49c6) ──────────────────

test('extractEmails: pulls addresses from mailto: hrefs in flattened bullet lists', () => {
  // Teams renders contact chips as <a href="mailto:..."> — the email regex
  // alone would glue list items together, so mailto extraction runs first.
  const html = `<at>admin-bot</at> add the following:
    <li><a href="mailto:alice@example.com">Alice</a></li>
    <li><a href="mailto:bob@example.com">Bob</a></li>
    <li><a href="mailto:carol@example.org">Carol</a></li>`;
  assert.deepEqual(extractEmails(html), [
    'alice@example.com',
    'bob@example.com',
    'carol@example.org',
  ]);
});

test('extractEmails: dedupes when the same address appears in both mailto and link text', () => {
  const html = `<a href="mailto:alice@example.com">alice@example.com</a>`;
  assert.deepEqual(extractEmails(html), ['alice@example.com']);
});

test('extractEmails: lowercases extracted addresses', () => {
  assert.deepEqual(extractEmails('add ALICE@Example.COM'), ['alice@example.com']);
});

// ── extractEmails: concatenated bullet emails (commit 74a0fc4) ────

test('extractEmails: splits emails Teams glued together at common TLD boundaries', () => {
  // Teams sometimes flattens <li> items with no separator at all, producing
  // "a@example.comb@example.com". A glued common-TLD-then-letter boundary
  // can only come from this concatenation, so we split there.
  const glued = 'add alice@example.combob@example.orgcarol@example.net';
  assert.deepEqual(extractEmails(glued), [
    'alice@example.com',
    'bob@example.org',
    'carol@example.net',
  ]);
});

test('extractEmails: decodes HTML entities before email extraction', () => {
  // Teams sometimes HTML-encodes the entire payload (&lt;li&gt; instead of <li>).
  // decodeEntities runs first so tag stripping and TLD-glue splitting still work.
  const encoded = 'add &lt;li&gt;alice@example.com&lt;/li&gt;&lt;li&gt;bob@example.com&lt;/li&gt;';
  assert.deepEqual(extractEmails(encoded), [
    'alice@example.com',
    'bob@example.com',
  ]);
});

// ── extractEmails: end-of-string TLD (commit 865b472) ─────────────

test('extractEmails: does not split a .com TLD at end of string into .co + m', () => {
  // Regression: earlier the lookahead only required a following letter, so
  // the regex backtracked from .com (lookahead fail at EOS) to .co + lone m,
  // producing "alice@example.co". Lookahead now requires another local@.
  assert.deepEqual(extractEmails('add alice@example.com'), ['alice@example.com']);
});

test('extractEmails: preserves trailing dot-co address when nothing follows', () => {
  // .co is itself a common TLD; it must remain intact when at end of string.
  assert.deepEqual(extractEmails('add alice@example.co'), ['alice@example.co']);
});

// ── extractEmails: RFC-5322 angle brackets (commit c8f3ff1) ───────

test('extractEmails: extracts addresses wrapped in <...> mailbox brackets', () => {
  // Outlook copy-paste produces "Name <email@domain>". The HTML tag stripper
  // used to consume the <...> span entirely; it now skips <...> spans with @.
  const text = 'add Alice Example <alice@example.com>';
  assert.deepEqual(extractEmails(text), ['alice@example.com']);
});

test('extractEmails: handles a semicolon-separated RFC-5322 list (the real failure case)', () => {
  // The exact message shape that triggered the silent-fail incident, with
  // names and domains anonymized.
  const text = `add the following email addressees to this chat

Alice Example <alice@example.com>; Bob Sample <bob.sample@partner.example>; Carol Demo <carol@example.com>; Dan Placeholder <dan_p@example.com>; Eve Tester <eve@example.com>`;
  assert.deepEqual(extractEmails(text), [
    'alice@example.com',
    'bob.sample@partner.example',
    'carol@example.com',
    'dan_p@example.com',
    'eve@example.com',
  ]);
});

test('extractEmails: handles HTML-encoded angle-bracket addresses', () => {
  const text = 'add Alice &lt;alice@example.com&gt;; Bob &lt;bob@example.com&gt;';
  assert.deepEqual(extractEmails(text), [
    'alice@example.com',
    'bob@example.com',
  ]);
});

// ── extractEmails: misc invariants ────────────────────────────────

test('extractEmails: strips <at> mention so bot name is not matched', () => {
  // The bot's own @mention markup includes its display name — must not be
  // parsed as an email even when it superficially resembles one.
  const text = '<at>admin-bot</at> add alice@example.com';
  assert.deepEqual(extractEmails(text), ['alice@example.com']);
});

test('extractEmails: returns empty array on empty/missing input', () => {
  assert.deepEqual(extractEmails(''), []);
  assert.deepEqual(extractEmails(null), []);
  assert.deepEqual(extractEmails(undefined), []);
});
