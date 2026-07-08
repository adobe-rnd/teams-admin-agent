import { test } from 'node:test';
import assert from 'node:assert/strict';
import { extractEmails, hasAddIntent, hasRemoveIntent } from './parser.js';

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

// ── hasRemoveIntent ───────────────────────────────────────────────

test('hasRemoveIntent: matches remove/delete/uninvite/revoke/kick verbs', () => {
  assert.equal(hasRemoveIntent('please remove alice@example.com'), true);
  assert.equal(hasRemoveIntent('delete Bob from the team'), true);
  assert.equal(hasRemoveIntent('uninvite carol@example.com'), true);
  assert.equal(hasRemoveIntent('revoke access for dan@example.com'), true);
  assert.equal(hasRemoveIntent('kick this user'), true);
});

test('hasRemoveIntent: an add request is not a remove request', () => {
  assert.equal(hasRemoveIntent('please add alice@example.com'), false);
  assert.equal(hasRemoveIntent('invite Bob to the team'), false);
});

test('hasRemoveIntent: returns false for empty/missing text', () => {
  assert.equal(hasRemoveIntent(''), false);
  assert.equal(hasRemoveIntent(null), false);
  assert.equal(hasRemoveIntent(undefined), false);
});

test('"uninvite" reads as remove, not add', () => {
  // No word boundary before "invite" inside "uninvite", so it must not
  // trip the add detector while it does trip the remove detector.
  assert.equal(hasAddIntent('uninvite carol@example.com'), false);
  assert.equal(hasRemoveIntent('uninvite carol@example.com'), true);
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

test('extractEmails: does not split a local part containing a common-TLD substring', () => {
  // Regression: "Vanessa.CalvoBarreto@..." has ".ca" (a common TLD) inside the
  // local part with a letter+@ following, so the glue splitter fired there and
  // dropped the "Vanessa.Ca" prefix. Glue splitting must only fire on a .tld
  // that ends a domain (after @), never inside a local part.
  assert.deepEqual(
    extractEmails('add Vanessa.CalvoBarreto@terumo-europe.com'),
    ['vanessa.calvobarreto@terumo-europe.com'],
  );
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

// ── extractEmails: separator support ──────────────────────────────

test('extractEmails: comma-separated inline list', () => {
  assert.deepEqual(
    extractEmails('add alice@example.com, bob@example.com, carol@example.com'),
    ['alice@example.com', 'bob@example.com', 'carol@example.com'],
  );
});

test('extractEmails: comma without surrounding space', () => {
  assert.deepEqual(
    extractEmails('alice@example.com,bob@example.com'),
    ['alice@example.com', 'bob@example.com'],
  );
});

test('extractEmails: semicolon-separated inline list', () => {
  assert.deepEqual(
    extractEmails('add alice@example.com; bob@example.com; carol@example.com'),
    ['alice@example.com', 'bob@example.com', 'carol@example.com'],
  );
});

test('extractEmails: whitespace separators (space, tab, newline, CRLF)', () => {
  assert.deepEqual(
    extractEmails('alice@example.com bob@example.com\tcarol@example.com\ndan@example.com\r\neve@example.com'),
    [
      'alice@example.com',
      'bob@example.com',
      'carol@example.com',
      'dan@example.com',
      'eve@example.com',
    ],
  );
});

// ── extractEmails: address format support ─────────────────────────

test('extractEmails: plus-tag in local part', () => {
  assert.deepEqual(
    extractEmails('add alice+team@example.com, bob+y@example.com'),
    ['alice+team@example.com', 'bob+y@example.com'],
  );
});

test('extractEmails: dots, underscores, hyphens, digits, percent in local part', () => {
  const text = 'alice.smith@example.com alice_smith@example.com alice-smith@example.com alice123@example.com alice%test@example.com';
  assert.deepEqual(extractEmails(text), [
    'alice.smith@example.com',
    'alice_smith@example.com',
    'alice-smith@example.com',
    'alice123@example.com',
    'alice%test@example.com',
  ]);
});

test('extractEmails: hyphenated and multi-label domains', () => {
  // Subdomain, hyphenated domain, digits in domain.
  const text = 'alice@mail.corp.example.com bob@my-company.example carol@example123.com';
  assert.deepEqual(extractEmails(text), [
    'alice@mail.corp.example.com',
    'bob@my-company.example',
    'carol@example123.com',
  ]);
});

test('extractEmails: multi-level country TLDs (.co.uk, .com.au)', () => {
  assert.deepEqual(
    extractEmails('alice@example.co.uk bob@example.com.au'),
    ['alice@example.co.uk', 'bob@example.com.au'],
  );
});

test('extractEmails: glued multi-level TLD splits at the trailing common TLD', () => {
  // When two .co.uk addresses get concatenated, the split fires at .uk
  // (a common TLD) — not at the .co — because the .co lookahead requires
  // [a-zA-Z] next, which fails against a literal dot.
  assert.deepEqual(
    extractEmails('alice@example.co.ukbob@example.com'),
    ['alice@example.co.uk', 'bob@example.com'],
  );
});

test('extractEmails: uppercase MAILTO: scheme', () => {
  assert.deepEqual(
    extractEmails('<a HREF="MAILTO:alice@example.com">Alice</a>'),
    ['alice@example.com'],
  );
});

test('extractEmails: mixed mailto chip and RFC-5322 inline in one message', () => {
  const text = '<a href="mailto:alice@example.com">Alice</a>; Bob Sample <bob@example.com>';
  assert.deepEqual(extractEmails(text), [
    'alice@example.com',
    'bob@example.com',
  ]);
});

// ── extractEmails: surrounding punctuation ────────────────────────

test('extractEmails: sentence punctuation and brackets around an address', () => {
  // Trailing period (end of sentence) and wrapping parens should not be
  // captured as part of the address.
  assert.deepEqual(extractEmails('Please add alice@example.com.'), ['alice@example.com']);
  assert.deepEqual(extractEmails('(alice@example.com)'), ['alice@example.com']);
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
