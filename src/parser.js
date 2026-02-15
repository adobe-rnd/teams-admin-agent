const EMAIL_RE = /[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}/g;

const INTENT_WORDS = /\b(add|invite|include|onboard|grant|give\s+access|join)\b/i;

/**
 * Check whether the message expresses intent to add members.
 * Strips <at>…</at> mention markup first.
 */
export function hasAddIntent(text) {
  if (!text) return false;
  const cleaned = text.replace(/<at[^>]*>.*?<\/at>/gi, ' ');
  return INTENT_WORDS.test(cleaned);
}

/**
 * Extract de-duplicated email addresses from a Teams message.
 * Strips <at>…</at> mention markup first so the bot's own name
 * isn't accidentally matched.
 */
export function extractEmails(text) {
  if (!text) return [];
  const cleaned = text.replace(/<at[^>]*>.*?<\/at>/gi, ' ');
  const matches = cleaned.match(EMAIL_RE);
  if (!matches) return [];
  return [...new Set(matches.map((e) => e.toLowerCase()))];
}
