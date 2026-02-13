const EMAIL_RE = /[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}/g;

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
