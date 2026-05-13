const EMAIL_RE = /[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}/g;

const INTENT_WORDS = /\b(add|invite|include|onboard|grant|give\s+access|join)\b/i;

// Strip <at>...</at> mentions entirely, then replace every remaining HTML
// tag with a space so adjacent <li> / <a> / <p> boundaries become real
// separators for the email regex.
function stripMarkup(text) {
  return text
    .replace(/<at[^>]*>.*?<\/at>/gi, ' ')
    .replace(/<[^>]+>/g, ' ');
}

export function hasAddIntent(text) {
  if (!text) return false;
  return INTENT_WORDS.test(stripMarkup(text));
}

/**
 * Extract de-duplicated email addresses from a Teams message.
 * Pulls emails from mailto: links first (so bullet lists that get
 * flattened don't lose their separators), then from the stripped text.
 */
export function extractEmails(text) {
  if (!text) return [];
  const found = [];
  for (const m of text.matchAll(/mailto:([^"'>\s,;]+@[^"'>\s,;]+)/gi)) {
    found.push(m[1]);
  }
  const cleaned = stripMarkup(text);
  const inline = cleaned.match(EMAIL_RE) ?? [];
  found.push(...inline);
  return [...new Set(found.map((e) => e.toLowerCase()))];
}
