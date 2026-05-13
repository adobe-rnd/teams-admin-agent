const EMAIL_RE = /[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}/g;

const INTENT_WORDS = /\b(add|invite|include|onboard|grant|give\s+access|join)\b/i;

// Common TLDs used to separate emails that Teams flattens together when
// users send bullet lists. We insert a space after these when they are
// immediately followed by another letter (which can only happen when two
// emails were concatenated).
const COMMON_TLDS = 'com|org|net|edu|gov|mil|io|co|us|uk|de|fr|jp|cn|au|in|br|ca|me|tv|info|biz|app|dev|ai|cloud';
const TLD_GLUE_RE = new RegExp(`\\.(${COMMON_TLDS})(?=[a-zA-Z])`, 'gi');

const HTML_ENTITIES = { lt: '<', gt: '>', amp: '&', quot: '"', apos: "'", nbsp: ' ' };

function decodeEntities(text) {
  return text.replace(/&(lt|gt|amp|quot|apos|nbsp|#\d+);/gi, (m, name) => {
    if (name.startsWith('#')) return String.fromCharCode(Number(name.slice(1)));
    return HTML_ENTITIES[name.toLowerCase()] ?? m;
  });
}

// Decode HTML entities, strip <at>...</at> mentions entirely, replace
// every remaining HTML tag with a space, then split emails that Teams
// concatenated together (no separator between bullet list items).
function stripMarkup(text) {
  return decodeEntities(text)
    .replace(/<at[^>]*>.*?<\/at>/gi, ' ')
    .replace(/<[^>]+>/g, ' ')
    .replace(TLD_GLUE_RE, '.$1 ');
}

export function hasAddIntent(text) {
  if (!text) return false;
  return INTENT_WORDS.test(stripMarkup(text));
}

/**
 * Extract de-duplicated email addresses from a Teams message.
 * Pulls emails from mailto: links first (when chips are present), then
 * from the stripped text (with concatenated bullet-list emails split).
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
