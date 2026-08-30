// Turning workbook prose into renderable pieces.

const URL_PATTERN = /\bhttps?:\/\/[^\s<>]+/gi;

const count = (s, ch) => (s.split(ch).length - 1);

/**
 * Splits text into plain and linkable runs:
 *
 *   linkify('see https://x.test/a for more')
 *     -> [{ text: 'see ' }, { text: 'https://x.test/a', href: '…' }, { text: ' for more' }]
 *
 * The caller renders each run itself, so the prose is never handed to set:html
 * — Astro escapes the text runs as usual and only the href we built is trusted.
 *
 * Workbook cells write URLs inline mid-sentence ("…sequences: https://…"), so
 * the tricky part is where the URL stops. Trailing sentence punctuation is
 * given back to the text; a closing bracket is only given back when the URL
 * didn't open one, which keeps links that legitimately contain brackets whole.
 */
export function linkify(text) {
  const source = String(text ?? '');
  const parts = [];
  let cursor = 0;

  for (const match of source.matchAll(URL_PATTERN)) {
    let url = match[0];

    for (let trimming = true; trimming && url; ) {
      const last = url[url.length - 1];
      if ('.,;:!?"\''.includes(last)) url = url.slice(0, -1);
      else if (last === ')' && count(url, ')') > count(url, '(')) url = url.slice(0, -1);
      else if (last === ']' && count(url, ']') > count(url, '[')) url = url.slice(0, -1);
      else trimming = false;
    }

    if (!url) continue;

    if (match.index > cursor) parts.push({ text: source.slice(cursor, match.index) });
    parts.push({ text: url, href: url });
    cursor = match.index + url.length;
  }

  if (cursor < source.length) parts.push({ text: source.slice(cursor) });
  return parts;
}
