/**
 * URL sanitizer for the ProseMirror schema boundary.
 *
 * Hrefs entering the model from pasted/parsed HTML are attacker-controlled.
 * This allowlist neutralizes javascript:, vbscript:, data: and other XSS
 * vectors before an href is stored on a `link` mark or serialized to an anchor.
 */

/** Schemes permitted for hyperlink hrefs. */
const SAFE_HREF_SCHEMES = ['http:', 'https:', 'mailto:', 'tel:'];

// ASCII whitespace + C0 control characters (0x00–0x20). Browsers ignore these
// inside a URL scheme, so `java\tscript:` runs as `javascript:`; strip them
// before comparing the scheme against the allowlist.
// eslint-disable-next-line no-control-regex -- stripping control chars is the point
const CONTROL_CHARS = /[\u0000-\u0020]/g;

/**
 * Sanitize a hyperlink href.
 *
 * Allows http:, https:, mailto:, tel:, and relative/anchor URLs. Rejects
 * javascript:, vbscript:, data:, and any other scheme by returning ''.
 */
export function sanitizeHref(href: string | null | undefined): string {
  if (!href) return '';
  const value = String(href).trim();
  if (value === '') return '';
  const colon = value.indexOf(':');
  if (colon === -1) return value; // relative / anchor / query — no scheme
  const beforeSlash = value.search(/[/?#]/);
  // A ':' after the first '/', '?' or '#' is part of the path, not a scheme.
  if (beforeSlash !== -1 && beforeSlash < colon) return value;
  const scheme = value.slice(0, colon + 1).replace(CONTROL_CHARS, '').toLowerCase();
  return SAFE_HREF_SCHEMES.includes(scheme) ? value : '';
}

/** Schemes permitted for image sources (plus a `data:image/` special case). */
const SAFE_IMAGE_SCHEMES = ['http:', 'https:'];

/**
 * Sanitize an image `src`. Allows http:, https:, `data:image/…`, and relative
 * URLs; rejects javascript: and everything else by returning ''. Prevents a
 * hidden editor view from firing a network request to an attacker URL.
 */
export function sanitizeImageSrc(src: string | null | undefined): string {
  if (!src) return '';
  const value = String(src).trim();
  if (value === '') return '';
  const colon = value.indexOf(':');
  if (colon === -1) return value; // relative / protocol-relative
  const beforeSlash = value.search(/[/?#]/);
  if (beforeSlash !== -1 && beforeSlash < colon) return value;
  const scheme = value.slice(0, colon + 1).replace(CONTROL_CHARS, '').toLowerCase();
  if (SAFE_IMAGE_SCHEMES.includes(scheme)) return value;
  if (/^data:image\//i.test(value.replace(CONTROL_CHARS, ''))) return value;
  return '';
}

const HEX_COLOR = /^#(?:[0-9a-fA-F]{3,4}|[0-9a-fA-F]{6}|[0-9a-fA-F]{8})$/;
// A whitelisted color function or var() reference over a safe character set.
// No leading `\s*` (avoids ReDoS — the class already matches whitespace).
const FN_COLOR =
  /^(?:rgb|rgba|hsl|hsla|hwb|lab|lch|oklab|oklch|color|color-mix|var)\([a-z0-9.,%/()#_\s-]+\)$/i;
const KEYWORD_COLOR = /^[a-zA-Z]+$/;

/**
 * Return `color` if it is a safe CSS color token, otherwise undefined. Blocks
 * CSS injection such as `red; background: url(//evil)` at the toDOM boundary.
 */
export function safeCssColor(color: string | null | undefined): string | undefined {
  if (!color) return undefined;
  const value = String(color).trim();
  if (value === '') return undefined;
  if (/url\(/i.test(value)) return undefined;
  if (HEX_COLOR.test(value) || FN_COLOR.test(value) || KEYWORD_COLOR.test(value)) {
    return value;
  }
  return undefined;
}

// font-family: letters, digits, spaces, commas, quotes, hyphens. Excludes the
// characters (';{}():') needed to break out of the declaration.
const SAFE_FONT_FAMILY = /^[a-zA-Z0-9\s,'"-]+$/;

/**
 * Return `value` if it is a safe CSS `font-family` list, otherwise undefined.
 */
export function safeFontFamily(value: string | null | undefined): string | undefined {
  if (!value) return undefined;
  const v = String(value).trim();
  if (v === '') return undefined;
  // Reject an actual url() call, not any name merely containing "url" (e.g. the
  // legitimate font "Curlz MT"). The char-class allowlist already forbids the
  // '(' a real url() needs, so this is just defense in depth.
  if (/url\(/i.test(v)) return undefined;
  return SAFE_FONT_FAMILY.test(v) ? v : undefined;
}

// A font NAME kept as data (the formatting before a tracked change), never put into CSS: any Unicode letter, digit or
// space and a few punctuation marks (so CJK names pass), none of the characters that could break out of a declaration.
const SAFE_FONT_NAME = /^[\p{L}\p{N}\p{M}\s.,'"&_-]+$/u;
const MAX_FONT_NAME = 100;

/**
 * Return `value` if it is usable as a stored font name, otherwise undefined. Unlike `safeFontFamily` it accepts non-ASCII
 * names, so a tracked formatting change keeps the font it came from; only turn the result into CSS through `safeFontFamily`.
 */
export function safeFontName(value: string | null | undefined): string | undefined {
  if (!value) return undefined;
  const v = String(value).trim();
  if (v === '' || v.length > MAX_FONT_NAME) return undefined;
  if (/url\(/i.test(v)) return undefined;
  return SAFE_FONT_NAME.test(v) ? v : undefined;
}

// The only keywords ever emitted for `text-align`.
const ALIGN_KEYWORDS = new Set(['left', 'right', 'center', 'justify', 'start', 'end']);

/**
 * Return a whitelisted CSS alignment keyword (lower-cased), or undefined.
 * Prevents `left; background: url(//evil)`-style injection at alignment sinks.
 */
export function safeCssKeyword(value: string | null | undefined): string | undefined {
  if (!value) return undefined;
  const v = String(value).trim().toLowerCase();
  return ALIGN_KEYWORDS.has(v) ? v : undefined;
}

/**
 * Coerce a value to a finite CSS length in px, or undefined. A number can never
 * carry a CSS-injection payload, so this neutralizes attrs like spaceBefore /
 * spaceAfter / fontSize that are interpolated into a style string.
 */
export function safeCssNumber(value: unknown): number | undefined {
  const n = typeof value === 'number' ? value : parseFloat(String(value));
  return Number.isFinite(n) ? n : undefined;
}
