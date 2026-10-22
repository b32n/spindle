/**
 * Where the editors get web fonts from. The editors never talk to a third-party
 * font host on their own: by default they ask the Spindle font mirror, and a
 * host can switch to its own stylesheet endpoint, a custom loader, or
 * `{ kind: 'local' }` for zero font requests (system fonts only).
 *
 * Every font stack the editors emit ends in a generic family, so text stays
 * readable if a font (or the whole source) is unreachable.
 */
export type FontSource =
  /** No font requests at all; only fonts installed on the user's system render. */
  | { kind: 'local' }
  /**
   * A stylesheet endpoint compatible with the Google Fonts `css2` API
   * (`<url>?family=Roboto:wght@400&family=Lato:wght@400&display=swap`), e.g. a
   * self-hosted mirror or Bunny Fonts (`https://fonts.bunny.net/css2`).
   */
  | { kind: 'css'; url: string }
  /** Load the fonts yourself (inject @font-face, call a service, ...). */
  | { kind: 'custom'; load: (families: string[]) => void | Promise<void> };

/** The font mirror used when no `fontSource` is given. */
export const DEFAULT_FONT_SOURCE_URL = 'https://fonts.b32n.io/css2';
export const DEFAULT_FONT_SOURCE: FontSource = { kind: 'css', url: DEFAULT_FONT_SOURCE_URL };

// Families that ship with common operating systems (or are generic keywords):
// never worth a network request.
const SYSTEM_FAMILIES = new Set([
  'arial', 'helvetica', 'helvetica neue', 'georgia', 'times new roman', 'times', 'verdana', 'tahoma',
  'trebuchet ms', 'courier new', 'courier', 'impact', 'comic sans ms', 'segoe ui', 'calibri', 'cambria',
  'consolas', 'monaco', 'menlo', 'palatino', 'garamond', 'system-ui', 'sans-serif', 'serif', 'monospace',
  'cursive', 'fantasy', 'major', 'minor',
]);

export function isSystemFont(family: string): boolean {
  return SYSTEM_FAMILIES.has(family.trim().toLowerCase());
}

/** The stylesheet URL for `families` at `baseUrl`, or null if `baseUrl` is not http(s)/relative. */
export function fontStylesheetUrl(baseUrl: string, families: string[]): string | null {
  if (/^\s*[a-z][a-z0-9+.-]*:/i.test(baseUrl) && !/^\s*https?:/i.test(baseUrl)) return null;
  const query = families
    .map((f) => `family=${encodeURIComponent(f.trim()).replace(/%20/g, '+')}:wght@400`)
    .join('&');
  return `${baseUrl}${baseUrl.includes('?') ? '&' : '?'}${query}&display=swap`;
}

/** Minimal slice of `Document` the loader needs (lets tests run without a DOM). */
export interface FontLoaderDocument {
  head: { appendChild(node: unknown): unknown };
  createElement(tag: 'link'): {
    rel: string; href: string; onerror: (() => void) | null;
    setAttribute(name: string, value: string): void; remove(): void;
  };
}

/**
 * Returns a function that makes sure the given families are loaded from
 * `source`. Each family is requested at most once per loader; unreachable
 * stylesheets are dropped silently so the fallback stack keeps working.
 */
export function createFontLoader(
  source: FontSource = DEFAULT_FONT_SOURCE,
  doc: FontLoaderDocument | undefined = typeof document === 'undefined' ? undefined : (document as unknown as FontLoaderDocument),
): (families: string[]) => void {
  const requested = new Set<string>();
  return (families) => {
    if (source.kind === 'local') return;
    const fresh = [...new Set(families.map((f) => f.trim()))].filter((f) => f && !isSystemFont(f) && !requested.has(f));
    if (fresh.length === 0) return;
    fresh.forEach((f) => requested.add(f));
    if (source.kind === 'custom') {
      try {
        Promise.resolve(source.load(fresh)).catch(() => { /* fall back to the font stack */ });
      } catch { /* fall back to the font stack */ }
      return;
    }
    if (!doc) return;
    for (let i = 0; i < fresh.length; i += 10) {
      const batch = fresh.slice(i, i + 10);
      const href = fontStylesheetUrl(source.url, batch);
      if (!href) return;
      const link = doc.createElement('link');
      link.rel = 'stylesheet';
      link.href = href;
      link.setAttribute('data-spindle-fonts', batch.join(','));
      link.onerror = () => link.remove();
      doc.head.appendChild(link);
    }
  };
}
