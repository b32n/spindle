import { createFontLoader, fontStylesheetUrl, DEFAULT_FONT_SOURCE, FontLoaderDocument } from './fonts';

function fakeDoc() {
  const links: Array<{ href: string; rel: string; onerror: (() => void) | null; removed: boolean; attrs: Record<string, string> }> = [];
  const doc: FontLoaderDocument = {
    head: { appendChild: () => undefined },
    createElement: () => {
      const l = {
        href: '', rel: '', onerror: null as (() => void) | null, removed: false, attrs: {} as Record<string, string>,
        setAttribute(n: string, v: string) { l.attrs[n] = v; },
        remove() { l.removed = true; },
      };
      links.push(l);
      return l;
    },
  };
  return { doc, links };
}

describe('font source', () => {
  it('defaults to the Spindle mirror', () => {
    expect(DEFAULT_FONT_SOURCE).toEqual({ kind: 'css', url: 'https://fonts.b32n.io/css2' });
  });

  it('builds css2-style URLs', () => {
    expect(fontStylesheetUrl('https://x.test/css2', ['Open Sans', 'Lato'])).toBe(
      'https://x.test/css2?family=Open+Sans:wght@400&family=Lato:wght@400&display=swap');
    expect(fontStylesheetUrl('/fonts?v=1', ['Lato'])).toBe('/fonts?v=1&family=Lato:wght@400&display=swap');
    expect(fontStylesheetUrl('javascript:alert(1)', ['Lato'])).toBeNull();
  });

  it('local makes no requests', () => {
    const { doc, links } = fakeDoc();
    createFontLoader({ kind: 'local' }, doc)(['Roboto']);
    expect(links).toHaveLength(0);
  });

  it('css requests each web font once, skips system fonts, and drops failed stylesheets', () => {
    const { doc, links } = fakeDoc();
    const load = createFontLoader({ kind: 'css', url: 'https://x.test/css2' }, doc);
    load(['Roboto', 'Arial', 'Georgia']);
    load(['Roboto', 'Lato']);
    expect(links.map((l) => l.href)).toEqual([
      'https://x.test/css2?family=Roboto:wght@400&display=swap',
      'https://x.test/css2?family=Lato:wght@400&display=swap',
    ]);
    links[0].onerror!();
    expect(links[0].removed).toBe(true);
  });

  it('batches in tens', () => {
    const { doc, links } = fakeDoc();
    createFontLoader({ kind: 'css', url: 'https://x.test/css2' }, doc)(Array.from({ length: 23 }, (_, i) => `F${i}`));
    expect(links).toHaveLength(3);
  });

  it('custom loader receives only new web families; its errors are swallowed', async () => {
    const calls: string[][] = [];
    const load = createFontLoader({ kind: 'custom', load: (f) => { calls.push(f); throw new Error('boom'); } });
    load(['Roboto', 'Arial']);
    load(['Roboto']);
    expect(calls).toEqual([['Roboto']]);
    const async = createFontLoader({ kind: 'custom', load: () => Promise.reject(new Error('x')) });
    expect(() => async(['Lato'])).not.toThrow();
  });
});

describe('font loading does not hold up page load', () => {
  it('waits for window load while the document is still loading', () => {
    const { doc, links } = fakeDoc();
    (doc as { readyState?: string }).readyState = 'interactive';
    const listeners: Array<() => void> = [];
    (globalThis as unknown as { window: unknown }).window = {
      addEventListener: (_: string, fn: () => void) => listeners.push(fn),
    };
    try {
      createFontLoader({ kind: 'css', url: 'https://x.test/css2' }, doc)(['Roboto']);
      expect(links).toHaveLength(0);
      listeners.forEach((fn) => fn());
      expect(links).toHaveLength(1);
    } finally {
      delete (globalThis as unknown as { window?: unknown }).window;
    }
  });
});
