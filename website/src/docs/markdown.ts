import { marked } from 'marked';
import { hasDoc } from './manifest';

const REPO_TREE = 'https://github.com/b32n/spindle/tree/master/documentation';

export interface TocItem {
  id: string;
  text: string;
  level: number;
}

const slugify = (s: string): string =>
  s
    .toLowerCase()
    .trim()
    .replace(/[^\w\s-]/g, '')
    .replace(/\s+/g, '-')
    .replace(/-+/g, '-');

// Resolve a relative markdown link (from inside `currentSlug`) to an in-app
// docs route. Returns null when it can't be mapped to a known doc page.
function resolveDocLink(href: string, currentSlug: string): string | null {
  const [pathPart, hash = ''] = href.split('#');
  if (!/\.md$/i.test(pathPart)) return null;

  const dir = currentSlug.includes('/') ? currentSlug.replace(/\/[^/]*$/, '') : '';
  const segs = dir ? dir.split('/') : [];
  for (const seg of pathPart.replace(/\.md$/i, '').split('/')) {
    if (seg === '' || seg === '.') continue;
    if (seg === '..') segs.pop();
    else segs.push(seg);
  }
  let target = segs.join('/');
  if (target === 'README') target = '';
  if (!hasDoc(target)) return null;
  return `/docs${target ? '/' + target : ''}${hash ? '#' + hash : ''}`;
}

// Parse markdown to HTML, assigning stable heading ids, collecting a TOC,
// promoting mermaid code blocks to renderable containers, and rewriting links:
// internal .md links become SPA routes, everything else opens externally
// (unmapped repo-relative links fall back to GitHub).
export function renderDoc(
  raw: string,
  currentSlug: string,
): { html: string; toc: TocItem[]; hasMermaid: boolean } {
  const rawHtml = marked.parse(raw, { async: false, gfm: true }) as string;
  const doc = new DOMParser().parseFromString(rawHtml, 'text/html');

  // ```mermaid blocks render as <pre><code class="language-mermaid">…</code></pre>.
  // Swap each for a <div class="mermaid"> holding the raw graph source so the
  // mermaid runtime can turn it into an SVG.
  let hasMermaid = false;
  doc.querySelectorAll('pre > code.language-mermaid').forEach((code) => {
    const div = doc.createElement('div');
    div.className = 'mermaid';
    div.textContent = code.textContent ?? '';
    code.parentElement!.replaceWith(div);
    hasMermaid = true;
  });

  const toc: TocItem[] = [];
  const seen = new Map<string, number>();
  doc.querySelectorAll('h1, h2, h3').forEach((h) => {
    const level = Number(h.tagName[1]);
    const text = h.textContent ?? '';
    let id = slugify(text) || 'section';
    const n = seen.get(id) ?? 0;
    seen.set(id, n + 1);
    if (n) id = `${id}-${n}`;
    h.id = id;
    if (level === 2 || level === 3) toc.push({ id, text, level });
  });

  doc.querySelectorAll('a[href]').forEach((a) => {
    const href = a.getAttribute('href') ?? '';
    if (href.startsWith('#') || href.startsWith('mailto:')) return;
    if (/^(https?:)?\/\//.test(href)) {
      a.setAttribute('target', '_blank');
      a.setAttribute('rel', 'noopener noreferrer');
      return;
    }
    const internal = resolveDocLink(href, currentSlug);
    if (internal) {
      a.setAttribute('href', internal);
    } else {
      // Repo-relative link we don't render on the site (e.g. examples/); send
      // it to the source on GitHub.
      const clean = href.replace(/^\.\//, '');
      a.setAttribute('href', `${REPO_TREE}/${clean}`);
      a.setAttribute('target', '_blank');
      a.setAttribute('rel', 'noopener noreferrer');
    }
  });

  return { html: doc.body.innerHTML, toc, hasMermaid };
}
