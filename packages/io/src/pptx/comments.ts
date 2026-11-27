import type { PptxComment, PptxCommentThread, PptxElement } from '@b32nio/spindle-shared/export/pptx';
import type { OoxmlPackage, Rel } from '../common/ooxml-package';
import { descendants, kid, kids, walk, type XmlElement } from '../xlsx/xml';
import { num } from './units';

// Comments (plan K10c): the legacy format (`ppt/commentAuthors.xml` and one `ppt/comments/*.xml` per slide, `p:cm` with a
// position and replies as `cm`s whose `p15:parentCm` names the comment they answer) and the modern one (`ppt/authors.xml`
// and a `p188:cmLst` per slide, each `p188:cm` anchored to the slide or to a shape, with `p188:reply`s and a status).
// Both become `PptxCommentThread`s of the slide the comments part hangs off.

const UNKNOWN_AUTHOR = 'Unknown author';
/** Legacy positions are in 1/576 inch; a px is 1/96 inch. */
const LEGACY_PER_PX = 6;
/** Comments (threads and replies together) kept per slide; the part itself is bounded by the package's part size limit. */
const MAX_COMMENTS = 10_000;

interface Author {
  name: string;
  initials?: string;
}

/** A timestamp as UTC ISO-8601: the files write local-looking times without a zone, which are read as UTC so the result does not depend on the machine. */
export function isoTime(value: string | undefined): string | undefined {
  const v = (value ?? '').trim();
  if (!/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}(?::\d{2}(?:\.\d+)?)?(?:Z|[+-]\d{2}:?\d{2})?$/.test(v)) return undefined;
  const [y, mo, d, h, mi] = (v.match(/\d+/g) as string[]).slice(0, 5).map(Number);
  const t = Date.parse(/(?:Z|[+-]\d{2}:?\d{2})$/.test(v) ? v : `${v}Z`);
  // Date.parse rolls 02-31 over to March and 24:00 to the next day: only a date that reads back as written is a date
  const day = new Date(Date.UTC(y, mo - 1, d));
  if (Number.isNaN(t) || day.getUTCMonth() !== mo - 1 || day.getUTCDate() !== d || h > 23 || mi > 59) return undefined;
  return new Date(t).toISOString();
}

const comment = (author: Author | undefined, body: string, created: string | undefined): PptxComment => ({
  author: author?.name ?? UNKNOWN_AUTHOR,
  ...(author?.initials ? { initials: author.initials } : {}),
  body,
  ...(isoTime(created) ? { createdAt: isoTime(created) as string } : {}),
});

/** The authors of the presentation's author parts, by id (an integer for legacy, a GUID for modern). */
export async function readAuthors(pkg: OoxmlPackage, presentationRels: Rel[]): Promise<Map<string, Author>> {
  const authors = new Map<string, Author>();
  const parts = presentationRels.filter((r) => (r.type === 'commentAuthors' || r.type === 'authors') && !r.external && pkg.has(r.target)).map((r) => r.target);
  for (const part of parts) {
    const root = await pkg.xml(part);
    for (const a of [...descendants(root, 'cmAuthor'), ...descendants(root, 'author')]) {
      const name = (a.attrs.name ?? '').trim();
      const initials = (a.attrs.initials ?? '').trim();
      if (a.attrs.id !== undefined && !authors.has(a.attrs.id)) authors.set(a.attrs.id, { name: name !== '' ? name : UNKNOWN_AUTHOR, ...(initials !== '' ? { initials } : {}) });
    }
  }
  return authors;
}

/** The text of a modern comment body (`p188:txBody`): the runs of each paragraph, the paragraphs joined by a line break. */
function modernBody(txBody: XmlElement | undefined): string {
  return kids(txBody, 'p')
    .map((p) => descendants(p, 't').map((t) => t.text).join(''))
    .join('\n')
    .trim();
}

/** Every `pptxId` in a slide's elements, groups included. */
export function elementIds(elements: PptxElement[], into = new Set<string>()): Set<string> {
  for (const e of elements) {
    into.add(e.pptxId);
    if (e.kind === 'group') elementIds(e.children, into);
  }
  return into;
}

/** The anchor of a modern comment: the shape its `*MkLst` marks (a `spMk`, `picMk`, `graphicFrameMk`... with an `id`) when the slide holds it, else the slide. */
function modernAnchor(cm: XmlElement, ids: ReadonlySet<string>): PptxCommentThread['anchor'] {
  for (const list of cm.children.filter((c) => /MkLst$/.test(c.name))) {
    let found: string | undefined;
    walk(list, (e) => {
      if (found === undefined && /Mk$/.test(e.name) && e.name !== 'docMk' && e.name !== 'sldMk' && e.attrs.id !== undefined && ids.has(e.attrs.id)) found = e.attrs.id;
    });
    if (found !== undefined) return { kind: 'element', pptxId: found };
  }
  return { kind: 'slide' };
}

function modernThread(cm: XmlElement, id: string, authors: ReadonlyMap<string, Author>, ids: ReadonlySet<string>): PptxCommentThread {
  const replies = descendants(kid(cm, 'replyLst'), 'reply').map((r) => comment(authors.get(r.attrs.authorId ?? ''), modernBody(kid(r, 'txBody')), r.attrs.created));
  const status = cm.attrs.status;
  return {
    id,
    format: 'modern',
    anchor: modernAnchor(cm, ids),
    ...(status === 'resolved' || status === 'closed' ? { resolved: true } : {}),
    comments: [comment(authors.get(cm.attrs.authorId ?? ''), modernBody(kid(cm, 'txBody')), cm.attrs.created), ...replies],
  };
}

/**
 * The threads of one legacy comments part: a `cm` with a `parentCm` is a reply to the comment with that author and index,
 * wherever that comment is in the part (a reply to a reply joins the thread of the comment at the top). A reply whose
 * parent is missing, or whose chain of parents loops, is kept as a thread of its own.
 */
function legacyThreads(comments: XmlElement[], slideId: string, authors: ReadonlyMap<string, Author>, taken: Set<string>): PptxCommentThread[] {
  const keyOf = (cm: XmlElement): string => `${cm.attrs.authorId ?? ''}.${cm.attrs.idx ?? ''}`;
  const parentOf = (cm: XmlElement): string | undefined => {
    const parent = descendants(cm, 'parentCm')[0];
    return parent ? `${parent.attrs.authorId ?? ''}.${parent.attrs.idx ?? ''}` : undefined;
  };
  // the first comment of each author and index is the one a reply can name
  const byKey = new Map<string, XmlElement>();
  for (const cm of comments) if (!byKey.has(keyOf(cm))) byKey.set(keyOf(cm), cm);
  /** The comment at the top of the chain of parents above `cm` (its own thread when it has none), or `cm` when the chain loops. */
  const topOf = (cm: XmlElement): XmlElement => {
    const seen = new Set<XmlElement>([cm]);
    let top = cm;
    for (;;) {
      const key = parentOf(top);
      const parent = key === undefined ? undefined : byKey.get(key);
      if (!parent) return top;
      if (seen.has(parent)) return cm;
      seen.add(parent);
      top = parent;
    }
  };
  const threads = new Map<XmlElement, PptxCommentThread>();
  const entryOf = (cm: XmlElement): PptxComment => comment(authors.get(cm.attrs.authorId ?? ''), (kid(cm, 'text')?.text ?? '').trim(), cm.attrs.dt);
  const roots = comments.filter((cm) => topOf(cm) === cm);
  for (const cm of roots) {
    const pos = kid(cm, 'pos');
    const [x, y] = [num(pos?.attrs.x), num(pos?.attrs.y)];
    let id = `${slideId}-${keyOf(cm)}`;
    for (let n = 2; taken.has(id); n++) id = `${slideId}-${keyOf(cm)}-${n}`;
    taken.add(id);
    threads.set(cm, { id, format: 'legacy', anchor: { kind: 'slide' }, ...(x !== undefined && y !== undefined ? { position: { x: Math.round((x / LEGACY_PER_PX) * 100) / 100, y: Math.round((y / LEGACY_PER_PX) * 100) / 100 } } : {}), comments: [entryOf(cm)] });
  }
  for (const cm of comments) if (!threads.has(cm)) threads.get(topOf(cm))?.comments.push(entryOf(cm));
  return roots.map((cm) => threads.get(cm) as PptxCommentThread);
}

/**
 * The comment threads of a slide: every comments part its relationships name (legacy or modern, told apart by the body
 * element of its `cm`s), in file order. `authors` comes from `readAuthors`; `ids` are the slide's element ids.
 */
export async function readThreads(pkg: OoxmlPackage, slidePart: string, slideId: string, authors: ReadonlyMap<string, Author>, ids: ReadonlySet<string>, warn: (m: string) => void): Promise<PptxCommentThread[]> {
  const threads: PptxCommentThread[] = [];
  const taken = new Set<string>();
  for (const rel of (await pkg.rels(slidePart)).filter((r) => r.type === 'comments' && !r.external)) {
    const list = await pkg.xml(rel.target);
    if (!list || list.name !== 'cmLst') continue;
    const modern = kids(list, 'cm').filter((cm) => kid(cm, 'txBody'));
    const legacy = kids(list, 'cm').filter((cm) => !kid(cm, 'txBody'));
    for (const cm of modern) {
      const base = `${slideId}-${cm.attrs.id ?? 'm'}`;
      let id = base;
      for (let n = 2; taken.has(id); n++) id = `${base}-${n}`;
      taken.add(id);
      threads.push(modernThread(cm, id, authors, ids));
    }
    threads.push(...legacyThreads(legacy, slideId, authors, taken));
  }
  const total = threads.reduce((n, t) => n + t.comments.length, 0);
  if (total > MAX_COMMENTS) {
    warn(`A slide has more than ${MAX_COMMENTS} comments; the rest are dropped.`);
    let left = MAX_COMMENTS;
    for (const thread of threads) {
      thread.comments.length = Math.min(thread.comments.length, left);
      left -= thread.comments.length;
    }
    threads.splice(0, threads.length, ...threads.filter((t) => t.comments.length > 0));
  }
  return threads;
}
