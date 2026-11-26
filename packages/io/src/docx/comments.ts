import type { DocComment, DocCommentThread } from '@b32nio/spindle-shared/export/docx';

import { attrsOf, childrenOf, find, findAll, parseXml, tagOf, type XmlNode } from './xml-tree';

// ---------------------------------------------------------------------------
// Comments (word/comments.xml, commentsExtended.xml, people.xml)
//
// comments.xml holds the comment bodies; the body of the document marks where each one is anchored with
// commentRangeStart / commentRangeEnd (ids from comments.xml). Threading (replies) and the resolved state
// live in commentsExtended.xml, keyed by the w14:paraId of a comment's last paragraph. people.xml maps an
// author name to a stable user id. Every part is optional and may be damaged: whatever can be read is kept.
// ---------------------------------------------------------------------------

/** Anchor quotes are display text only; cap them so a comment on a whole chapter does not copy the chapter. */
const MAX_QUOTE_LENGTH = 200;

/** Concatenated text of one comment paragraph: runs, tabs and line breaks; deleted text and field codes are skipped. */
function paragraphText(node: XmlNode): string {
  let out = '';
  const visit = (n: XmlNode) => {
    for (const child of childrenOf(n)) {
      const tag = tagOf(child);
      if (tag === '#text') out += String((child as Record<string, unknown>)['#text'] ?? '');
      else if (tag === 'w:del' || tag === 'w:moveFrom' || tag === 'w:instrText' || tag === 'w:delText') continue;
      else if (tag === 'w:tab') out += '\t';
      else if (tag === 'w:br' || tag === 'w:cr') out += '\n';
      else visit(child);
    }
  };
  visit(node);
  return out;
}

interface ParsedComment {
  /** w:id */
  id: string;
  comment: DocComment;
  /** Position in comments.xml. */
  index: number;
  /** w14:paraId of the last paragraph (what commentsExtended refers to). */
  paraId?: string;
}

function parseCommentsPart(xml: string | undefined, authorIds: Map<string, string>, warnings: string[]): ParsedComment[] {
  if (!xml) return [];
  let root: XmlNode[];
  try {
    root = parseXml(xml);
  } catch {
    warnings.push('The comments part could not be read, so its comments were not imported.');
    return [];
  }
  const commentsEl = find(root, 'w:comments');
  if (!commentsEl) {
    warnings.push('The comments part could not be read, so its comments were not imported.');
    return [];
  }
  const out: ParsedComment[] = [];
  const seen = new Set<string>();
  for (const el of findAll(childrenOf(commentsEl), 'w:comment')) {
    const a = attrsOf(el);
    const id = a['@_w:id'];
    if (id === undefined || seen.has(id)) continue;
    seen.add(id);
    const paragraphs = findAll(childrenOf(el), 'w:p');
    const text = paragraphs.map(paragraphText).join('\n');
    const lastPara = paragraphs[paragraphs.length - 1];
    const author = a['@_w:author'] || 'Unknown';
    const comment: DocComment = {
      id,
      author,
      authorId: authorIds.get(author) ?? author,
      text,
      ...(a['@_w:initials'] && { initials: a['@_w:initials'] }),
      ...(a['@_w:date'] && { date: a['@_w:date'] }),
    };
    out.push({ id, comment, index: out.length, paraId: lastPara ? attrsOf(lastPara)['@_w14:paraId'] : undefined });
  }
  return out;
}

function parsePeople(xml: string | undefined): Map<string, string> {
  const ids = new Map<string, string>();
  if (!xml) return ids;
  try {
    const people = find(parseXml(xml), 'w15:people');
    for (const person of people ? findAll(childrenOf(people), 'w15:person') : []) {
      const author = attrsOf(person)['@_w15:author'];
      const info = find(childrenOf(person), 'w15:presenceInfo');
      const userId = info ? attrsOf(info)['@_w15:userId'] : undefined;
      if (author && userId) ids.set(author, userId);
    }
  } catch {
    // people.xml only improves author identity; ignore a damaged one.
  }
  return ids;
}

interface ExtendedEntry {
  parent?: string;
  done: boolean;
}

function parseExtended(xml: string | undefined): Map<string, ExtendedEntry> {
  const entries = new Map<string, ExtendedEntry>();
  if (!xml) return entries;
  try {
    const root = find(parseXml(xml), 'w15:commentsEx');
    for (const el of root ? findAll(childrenOf(root), 'w15:commentEx') : []) {
      const a = attrsOf(el);
      const paraId = a['@_w15:paraId'];
      if (!paraId) continue;
      entries.set(paraId, { parent: a['@_w15:paraIdParent'], done: a['@_w15:done'] === '1' || a['@_w15:done'] === 'true' });
    }
  } catch {
    // Threading is an enhancement: a damaged commentsExtended leaves every comment a thread of its own.
  }
  return entries;
}

export interface ParsedComments {
  /** Threads in comments.xml order of their root comment, quotes still empty. */
  threads: DocCommentThread[];
  /** w:id of every comment -> id of the thread it belongs to. */
  threadOf: Map<string, string>;
}

export function parseComments(
  parts: { comments?: string; extended?: string; people?: string },
  warnings: string[]
): ParsedComments {
  const parsed = parseCommentsPart(parts.comments, parsePeople(parts.people), warnings);
  const extended = parseExtended(parts.extended);
  const byParaId = new Map<string, ParsedComment>();
  for (const c of parsed) if (c.paraId) byParaId.set(c.paraId, c);

  const rootOf = (c: ParsedComment): ParsedComment => {
    let current = c;
    const chain: ParsedComment[] = [current];
    for (;;) {
      const parentPara = current.paraId ? extended.get(current.paraId)?.parent : undefined;
      const parent = parentPara ? byParaId.get(parentPara) : undefined;
      if (!parent) return current;
      // A parent loop has no natural root; take the first comment of the loop so every member agrees.
      if (chain.includes(parent)) return chain.slice(chain.indexOf(parent)).reduce((a, b) => (b.index < a.index ? b : a));
      chain.push(parent);
      current = parent;
    }
  };

  const threads = new Map<string, DocCommentThread>();
  const threadOf = new Map<string, string>();
  for (const c of parsed) {
    const root = rootOf(c);
    const threadId = `docx-thread-${root.id}`;
    let thread = threads.get(threadId);
    if (!thread) {
      const done = root.paraId ? extended.get(root.paraId)?.done : false;
      thread = { id: threadId, quote: '', comments: [], ...(done && { resolved: true }) };
      threads.set(threadId, thread);
    }
    thread.comments.push(c.comment);
    threadOf.set(c.id, threadId);
  }
  return { threads: [...threads.values()], threadOf };
}

/**
 * Follows the comment ranges while the document body is walked in order. A run that is inside one or more
 * ranges is tagged with a thread id (the innermost range wins, since the `comment` mark carries one id);
 * the text of every open range is also gathered as that thread's anchor quote.
 */
export interface CommentTracker {
  /** Consume a body/paragraph child that may hold range markers anywhere inside it. */
  mark(node: XmlNode): void;
  /** A new paragraph starts: text that follows is separated from the quote so far. */
  paragraphStart(): void;
  /** Text imported at the current position: records it in the open threads' quotes, returns the thread to tag it with. */
  anchorText(text: string): string | undefined;
  /** Threads with their anchor quotes filled in. */
  threads(): DocCommentThread[];
}

export function createCommentTracker(parsed: ParsedComments, documentXml: string): CommentTracker {
  // A range whose end is missing would swallow the rest of the document; treat it as having no anchor.
  const starts = new Set([...documentXml.matchAll(/<w:commentRangeStart\b[^>]*\bw:id="([^"]*)"/g)].map((m) => m[1]));
  const ends = new Set([...documentXml.matchAll(/<w:commentRangeEnd\b[^>]*\bw:id="([^"]*)"/g)].map((m) => m[1]));
  const balanced = new Set([...starts].filter((id) => ends.has(id) && parsed.threadOf.has(id)));

  const open: string[] = []; // comment ids, outermost first
  const quotes = new Map<string, string>();
  const quoteFull = new Set<string>();
  let paragraphBreak = false;

  const mark = (node: XmlNode) => {
    const tag = tagOf(node);
    if (tag === 'w:commentRangeStart' || tag === 'w:commentRangeEnd') {
      const id = attrsOf(node)['@_w:id'];
      if (id === undefined || !balanced.has(id)) return;
      if (tag === 'w:commentRangeStart') {
        if (!open.includes(id)) open.push(id);
      } else {
        const i = open.indexOf(id);
        if (i >= 0) open.splice(i, 1);
      }
      return;
    }
    for (const child of childrenOf(node)) mark(child);
  };

  const anchorText = (text: string): string | undefined => {
    if (open.length === 0) return undefined;
    const threadIds = [...new Set(open.map((id) => parsed.threadOf.get(id)!))];
    for (const t of threadIds) {
      if (quoteFull.has(t)) continue;
      // Text of separate paragraphs is joined with a space, like the selection quote the editor itself makes.
      const joiner = paragraphBreak && quotes.has(t) ? ' ' : '';
      const next = (quotes.get(t) ?? '') + joiner + text;
      if (next.length > MAX_QUOTE_LENGTH) {
        quotes.set(t, next.slice(0, MAX_QUOTE_LENGTH) + '…');
        quoteFull.add(t);
      } else {
        quotes.set(t, next);
      }
    }
    paragraphBreak = false;
    return parsed.threadOf.get(open[open.length - 1]);
  };

  return {
    mark,
    paragraphStart: () => {
      paragraphBreak = true;
    },
    anchorText,
    threads: () => parsed.threads.map((t) => ({ ...t, quote: quotes.get(t.id) ?? '' })),
  };
}
