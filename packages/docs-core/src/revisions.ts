// Tracked changes: how they are told apart on screen and read out of a document.

import type { Node as PmNode } from 'prosemirror-model';
import { readBlockRevision } from './prosemirror/block-revision';

/** Colours the authors of tracked changes are drawn in, dark enough to read on white. */
export const REVISION_COLORS: readonly string[] = ['#c5221f', '#1a73e8', '#188038', '#a142f4', '#e8710a', '#129eaf', '#d01884', '#7b5e00'];

/**
 * The colour of an author's changes. It follows from the name alone (not from the order authors appear in), so every
 * client and every session shows the same author in the same colour, and nothing has to be stored for it.
 */
export function authorColor(author: string): string {
  // FNV-1a over the UTF-16 code units.
  let hash = 0x811c9dc5;
  for (let i = 0; i < author.length; i++) {
    hash ^= author.charCodeAt(i);
    hash = Math.imul(hash, 0x01000193) >>> 0;
  }
  return REVISION_COLORS[hash % REVISION_COLORS.length];
}

/**
 * The colours of a document's authors: each author is drawn in the colour `authorColor` gives them, unless another
 * author already has it, then in the next free colour of the palette. It follows from the set of authors alone (not from
 * the order they appear in), so every client shows the same colours; with more authors than colours they are reused.
 * `preferred` gives some authors a colour of their own (`collabAuthorColors`: the colour of their cursor in a shared document).
 */
export function assignAuthorColors(authors: Iterable<string>, preferred?: ReadonlyMap<string, string>): Map<string, string> {
  const names = [...new Set(authors)].sort();
  const colors = new Map<string, string>();
  const taken = new Set<string>();
  // Authors with a colour of their own (the colour of their cursor while they are in the room) keep it, and nobody else takes it.
  for (const name of names) {
    const color = preferred?.get(name);
    if (color && /^#[0-9a-f]{6}$/i.test(color) && !taken.has(color.toLowerCase())) {
      colors.set(name, color.toLowerCase());
      taken.add(color.toLowerCase());
    }
  }
  for (const name of names) {
    if (colors.has(name)) continue;
    const start = REVISION_COLORS.indexOf(authorColor(name));
    let color = REVISION_COLORS[start];
    for (let i = 0; i < REVISION_COLORS.length; i++) {
      const candidate = REVISION_COLORS[(start + i) % REVISION_COLORS.length];
      if (!taken.has(candidate)) {
        color = candidate;
        break;
      }
    }
    taken.add(color);
    colors.set(name, color);
  }
  return colors;
}

/**
 * The authors with tracked changes in a document, in order of first appearance: the marks of its text (insertions, deletions,
 * formatting changes) and the changes recorded on paragraphs, headings and table rows (a paragraph mark, a row, formatting).
 */
export function documentRevisionAuthors(doc: PmNode): string[] {
  const authors = new Set<string>();
  doc.descendants((node) => {
    if (node.isText) {
      for (const mark of node.marks) if (mark.type.name === 'insertion' || mark.type.name === 'deletion' || mark.type.name === 'formatChange') authors.add(String(mark.attrs.author));
      return true;
    }
    const revision = readBlockRevision(node.attrs);
    if (revision) {
      for (const kind of ['insertion', 'deletion', 'formatChange'] as const) {
        const author = revision[kind]?.author;
        if (typeof author === 'string') authors.add(author);
      }
    }
    return true;
  });
  return [...authors];
}

/** A tracked change in a ProseMirror document: the text it covers and where. */
export interface RevisionSummary {
  kind: 'insertion' | 'deletion';
  /** Groups the runs of one change. */
  id: string;
  author: string;
  date?: string;
  /** The text covered, in document order. */
  text: string;
  /** Position of the first and the end of the last run covered. */
  from: number;
  to: number;
}

/** The authors with tracked changes in the document, in order of first appearance, and how many changes each has. */
export interface RevisionAuthor {
  name: string;
  color: string;
  count: number;
}

/**
 * The tracked changes of the document in order, runs of one change (same kind, id and author) joined; the same id from two authors is two changes, as the DOCX writer treats it. With `withText` false the
 * `text` of each is left empty, which spares building the strings when only the authors and counts are wanted.
 * Runs are joined only when they sit next to each other: a change that crosses a paragraph boundary (pasted tracked text
 * of several paragraphs keeps one id) is one entry per paragraph, so the count can be higher than Word's.
 */
export function collectRevisions(doc: PmNode, withText = true): RevisionSummary[] {
  const found: RevisionSummary[] = [];
  const open = new Map<string, RevisionSummary>();
  doc.descendants((node, pos) => {
    if (!node.isText) return true;
    for (const mark of node.marks) {
      const kind = mark.type.name;
      if (kind !== 'insertion' && kind !== 'deletion') continue;
      const id = String(mark.attrs.revisionId);
      const key = `${kind}:${id}:${mark.attrs.author}`;
      const last = open.get(key);
      const text = withText ? (node.text ?? '') : '';
      if (last && last.to === pos) {
        last.text += text;
        last.to = pos + node.nodeSize;
      } else {
        const date = mark.attrs.date;
        const entry: RevisionSummary = { kind, id, author: String(mark.attrs.author), ...(typeof date === 'string' && date && { date }), text, from: pos, to: pos + node.nodeSize };
        open.set(key, entry);
        found.push(entry);
      }
    }
    return true;
  });
  return found;
}

/**
 * The authors of the given changes, in order of first appearance, each with the colour `assignAuthorColors` gives them.
 * `palette` is every author of the document, when that is more than the authors of these changes (the colours are shared
 * out among all of them, so they match the page). `preferred` are colours some authors are shown in already (cursor colours).
 */
export function revisionAuthors(revisions: ReadonlyArray<{ author: string }>, palette?: Iterable<string>, preferred?: ReadonlyMap<string, string>): RevisionAuthor[] {
  const colors = assignAuthorColors(palette ?? revisions.map((r) => r.author), preferred);
  const byName = new Map<string, RevisionAuthor>();
  for (const r of revisions) {
    const entry = byName.get(r.author) ?? { name: r.author, color: colors.get(r.author) as string, count: 0 };
    entry.count++;
    byName.set(r.author, entry);
  }
  return [...byName.values()];
}

/** True when two lists of authors show the same names and counts, so a list in state can be kept (and a render skipped). */
export function sameRevisionAuthors(a: readonly RevisionAuthor[], b: readonly RevisionAuthor[]): boolean {
  return a.length === b.length && a.every((author, i) => author.name === b[i].name && author.count === b[i].count && author.color === b[i].color);
}
