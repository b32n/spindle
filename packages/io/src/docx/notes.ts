import { attrsOf, childrenOf, find, findAll, parseXml, type XmlNode } from './xml-tree';

/**
 * Footnotes and endnotes (`word/footnotes.xml`, `word/endnotes.xml`). Each note is a list of block elements; the
 * separator notes Word puts at ids -1 and 0 (or 0 and 1) are not notes. The parts are read leniently: whatever can be
 * parsed from a damaged part is kept.
 */

export type NoteKind = 'footnote' | 'endnote';

const SEPARATOR_TYPES = new Set(['separator', 'continuationSeparator', 'continuationNotice']);

/** The block elements of each real note of a notes part, by note id (the first one wins when an id repeats). */
export function parseNotes(xml: string | undefined, kind: NoteKind): Map<string, XmlNode[]> {
  const notes = new Map<string, XmlNode[]>();
  if (!xml) return notes;
  let root: XmlNode[];
  try {
    root = parseXml(xml);
  } catch {
    return notes;
  }
  const container = find(root, kind === 'footnote' ? 'w:footnotes' : 'w:endnotes');
  if (!container) return notes;
  for (const note of findAll(childrenOf(container), `w:${kind}`)) {
    const attrs = attrsOf(note);
    const type = attrs['@_w:type'];
    const id = attrs['@_w:id'];
    if ((type && SEPARATOR_TYPES.has(type)) || id === undefined || notes.has(id)) continue;
    notes.set(id, childrenOf(note));
  }
  return notes;
}

/** The ids of the notes of a kind that the body of the document refers to (`w:footnoteReference` / `w:endnoteReference`). */
export function referencedNoteIds(documentXml: string, kind: NoteKind): Set<string> {
  const ids = new Set<string>();
  const re = new RegExp(`<w:${kind}Reference\\b[^>]*?\\bw:id="([^"]*)"`, 'g');
  for (const m of documentXml.matchAll(re)) ids.add(m[1]);
  return ids;
}
