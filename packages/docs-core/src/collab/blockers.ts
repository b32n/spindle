// What the collaboration (Y.Doc) layout cannot hold.
//
// A document that contains such a construct must not be attached to a room: the shared copy would silently lose it.
// `getCollabBlockers` lists those constructs so a host can tell the user why collaboration is unavailable for this
// document, and `DocumentImpl.attachCollab` refuses with the same list. The rule for new document constructs is that
// they are representable in the Y.Doc (a ProseMirror node or mark, or an additive Y map) before they ship; a construct
// that is temporarily model-only is added here as a check and removed again when its Y representation lands.

import type { Block, Document, DocumentData, InlineContent } from '../types';

/** One kind of content that collaboration cannot carry, with how many places in the document have it. */
export interface CollabBlocker {
  /** Stable id, e.g. `collab.duplicateSectionIds`. */
  id: string;
  /** A sentence a host can show to the user. */
  message: string;
  /** How many places in the document are affected. */
  count: number;
}

/** Thrown by `attachCollab` when the document has content collaboration cannot hold. */
export class CollabBlockedError extends Error {
  readonly blockers: CollabBlocker[];
  constructor(blockers: CollabBlocker[]) {
    super(
      'Collaboration is not available for this document: ' +
        blockers.map((b) => `${b.message} (${b.count})`).join('; '),
    );
    this.name = 'CollabBlockedError';
    this.blockers = blockers;
  }
}

const KNOWN_BLOCK_TYPES: ReadonlySet<string> = new Set([
  'paragraph',
  'heading',
  'list-item',
  'table',
  'image',
  'horizontal-rule',
  'page-break',
  'content-control',
  'toc',
]);

const KNOWN_INLINE_TYPES: ReadonlySet<string> = new Set(['text', 'image', 'link', 'dynamicField', 'field', 'bookmark', 'contentControl', 'noteRef']);

interface Tally {
  unknownBlocks: number;
  unknownInlines: number;
}

function countInlines(content: readonly InlineContent[] | undefined, tally: Tally): void {
  for (const item of content ?? []) {
    if (!KNOWN_INLINE_TYPES.has((item as { type: string }).type)) tally.unknownInlines++;
    else if (item.type === 'contentControl' || item.type === 'noteRef') countInlines(item.content, tally);
  }
}

function countBlocks(blocks: readonly Block[] | undefined, tally: Tally): void {
  for (const block of blocks ?? []) {
    if (!KNOWN_BLOCK_TYPES.has(block.type)) {
      tally.unknownBlocks++;
      continue;
    }
    if (block.type === 'content-control' || block.type === 'toc') {
      countBlocks(block.blocks, tally);
    } else if (block.type === 'table') {
      for (const row of block.rows ?? []) {
        for (const cell of row.cells ?? []) {
          countInlines(cell.content, tally);
          countBlocks(cell.blocks, tally);
        }
      }
    } else if ('content' in block) {
      countInlines(block.content, tally);
    }
  }
}

/**
 * The content of a document that collaboration cannot hold (empty when the document can be attached).
 *
 * Multi-section documents are not a blocker: sections are carried by `section_break` nodes. What remains are structural
 * problems the shared layout cannot represent: no sections at all, two sections with the same id (the setup of a
 * section is looked up by id), and block or inline types this version does not know.
 */
export function getCollabBlockers(document: Pick<Document | DocumentData, 'sections'>): CollabBlocker[] {
  const blockers: CollabBlocker[] = [];
  const sections = document.sections ?? [];

  if (sections.length === 0) {
    blockers.push({ id: 'collab.noSections', message: 'The document has no sections', count: 1 });
    return blockers;
  }

  const seen = new Set<string>();
  let duplicates = 0;
  for (const section of sections) {
    if (seen.has(section.id)) duplicates++;
    seen.add(section.id);
  }
  if (duplicates > 0) {
    blockers.push({
      id: 'collab.duplicateSectionIds',
      message: 'Two or more sections share the same id',
      count: duplicates,
    });
  }

  const tally: Tally = { unknownBlocks: 0, unknownInlines: 0 };
  for (const section of sections) countBlocks(section.blocks, tally);
  if (tally.unknownBlocks > 0) {
    blockers.push({
      id: 'collab.unsupportedBlock',
      message: 'The document has block types that collaboration does not support',
      count: tally.unknownBlocks,
    });
  }
  if (tally.unknownInlines > 0) {
    blockers.push({
      id: 'collab.unsupportedInline',
      message: 'The document has inline content types that collaboration does not support',
      count: tally.unknownInlines,
    });
  }
  return blockers;
}
