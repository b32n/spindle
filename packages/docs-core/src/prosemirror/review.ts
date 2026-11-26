// Reviewing tracked changes: the changes of a document as the individual things that can be accepted or rejected, and the
// commands that accept or reject them.
//
// Accepting makes a change part of the document (an inserted run stays, a deleted run goes, new formatting stays, an
// inserted paragraph mark or row stays, a deleted one goes); rejecting undoes it (an inserted run goes, a deleted run
// stays, the old formatting comes back, an inserted paragraph mark or row goes, a deleted one stays). The two ends of a
// move are one change: accepting it deletes the text where it was and keeps it where it went. All the units a command
// acts on are changed in one transaction, so one undo takes it back.

import { type Mark, type MarkType, type Node as PmNode, type NodeType } from 'prosemirror-model';
import type { Command, EditorState, Transaction } from 'prosemirror-state';
import type { BlockRevision, ParagraphFormat, RunFormat } from '../types';
import { docsSchema } from './schema';
import { readBlockRevision, withBlockRevision, withBlockRevisionKind } from './block-revision';
import { paragraphFormatAttrs, runFormatMarks } from './sync';

/** What one unit of change is: text (an insertion, a deletion, a change of formatting), a paragraph mark, a paragraph's formatting, a table row. */
export type ChangeUnitKind = 'insertion' | 'deletion' | 'format' | 'markInsertion' | 'markDeletion' | 'paragraphFormat' | 'rowInsertion' | 'rowDeletion';

/** One thing a tracked change did, at one place in the document. */
export interface ChangeUnit {
  /** Which change it belongs to: the kind and the revision id (as in the file). */
  key: string;
  kind: ChangeUnitKind;
  id: string;
  author: string;
  date?: string;
  /** The name of the move when it is one end of a move. */
  moveId?: string;
  /** Text units: the range of the run. Block units: the node, from its start to its end. */
  from: number;
  to: number;
  /** The text covered (text units). */
  text: string;
  /** The mark (text units) the unit comes from. */
  mark?: Mark;
}

const KIND_LABEL: Record<ChangeUnitKind, string> = {
  insertion: 'ins',
  deletion: 'del',
  format: 'fmt',
  markInsertion: 'pins',
  markDeletion: 'pdel',
  paragraphFormat: 'pfmt',
  rowInsertion: 'rins',
  rowDeletion: 'rdel',
};

const keyOf = (kind: ChangeUnitKind, id: string) => `${KIND_LABEL[kind]}:${id}`;

const isTextblock = (node: PmNode) => node.type.name === 'paragraph' || node.type.name === 'heading';

/** Every unit of tracked change in the document, in document order (a block's before the text in it). */
export function changeUnits(doc: PmNode): ChangeUnit[] {
  const units: ChangeUnit[] = [];
  doc.descendants((node, pos) => {
    if (node.isText) {
      for (const mark of node.marks) {
        const name = mark.type.name;
        const kind = name === 'insertion' ? 'insertion' : name === 'deletion' ? 'deletion' : name === 'formatChange' ? 'format' : undefined;
        if (!kind) continue;
        const id = String(mark.attrs.revisionId);
        const date = mark.attrs.date;
        const moveId = mark.attrs.moveId;
        units.push({
          key: keyOf(kind, id),
          kind,
          id,
          author: String(mark.attrs.author),
          ...(typeof date === 'string' && date && { date }),
          ...(typeof moveId === 'string' && moveId && { moveId }),
          from: pos,
          to: pos + node.nodeSize,
          text: node.text ?? '',
          mark,
        });
      }
      return false;
    }
    const revision = isTextblock(node) || node.type.name === 'table_row' ? readBlockRevision(node.attrs, false) : undefined;
    if (revision) {
      const add = (kind: ChangeUnitKind, meta: { id?: string; author?: string; date?: string; moveId?: string } | undefined) => {
        if (!meta) return;
        const id = String(meta.id ?? '');
        units.push({
          key: keyOf(kind, id),
          kind,
          id,
          author: String(meta.author ?? ''),
          ...(meta.date && { date: meta.date }),
          ...(meta.moveId && { moveId: meta.moveId }),
          from: pos,
          to: pos + node.nodeSize,
          text: node.type.name === 'table_row' ? node.textContent : '',
        });
      };
      if (node.type.name === 'table_row') {
        add('rowInsertion', revision.insertion);
        add('rowDeletion', revision.deletion);
      } else {
        add('markInsertion', revision.insertion);
        add('markDeletion', revision.deletion);
        add('paragraphFormat', revision.formatChange);
      }
    }
    return true;
  });
  return units;
}

/** A change as a reviewer sees it: the units that are accepted or rejected together. */
export interface ReviewChange {
  /** Unique within the document: the keys of its units and where the first one is (changes can share a revision id). */
  id: string;
  kind: 'insertion' | 'deletion' | 'move' | 'format' | 'paragraphInserted' | 'paragraphDeleted' | 'paragraphFormat' | 'rowInserted' | 'rowDeleted';
  author: string;
  date?: string;
  /** The text it covers (the text moved, for a move). */
  text: string;
  /** Where it is (where a move goes to). */
  from: number;
  to: number;
  units: ChangeUnit[];
}

const BLOCK_KINDS: Record<string, ReviewChange['kind']> = {
  markInsertion: 'paragraphInserted',
  markDeletion: 'paragraphDeleted',
  paragraphFormat: 'paragraphFormat',
  rowInsertion: 'rowInserted',
  rowDeletion: 'rowDeleted',
};

/**
 * The changes of the document, in order: the runs of one insertion, deletion or formatting change (same revision id) that
 * touch are one change, the two ends of a move are one change, and each block change (a paragraph mark, a paragraph's
 * formatting, a row) is a change of its own.
 */
export function listChanges(doc: PmNode): ReviewChange[] {
  const changes: ReviewChange[] = [];
  const open = new Map<string, ReviewChange>();
  const moves = new Map<string, ReviewChange>();
  for (const unit of changeUnits(doc)) {
    if (BLOCK_KINDS[unit.kind]) {
      changes.push({ id: `${unit.key}@${unit.from}`, kind: BLOCK_KINDS[unit.kind], author: unit.author, ...(unit.date && { date: unit.date }), text: unit.text, from: unit.from, to: unit.to, units: [unit] });
      continue;
    }
    const joined = open.get(unit.key);
    if (joined && joined.to === unit.from) {
      joined.to = unit.to;
      joined.text += unit.text;
      joined.units.push(unit);
      continue;
    }
    const moveKey = unit.moveId && unit.kind !== 'format' ? `move:${unit.moveId}` : undefined;
    const move = moveKey ? moves.get(moveKey) : undefined;
    if (move) {
      // The other end of a move already listed: the change is where the text went to, and covers both ends.
      move.units.push(unit);
      move.id = `${move.units.map((u) => u.key).join('+')}@${move.units[0].from}`;
      if (unit.kind === 'insertion') {
        move.text = unit.text;
        move.from = unit.from;
        move.to = unit.to;
        move.author = unit.author;
        if (unit.date) move.date = unit.date;
      }
      open.set(unit.key, move);
      continue;
    }
    const change: ReviewChange = {
      id: `${unit.key}@${unit.from}`,
      kind: moveKey ? 'move' : (unit.kind as 'insertion' | 'deletion' | 'format'),
      author: unit.author,
      ...(unit.date && { date: unit.date }),
      text: unit.text,
      from: unit.from,
      to: unit.to,
      units: [unit],
    };
    open.set(unit.key, change);
    if (moveKey) moves.set(moveKey, change);
    changes.push(change);
  }
  return changes;
}

// ---------------------------------------------------------------------------
// Accepting and rejecting
// ---------------------------------------------------------------------------

const FORMAT_MARKS = ['bold', 'italic', 'underline', 'strikethrough', 'superscript', 'subscript', 'smallCaps', 'allCaps', 'hidden', 'textStyle', 'namedStyle'];

/**
 * Join the node at `pos` with the one after it, when both are paragraphs or headings of one parent: the mark between them
 * goes. The paragraph that results ends with the second one's mark, so it takes that mark's changes (inserted, deleted) and
 * keeps the first one's formatting and the change of it (the second one's formatting is not kept, so neither is its change).
 */
function joinWithNext(tr: Transaction, pos: number): void {
  const node = tr.doc.nodeAt(pos);
  if (!node) return;
  const end = pos + node.nodeSize;
  const next = tr.doc.resolve(end).nodeAfter;
  if (!next || !isTextblock(node) || !isTextblock(next)) return;
  const first = readBlockRevision(node.attrs, false) ?? {};
  const second = readBlockRevision(next.attrs, false) ?? {};
  const revision: BlockRevision = {
    ...(second.insertion && { insertion: second.insertion }),
    ...(second.deletion && { deletion: second.deletion }),
    ...(first.formatChange && { formatChange: first.formatChange }),
  };
  tr.join(end);
  const joined = tr.doc.nodeAt(pos);
  if (joined) tr.setNodeMarkup(pos, undefined, withBlockRevision(joined.attrs, revision));
}

/** The type and attrs a paragraph gets back when its formatting change is rejected. */
function previousParagraph(change: NonNullable<BlockRevision['formatChange']>): { type: NodeType; attrs: Record<string, unknown> } {
  const { headingLevel, ...previous } = change.previous;
  const level = headingLevel && headingLevel >= 1 && headingLevel <= 6 ? headingLevel : undefined;
  const format: ParagraphFormat = previous;
  const type = level ? docsSchema.nodes.heading : docsSchema.nodes.paragraph;
  // The paragraph keeps what the formatting change did not touch (nothing else is stored in a paragraph's attrs).
  return { type, attrs: { ...type.create().attrs, ...paragraphFormatAttrs(level ? 'heading' : 'paragraph', format, level) } };
}

/** Accept or reject one unit in `tr`. Units must be settled from the end of the document to its start (see `settle`). */
function settleUnit(tr: Transaction, unit: ChangeUnit, action: 'accept' | 'reject'): void {
  const accept = action === 'accept';
  switch (unit.kind) {
    case 'insertion':
    case 'deletion': {
      // Accepting an insertion or rejecting a deletion keeps the text and drops the mark; the other two remove the text.
      const keep = (unit.kind === 'insertion') === accept;
      if (keep) tr.removeMark(unit.from, unit.to, unit.mark);
      else tr.delete(unit.from, unit.to);
      return;
    }
    case 'format': {
      if (!accept) {
        const previous = (unit.mark?.attrs.previous ?? {}) as RunFormat;
        for (const name of FORMAT_MARKS) {
          const type: MarkType | undefined = docsSchema.marks[name];
          if (type) tr.removeMark(unit.from, unit.to, type);
        }
        for (const json of runFormatMarks(previous)) tr.addMark(unit.from, unit.to, docsSchema.mark(json.type, json.attrs));
      }
      tr.removeMark(unit.from, unit.to, unit.mark);
      return;
    }
    case 'markInsertion':
    case 'markDeletion': {
      const node = tr.doc.nodeAt(unit.from);
      if (!node) return;
      const kind = unit.kind === 'markInsertion' ? 'insertion' : 'deletion';
      // The mark is accepted when it stays: an inserted mark stays when accepted, a deleted one when rejected.
      const stays = (kind === 'insertion') === accept;
      tr.setNodeMarkup(unit.from, undefined, withBlockRevisionKind(node.attrs, kind, null));
      if (!stays) joinWithNext(tr, unit.from);
      return;
    }
    case 'paragraphFormat': {
      const node = tr.doc.nodeAt(unit.from);
      if (!node) return;
      const revision = readBlockRevision(node.attrs, false) ?? {};
      if (accept || !revision.formatChange) {
        tr.setNodeMarkup(unit.from, undefined, withBlockRevisionKind(node.attrs, 'formatChange', null));
        return;
      }
      const previous = previousParagraph(revision.formatChange);
      // A list item holds a paragraph first: a paragraph that was a heading before stays a paragraph there, with its other formatting back.
      const $pos = tr.doc.resolve(unit.from);
      const fits = $pos.parent.canReplaceWith($pos.index(), $pos.index() + 1, previous.type);
      const type = fits ? previous.type : docsSchema.nodes.paragraph;
      const attrs = fits ? previous.attrs : previousParagraph({ ...revision.formatChange, previous: { ...revision.formatChange.previous, headingLevel: undefined } }).attrs;
      // The node's own record of an inserted or deleted mark outlives the restoring, and so does its indent level (not part of the formatting).
      const kept = { ...attrs, ...(node.attrs.indent !== undefined && 'indent' in type.spec.attrs! && { indent: node.attrs.indent }) };
      tr.setNodeMarkup(unit.from, type, withBlockRevisionKind({ ...kept, revisionInsertion: node.attrs.revisionInsertion, revisionDeletion: node.attrs.revisionDeletion }, 'formatChange', null));
      return;
    }
    case 'rowInsertion':
    case 'rowDeletion': {
      const node = tr.doc.nodeAt(unit.from);
      if (!node) return;
      const kind = unit.kind === 'rowInsertion' ? 'insertion' : 'deletion';
      const stays = (kind === 'insertion') === accept;
      if (stays) {
        tr.setNodeMarkup(unit.from, undefined, withBlockRevisionKind(node.attrs, kind, null));
        return;
      }
      const $row = tr.doc.resolve(unit.from);
      if ($row.parent.childCount === 1) tr.delete($row.before(), $row.after());
      else tr.delete(unit.from, unit.from + node.nodeSize);
    }
  }
}

/** Where a unit is settled: its text's start, or for a paragraph mark that goes, the end of its paragraph (the join is there). */
const anchorOf = (unit: ChangeUnit) => (unit.kind === 'markInsertion' || unit.kind === 'markDeletion' ? unit.to : unit.from);
/**
 * The order units are settled in when they are at the same place: what keeps text first and what removes it after, so a
 * run that two units cover (inserted by one author, deleted by another) is not read after it is gone; then paragraph
 * formatting, rows, and last the paragraph marks, which may join paragraphs.
 */
function rank(unit: ChangeUnit, action: 'accept' | 'reject'): number {
  switch (unit.kind) {
    case 'format':
      return 0;
    case 'insertion':
    case 'deletion':
      return (unit.kind === 'insertion') === (action === 'accept') ? 0 : 1;
    case 'paragraphFormat':
      return 2;
    case 'rowInsertion':
    case 'rowDeletion':
      return 3;
    default:
      return 4;
  }
}

/**
 * A transaction that accepts or rejects `units`, or null when there are none. Units are settled from the end of the
 * document to its start, so what one does never moves what another still has to do.
 */
export function settle(state: EditorState, units: readonly ChangeUnit[], action: 'accept' | 'reject'): Transaction | null {
  if (units.length === 0) return null;
  const tr = state.tr;
  const ordered = [...units].sort((a, b) => anchorOf(b) - anchorOf(a) || rank(a, action) - rank(b, action));
  for (const unit of ordered) settleUnit(tr, unit, action);
  return tr.docChanged ? tr.scrollIntoView() : tr;
}

/** All units that belong to the same change as `unit`: its other runs and the other end of a move. */
function wholeChange(state: EditorState, changeId: string): ChangeUnit[] {
  const change = listChanges(state.doc).find((c) => c.id === changeId);
  return change ? change.units : [];
}

type Dispatch = ((tr: Transaction) => void) | undefined;

function run(state: EditorState, dispatch: Dispatch, units: readonly ChangeUnit[], action: 'accept' | 'reject'): boolean {
  const tr = settle(state, units, action);
  if (!tr) return false;
  dispatch?.(tr);
  return true;
}

/** Accept or reject one change (a `ReviewChange.id` of `listChanges`). */
export const resolveChange =
  (changeId: string, action: 'accept' | 'reject'): Command =>
  (state, dispatch) =>
    run(state, dispatch, wholeChange(state, changeId), action);

/** Accept or reject every change, or every change of one author (a move is one change, by the author of the end it went to). */
export const resolveAll =
  (action: 'accept' | 'reject', author?: string): Command =>
  (state, dispatch) =>
    run(
      state,
      dispatch,
      author === undefined ? changeUnits(state.doc) : listChanges(state.doc).filter((c) => c.author === author).flatMap((c) => c.units),
      action
    );

/** Accept or reject the changes the selection touches (a caret touches the change it is in or at the edge of). */
export const resolveSelection =
  (action: 'accept' | 'reject'): Command =>
  (state, dispatch) => {
    const { from, to } = state.selection;
    const touched = new Set(
      listChanges(state.doc)
        .filter((c) => c.units.some((u) => u.from <= to && u.to >= from && (u.to > from || u.from < to || from === to)))
        .map((c) => c.id)
    );
    const units = listChanges(state.doc)
      .filter((c) => touched.has(c.id))
      .flatMap((c) => c.units);
    return run(state, dispatch, units, action);
  };

