import { Fragment, type Node as PmNode, type Schema } from 'prosemirror-model';
import type { Command } from 'prosemirror-state';
import { TextSelection } from 'prosemirror-state';

// ---------------------------------------------------------------------------
// Footnote and endnote references (`note_ref`): finding, inserting and rewriting them. A note is numbered by where its
// reference stands (per kind, in document order), so inserting or deleting a reference renumbers the rest.
// ---------------------------------------------------------------------------

export type NoteKind = 'footnote' | 'endnote';

/** The inline node types a note's text can hold (the content of `note_ref`). */
const NOTE_INLINE_TYPES = new Set(['text', 'hard_break', 'inline_image', 'bookmark', 'field', 'dynamic_field']);

/** The nth reference of a kind (1-based) in the document, with its position, or null. */
export function findNoteRef(doc: PmNode, kind: NoteKind, ordinal: number): { pos: number; node: PmNode } | null {
  let seen = 0;
  let found: { pos: number; node: PmNode } | null = null;
  doc.descendants((node, pos) => {
    if (found) return false;
    if (node.type.name === 'note_ref') {
      if (node.attrs.kind === kind && ++seen === ordinal) found = { pos, node };
      return false;
    }
    return true;
  });
  return found;
}

/** Which reference of its kind (1-based, document order) the one at `pos` is, or null when no reference starts there. */
export function noteOrdinalAt(doc: PmNode, pos: number): { kind: NoteKind; ordinal: number } | null {
  if (pos < 0 || pos > doc.content.size) return null;
  const target = doc.nodeAt(pos);
  if (!target || target.type.name !== 'note_ref') return null;
  const kind: NoteKind = target.attrs.kind === 'endnote' ? 'endnote' : 'footnote';
  let ordinal = 0;
  let result: number | null = null;
  doc.descendants((node, at) => {
    if (result !== null) return false;
    if (node.type.name === 'note_ref') {
      if (node.attrs.kind === kind) ordinal++;
      if (at === pos) result = ordinal;
      return false;
    }
    return true;
  });
  return result === null ? null : { kind, ordinal: result };
}

/** An id for a new note of a kind: one more than the largest numeric id among the document's notes of that kind. */
export function nextNoteId(doc: PmNode, kind: NoteKind): string {
  let max = 0;
  doc.descendants((node) => {
    if (node.type.name === 'note_ref' && node.attrs.kind === kind) {
      const n = Number(node.attrs.noteId);
      if (Number.isInteger(n) && n > max) max = n;
    }
    return node.type.name !== 'note_ref';
  });
  return String(max + 1);
}

/**
 * Insert an empty note reference at the end of the selection and put the caret after it. The new note's text is
 * written afterwards (see `noteTextFragment`). Not applicable inside a note.
 */
export function insertNoteReference(kind: NoteKind): Command {
  return (state, dispatch) => {
    const { $to } = state.selection;
    const parent = $to.parent;
    const type = state.schema.nodes.note_ref;
    if (!type || !parent.inlineContent || !parent.type.contentMatch.matchType(type)) return false;
    if (dispatch) {
      const node = type.create({ kind, noteId: nextNoteId(state.doc, kind) });
      const tr = state.tr.insert($to.pos, node);
      tr.setSelection(TextSelection.near(tr.doc.resolve($to.pos + node.nodeSize)));
      dispatch(tr.scrollIntoView());
    }
    return true;
  };
}

/**
 * The content for a note from a document written by the note editor: its paragraphs joined by line breaks, with only the
 * inline content a note can hold (anything else is replaced by its own text).
 */
export function noteTextFragment(schema: Schema, doc: PmNode): Fragment {
  const out: PmNode[] = [];
  const addInline = (node: PmNode) => {
    if (NOTE_INLINE_TYPES.has(node.type.name)) {
      out.push(node);
    } else if (node.content.size > 0) {
      node.forEach(addInline);
    }
  };
  let first = true;
  doc.descendants((block) => {
    if (!block.isTextblock) return true;
    if (!first) out.push(schema.nodes.hard_break.create());
    first = false;
    block.forEach(addInline);
    return false;
  });
  return Fragment.from(out);
}
