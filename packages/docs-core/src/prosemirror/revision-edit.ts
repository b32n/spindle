// Editing around tracked changes: pasted tracked text gets change ids of its own, and typing inside a deleted run stays live.

import { Fragment, Slice, type Node as PmNode, type Schema } from 'prosemirror-model';
import { Plugin, type EditorState, type Transaction } from 'prosemirror-state';

let counter = 0;

/** A change id that no other change of a document is likely to have (also not one made by another client). */
export function newRevisionId(): string {
  return `n${Date.now().toString(36)}${(counter++).toString(36)}${Math.random().toString(36).slice(2, 6)}`;
}

const TRACKED = ['insertion', 'deletion'] as const;

/**
 * The slice with a fresh change id for every distinct change in it (per kind), so pasting a copy of tracked text makes a
 * change of its own: accepting or rejecting one place never touches the other. Author and date stay as in the source.
 * The runs of one change inside the slice still share one id.
 */
export function freshRevisionIds(slice: Slice, makeId: () => string = newRevisionId): Slice {
  const ids = new Map<string, string>();
  const freshId = (kind: string, id: string): string => {
    const key = `${kind}:${id}`;
    let fresh = ids.get(key);
    if (!fresh) {
      fresh = makeId();
      ids.set(key, fresh);
    }
    return fresh;
  };
  const mapNode = (node: PmNode): PmNode => {
    if (node.isText) {
      if (!node.marks.some((m) => (TRACKED as readonly string[]).includes(m.type.name))) return node;
      return node.mark(
        node.marks.map((m) =>
          (TRACKED as readonly string[]).includes(m.type.name) ? m.type.create({ ...m.attrs, revisionId: freshId(m.type.name, String(m.attrs.revisionId)) }) : m
        )
      );
    }
    return node.content.size ? node.copy(mapFragment(node.content)) : node;
  };
  const mapFragment = (fragment: Fragment): Fragment => {
    const out: PmNode[] = [];
    fragment.forEach((child) => out.push(mapNode(child)));
    return Fragment.fromArray(out);
  };
  return new Slice(mapFragment(slice.content), slice.openStart, slice.openEnd);
}

/**
 * The transaction for typing `text` over `from`-`to`, or null when the plain insertion is right. Text typed strictly
 * inside a deleted run would inherit the deletion (a mark's `inclusive: false` only protects its edges) and be written to
 * the file as deleted text; it is inserted as live text instead. Text typed inside an insertion is more inserted text.
 */
export function typeOutsideDeletion(state: EditorState, from: number, to: number, text: string): Transaction | null {
  const type = state.schema.marks.deletion;
  if (!type) return null;
  const marks = state.storedMarks ?? state.doc.resolve(from).marks();
  if (!type.isInSet(marks)) return null;
  const tr = state.tr.insertText(text, from, to);
  tr.removeMark(from, from + text.length, type);
  tr.setStoredMarks(marks.filter((m) => m.type !== type));
  return tr.setMeta('uiEvent', 'input');
}

/**
 * The transaction for pasting plain text (the clipboard has no HTML) into a deleted run, or null when the default paste is
 * right. Plain text takes the marks around the caret, so inside a deletion it would become more deleted text, as typing did;
 * it is pasted as live text instead. Pasted HTML keeps its own marks (a copied deletion stays a deletion).
 */
export function pasteOutsideDeletion(state: EditorState, slice: Slice): Transaction | null {
  const type = state.schema.marks.deletion;
  if (!type) return null;
  const { from, to } = state.selection;
  const marks = state.storedMarks ?? state.doc.resolve(from).marks();
  if (!type.isInSet(marks)) return null;
  const tr = state.tr.replaceSelection(slice);
  tr.removeMark(from, tr.mapping.map(to, 1), type);
  tr.setStoredMarks(marks.filter((m) => m.type !== type));
  return tr.scrollIntoView();
}

/**
 * Shift-Enter inside tracked text: the line break is a "\n" in the text (like the DOCX reader makes it), because a marked
 * `hard_break` node would lose its change in collaboration (marks sync on text only).
 */
export function insertTrackedLineBreak(state: EditorState, dispatch?: (tr: Transaction) => void): boolean {
  if (!state.selection.empty) return false;
  const marks = state.storedMarks ?? state.selection.$from.marks();
  const tracked = marks.filter((m) => (TRACKED as readonly string[]).includes(m.type.name));
  if (!tracked.length) return false;
  if (dispatch) {
    const keep = marks.filter((m) => m.type.name !== 'deletion');
    dispatch(state.tr.replaceSelectionWith(state.schema.text('\n', keep), false).scrollIntoView());
  }
  return true;
}

/** Plugin: pasted tracked text gets fresh change ids, and typing or pasting plain text inside a deletion is not part of it. */
export function createRevisionEditPlugin(schema: Schema): Plugin | null {
  if (!schema.marks.insertion && !schema.marks.deletion) return null;
  return new Plugin({
    props: {
      transformPasted: (slice) => freshRevisionIds(slice),
      handlePaste(view, event, slice) {
        if (event.clipboardData?.getData('text/html')) return false;
        const tr = pasteOutsideDeletion(view.state, slice);
        if (!tr) return false;
        view.dispatch(tr.setMeta('paste', true).setMeta('uiEvent', 'paste'));
        return true;
      },
      handleTextInput(view, from, to, text) {
        const tr = typeOutsideDeletion(view.state, from, to, text);
        if (!tr) return false;
        view.dispatch(tr);
        return true;
      },
    },
  });
}
