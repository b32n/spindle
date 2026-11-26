// Editing around tracked changes: pasted tracked text gets change ids of its own, and typing inside a deleted run stays live.

import { Fragment, Slice, type Mark, type Node as PmNode, type Schema } from 'prosemirror-model';
import { ReplaceStep } from 'prosemirror-transform';
import { Plugin, type EditorState, type Transaction } from 'prosemirror-state';

let counter = 0;

/** A change id that no other change of a document is likely to have (also not one made by another client). */
export function newRevisionId(): string {
  return `n${Date.now().toString(36)}${(counter++).toString(36)}${Math.random().toString(36).slice(2, 6)}`;
}

const TRACKED = ['insertion', 'deletion', 'formatChange'] as const;
/** The change kind of each paragraph and row attribute (ids are shared with the text marks of the same kind). */
const BLOCK_KINDS: ReadonlyArray<[string, string]> = [
  ['insertion', 'revisionInsertion'],
  ['deletion', 'revisionDeletion'],
  ['formatChange', 'revisionFormat'],
];

/**
 * The slice with a fresh change id for every distinct change in it (per kind), so pasting a copy of tracked text makes a
 * change of its own: accepting or rejecting one place never touches the other. Author and date stay as in the source.
 * Records on paragraphs and table rows get the same treatment, and share the new id of text marks of the same change and kind.
 * The runs of one change inside the slice still share one id.
 * The two ends of a move that are both in the slice get one new move name.
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
  const moves = new Map<string, string>();
  const freshMove = (name: string): string => {
    let fresh = moves.get(name);
    if (!fresh) {
      fresh = makeId();
      moves.set(name, fresh);
    }
    return fresh;
  };
  const freshAttrs = (m: Mark): Record<string, unknown> => ({
    ...m.attrs,
    revisionId: freshId(m.type.name, String(m.attrs.revisionId)),
    ...(typeof m.attrs.moveId === 'string' && m.attrs.moveId && { moveId: freshMove(m.attrs.moveId) }),
  });
  const mapNode = (node: PmNode): PmNode => {
    if (node.isText) {
      if (!node.marks.some((m) => (TRACKED as readonly string[]).includes(m.type.name))) return node;
      return node.mark(
        node.marks.map((m) =>
          (TRACKED as readonly string[]).includes(m.type.name) ? m.type.create(freshAttrs(m)) : m
        )
      );
    }
    const content = node.content.size ? mapFragment(node.content) : node.content;
    const attrs = freshBlockAttrs(node);
    return attrs ? node.type.create(attrs, content, node.marks) : node.copy(content);
  };
  /** The attributes of a paragraph or row with the ids of its change records renewed, or null when it has none. */
  const freshBlockAttrs = (node: PmNode): Record<string, unknown> | null => {
    let out: Record<string, unknown> | null = null;
    for (const [kind, attr] of BLOCK_KINDS) {
      const record = node.attrs[attr];
      const id = record && typeof record === 'object' ? (record as { id?: unknown }).id : undefined;
      // `readBlockRevision` accepts a number id from another client's document, so a number is renewed too.
      if (typeof id !== 'string' && typeof id !== 'number') continue;
      out ??= { ...node.attrs };
      out[attr] = { ...(record as object), id: freshId(kind, String(id)) };
    }
    return out;
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

/** The fragment without text that is deleted by a tracked change (the final text). */
function withoutDeletedText(fragment: Fragment, deletion: string): Fragment {
  const out: PmNode[] = [];
  fragment.forEach((child) => {
    if (child.isText) {
      if (!child.marks.some((m) => m.type.name === deletion)) out.push(child);
    } else out.push(child.content.size ? child.copy(withoutDeletedText(child.content, deletion)) : child);
  });
  return Fragment.fromArray(out);
}

/** The plain text of a slice as it is copied to the clipboard: text deleted by a tracked change is left out. */
export function finalClipboardText(slice: Slice, blockSeparator = '\n\n'): string {
  const content = withoutDeletedText(slice.content, 'deletion');
  return content.textBetween(0, content.size, blockSeparator);
}

/**
 * Typing that the browser reports as a DOM change (IME composition, dead keys, Android keyboards) does not always pass
 * `handleTextInput`; the text it inserts inside or beside a deletion is parsed from the deleted span and carries the
 * deletion. This takes the deletion off text inserted by such an input, as typing does. Pastes, drops, undo and
 * collaboration keep the marks of what they insert.
 */
export function stripDeletionFromInput(trs: readonly Transaction[], newState: EditorState): Transaction | null {
  const type = newState.schema.marks.deletion;
  if (!type) return null;
  const ranges: [number, number][] = [];
  trs.forEach((tr, t) => {
    if (!tr.docChanged || tr.getMeta('paste') || tr.getMeta('history$') || tr.getMeta('addToHistory') === false) return;
    if (tr.getMeta('composition') === undefined && tr.getMeta('uiEvent') !== 'input') return;
    const later = trs.slice(t + 1);
    tr.steps.forEach((step, i) => {
      if (!(step instanceof ReplaceStep)) return;
      let hasDeleted = false;
      step.slice.content.descendants((n) => {
        if (n.isText && type.isInSet(n.marks)) hasDeleted = true;
      });
      if (!hasDeleted) return;
      step.getMap().forEach((_oldStart, _oldEnd, from, to) => {
        const rest = tr.mapping.slice(i + 1);
        let a = rest.map(from, 1);
        let b = rest.map(to, -1);
        for (const l of later) {
          a = l.mapping.map(a, 1);
          b = l.mapping.map(b, -1);
        }
        if (b > a) ranges.push([a, b]);
      });
    });
  });
  if (!ranges.length) return null;
  const tr = newState.tr;
  for (const [a, b] of ranges) tr.removeMark(a, Math.min(b, tr.doc.content.size), type);
  return tr.docChanged ? tr : null;
}

/** Plugin: pasted tracked text gets fresh change ids, and typing or pasting plain text inside a deletion is not part of it. */
export function createRevisionEditPlugin(schema: Schema): Plugin | null {
  if (!schema.marks.insertion && !schema.marks.deletion && !schema.marks.formatChange) return null;
  return new Plugin({
    appendTransaction: (trs, _oldState, newState) => stripDeletionFromInput(trs, newState),
    props: {
      transformPasted: (slice) => freshRevisionIds(slice),
      clipboardTextSerializer: (slice) => finalClipboardText(slice),
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
