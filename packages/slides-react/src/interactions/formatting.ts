// Applies a TextFormatSpec either to the live ProseMirror editor (range
// formatting via transactions) or, when nothing is being edited, to the
// selected element's stored JSON (idle path, whole-body).

import { toggleMark } from 'prosemirror-commands';
import type { EditorView } from 'prosemirror-view';
import type { DeckImpl, TextFormatSpec } from '@b32nio/spindle-slides-core';

function applyToView(view: EditorView, spec: TextFormatSpec): void {
  const { state } = view;
  const schema = state.schema;

  if (spec.toggleMark) {
    const mt = schema.marks[spec.toggleMark];
    if (mt) toggleMark(mt)(state, view.dispatch);
  } else if (spec.setMark) {
    const mt = schema.marks[spec.setMark.type];
    if (mt) {
      const { from, to, empty } = state.selection;
      if (empty) {
        view.dispatch(state.tr.addStoredMark(mt.create(spec.setMark.attrs)));
      } else {
        view.dispatch(state.tr.removeMark(from, to, mt).addMark(from, to, mt.create(spec.setMark.attrs)));
      }
    }
  } else if (spec.removeMark) {
    const mt = schema.marks[spec.removeMark];
    if (mt) {
      const { from, to } = state.selection;
      view.dispatch(state.tr.removeMark(from, to, mt));
    }
  } else if (spec.paragraph) {
    const { from, to } = state.selection;
    let tr = state.tr;
    state.doc.nodesBetween(from, to, (node, pos) => {
      if (node.type.name === 'paragraph') {
        tr = tr.setNodeMarkup(pos, undefined, { ...node.attrs, ...spec.paragraph });
      }
    });
    view.dispatch(tr);
  }
  view.focus();
}

export interface FormatTarget {
  editingView: EditorView | null;
  deck: DeckImpl;
  /** Selected element (idle path) when not editing. */
  elementId: string | null;
  /** A table cell-range to format (idle path) — takes priority over elementId
   *  when no cell is being edited, so a whole header row formats at once. */
  tableCells?: { tableId: string; cells: Array<[number, number]> } | null;
}

export function applyFormat(target: FormatTarget, spec: TextFormatSpec): void {
  if (target.editingView) applyToView(target.editingView, spec);
  else if (target.tableCells && target.tableCells.cells.length) target.deck.applyTableCellsFormat(target.tableCells.tableId, target.tableCells.cells, spec);
  else if (target.elementId) target.deck.applyTextFormat(target.elementId, spec);
}
