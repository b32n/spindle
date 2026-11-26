/**
 * NoteEditor: edits the text of one footnote or endnote in place, over its entry in the notes area.
 *
 * It is a small ProseMirror editor of its own (like the header and footer editor) on a document of one paragraph, so the
 * note keeps its formatting, links and line breaks. The main document is not touched while editing; closing the editor
 * (blur, Escape) hands the text back, and the caller writes it into the note's reference in one step.
 */

import { memo, useEffect, useRef } from 'react';
import { EditorState, Selection, type Command } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import type { Node as PmNode } from 'prosemirror-model';
import { baseKeymap, toggleMark } from 'prosemirror-commands';
import { keymap } from 'prosemirror-keymap';
import { history, undo, redo } from 'prosemirror-history';
import { docsSchema } from '@b32nio/spindle-docs-core';

export interface NoteEditorProps {
  /** The note reference being edited: its content is the text. */
  note: PmNode;
  /** Where the note is drawn, in the coordinates of the element the editor is placed in. */
  position: { top: number; left: number; width: number; minHeight: number };
  /** Zoom factor of the page, so the text has the size of the note's text on the page. */
  scale: number;
  /** Called once when editing ends, with the document of the note editor. */
  onClose: (doc: PmNode) => void;
}

/** Enter is a line break inside a note (a note is one paragraph), as in Word. */
const insertBreak: Command = (state, dispatch) => {
  if (dispatch) dispatch(state.tr.replaceSelectionWith(docsSchema.nodes.hard_break.create()).scrollIntoView());
  return true;
};

export const NoteEditor = memo(function NoteEditor({ note, position, scale, onClose }: NoteEditorProps) {
  const hostRef = useRef<HTMLDivElement>(null);
  const closeRef = useRef(onClose);
  closeRef.current = onClose;

  useEffect(() => {
    const host = hostRef.current;
    if (!host) return;
    const { paragraph, doc } = docsSchema.nodes;
    let closed = false;
    const view = new EditorView(host, {
      state: EditorState.create({
        doc: doc.create(null, [paragraph.create(null, note.content)]),
        plugins: [
          history(),
          keymap({
            'Mod-z': undo,
            'Mod-y': redo,
            'Shift-Mod-z': redo,
            'Mod-b': toggleMark(docsSchema.marks.bold),
            'Mod-i': toggleMark(docsSchema.marks.italic),
            'Mod-u': toggleMark(docsSchema.marks.underline),
            Enter: insertBreak,
            'Shift-Enter': insertBreak,
            Escape: () => {
              finish();
              return true;
            },
          }),
          keymap(baseKeymap),
        ],
      }),
    });
    function finish() {
      if (closed) return;
      closed = true;
      closeRef.current(view.state.doc);
    }
    view.dom.addEventListener('blur', finish);
    view.focus();
    // The caret goes to the end of the note's text.
    view.dispatch(view.state.tr.setSelection(Selection.atEnd(view.state.doc)));
    return () => {
      closed = true;
      view.dom.removeEventListener('blur', finish);
      view.destroy();
    };
    // The note reference (and so its text) is fixed for the life of one editing session.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  return (
    <div
      ref={hostRef}
      className="note-editor"
      data-testid="note-editor"
      style={{
        position: 'absolute',
        top: position.top,
        left: position.left,
        width: position.width,
        minHeight: position.minHeight,
        boxSizing: 'border-box',
        background: '#fff',
        border: '1px solid #1a73e8',
        borderRadius: 2,
        boxShadow: '0 2px 8px rgba(26, 115, 232, 0.25)',
        fontFamily: 'Arial, sans-serif',
        fontSize: 10 * scale,
        lineHeight: 1.5,
        zIndex: 200,
        padding: '0 2px',
      }}
    />
  );
});
