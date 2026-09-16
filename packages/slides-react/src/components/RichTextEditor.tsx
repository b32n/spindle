// RichTextEditor — the single live ProseMirror mount. It swaps in over the
// actively-edited element with identical box metrics (no visual jump vs
// StaticRichText).
//
// Non-collab: builds state from the element's JSON, keeps a local
// prosemirror-history for the session, and commits on blur/Escape as one engine
// history entry.
//
// Collab: binds directly to the element's Y.XmlFragment via y-prosemirror
// (ySyncPlugin) — the fragment is the source of truth, so there's no commit;
// the binding's observer keeps the engine record in sync. yUndoPlugin routes
// undo through the shared Y.UndoManager. Per-element fragments mean yCursorPlugin
// is skipped (it assumes one fragment per doc); remote presence is shown via
// awareness + RemotePresenceOverlay instead.
//
// Never unmounts mid-IME (composing guard on blur).

import React, { useEffect, useRef } from 'react';
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { keymap } from 'prosemirror-keymap';
import { baseKeymap, toggleMark } from 'prosemirror-commands';
import { history, undo, redo } from 'prosemirror-history';
import { ySyncPlugin, yUndoPlugin, undo as yUndo, redo as yRedo } from 'y-prosemirror';
import { slidesSchema, resolveColor, resolveFont, emptyRichText, type ThemeData, type BodyStyle, type RichTextDoc } from '@b32nio/spindle-slides-core';
import { useDeck } from '../hooks';
import { useDeckContext } from '../context/DeckContext';

const V_ALIGN = { top: 'flex-start', middle: 'center', bottom: 'flex-end' } as const;

// Inject the editor's rendering rules once: no focus outline, list markers via
// ::before (the flat PM schema has no list nodes; StaticRichText draws markers
// itself, so this keeps edit mode consistent with idle rendering).
function ensureEditorStyles(): void {
  if (typeof document === 'undefined' || document.getElementById('spindle-pm-styles')) return;
  const style = document.createElement('style');
  style.id = 'spindle-pm-styles';
  style.textContent = `
.spindle-pm-editor .ProseMirror { outline: none; white-space: pre-wrap; word-break: break-word; line-height: 1.2; }
.spindle-pm-editor .ProseMirror p { margin: 0; }
.spindle-pm-editor .ProseMirror { counter-reset: pm-list; }
.spindle-pm-editor .ProseMirror p:not([data-list="number"]) { counter-reset: pm-list; }
.spindle-pm-editor .ProseMirror p[data-list="bullet"]::before { content: "•"; position: absolute; left: 8px; opacity: 0.9; }
.spindle-pm-editor .ProseMirror p[data-list="number"] { counter-increment: pm-list; }
.spindle-pm-editor .ProseMirror p[data-list="number"]::before { content: counter(pm-list) "."; position: absolute; left: 2px; opacity: 0.9; }`;
  document.head.appendChild(style);
}

export function RichTextEditor({
  elementId,
  theme,
  bodyStyle,
  centered = false,
  cell,
}: {
  elementId: string;
  theme: ThemeData;
  bodyStyle?: BodyStyle;
  centered?: boolean;
  /** When set, edit this [row, col] cell of a table element instead of the
   *  element's own rich text. Table cells commit via snapshot (no collab v1). */
  cell?: readonly [number, number];
}): React.ReactElement {
  const deck = useDeck();
  const { editing } = useDeckContext();
  const mountRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    ensureEditorStyles();
    const mount = mountRef.current;
    const el = deck.getElement(elementId);
    // text and shape elements can hold rich text; a shape may not have any yet,
    // so fall back to an empty body rather than bailing. A table cell edits its
    // own rich text via the `cell` prop.
    if (!mount || !el) return;
    if (cell ? el.type !== 'table' : el.type !== 'text' && el.type !== 'shape') return;
    // Shapes default to centered text; a fresh (empty) shape body starts with a
    // centre-aligned paragraph so typing lands centre-centre, not top-left.
    const emptyDoc: RichTextDoc = centered
      ? { type: 'doc', content: [{ type: 'paragraph', attrs: { align: 'center' } }] }
      : emptyRichText();
    const cellDoc = cell && el.type === 'table' ? el.cells[cell[0]]?.[cell[1]]?.richText : undefined;
    const initialDoc = cellDoc ?? (el as { richText?: RichTextDoc }).richText ?? emptyDoc;

    const bold = slidesSchema.marks.bold;
    const italic = slidesSchema.marks.italic;
    const underline = slidesSchema.marks.underline;

    const handle = deck.getCollabHandle();
    // Table cells have no per-element CRDT fragment yet — commit via snapshot.
    const fragment = cell ? undefined : handle?.getElementFragment(elementId);
    const collab = !!(handle && fragment);

    let view: EditorView;

    const commit = () => {
      const doc = view.state.doc.toJSON() as RichTextDoc;
      if (cell) deck.setTableCellRichText(elementId, cell[0], cell[1], doc);
      else if (!collab) deck.setElementRichText(elementId, doc);
    };
    const exit = () => {
      commit();
      editing.setEditingId(null);
    };

    const markKeys = {
      'Mod-b': toggleMark(bold),
      'Mod-i': toggleMark(italic),
      'Mod-u': toggleMark(underline),
      Escape: () => {
        exit();
        return true;
      },
    };

    if (collab) {
      const state = EditorState.create({
        schema: slidesSchema,
        plugins: [
          ySyncPlugin(fragment!),
          yUndoPlugin({ undoManager: handle!.undoManager }),
          keymap({ ...markKeys, 'Mod-z': yUndo, 'Mod-y': yRedo, 'Shift-Mod-z': yRedo }),
          keymap(baseKeymap),
        ],
      });
      view = new EditorView(mount, { state });
      handle!.awareness.setLocalStateField('editing', { slideId: el.containerId, elementId });
    } else {
      const state = EditorState.create({
        doc: slidesSchema.nodeFromJSON(initialDoc),
        plugins: [
          history(),
          keymap({ ...markKeys, 'Mod-z': undo, 'Mod-y': redo, 'Shift-Mod-z': redo }),
          keymap(baseKeymap),
        ],
      });
      view = new EditorView(mount, { state });
      view.dispatch(view.state.tr.setSelection(TextSelection.atEnd(view.state.doc)));
    }

    editing.setView(view);

    // Remove the browser's default focus outline on the contenteditable — it
    // draws a second (mismatched) box inside the selection box. Fill the box so
    // the caret/click target matches the element frame.
    const pmDom = view.dom as HTMLElement;
    pmDom.style.outline = 'none';
    pmDom.style.minHeight = '1em';
    pmDom.style.width = '100%';
    pmDom.style.whiteSpace = 'pre-wrap';
    pmDom.style.wordBreak = 'break-word';

    // Focus on the next frame, after the pointer sequence that opened the
    // editor (the double-click's mouseup would otherwise blur the freshly
    // focused contenteditable and immediately exit). Ignore any blur that
    // fires before we've actually taken focus.
    let focused = false;
    const raf = requestAnimationFrame(() => {
      focused = true;
      view.focus();
    });

    const onBlur = () => {
      if (view.composing || !focused) return; // never commit mid-IME or pre-focus
      // Commit on blur, but DON'T leave edit mode here — a transient blur during
      // a re-render must not close the editor. Exit is driven explicitly by
      // Escape or a pointerdown outside the editor (see InteractiveSlide).
      commit();
    };
    view.dom.addEventListener('blur', onBlur);

    return () => {
      cancelAnimationFrame(raf);
      view.dom.removeEventListener('blur', onBlur);
      if (collab) {
        handle!.awareness.setLocalStateField('editing', null);
      } else if (editing.getView() === view) {
        // Commit any uncommitted edits (e.g. slide switch) before tearing down.
        const doc = view.state.doc.toJSON() as RichTextDoc;
        if (cell) deck.setTableCellRichText(elementId, cell[0], cell[1], doc);
        else deck.setElementRichText(elementId, doc);
      }
      if (editing.getView() === view) editing.setView(null);
      view.destroy();
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [elementId, cell?.[0], cell?.[1]]);

  // Bind the current theme's color slots + fonts to CSS variables the schema's
  // toDOM references, so theme-slot text colors and major/minor fonts render in
  // the editor exactly as StaticRichText resolves them.
  const themeVars: Record<string, string> = { '--font-major': theme.fonts.major, '--font-minor': theme.fonts.minor };
  for (const [slot, hex] of Object.entries(theme.colors)) themeVars[`--slot-${slot}`] = hex;

  return (
    <div
      className="spindle-pm-editor"
      onPointerDown={(e) => e.stopPropagation()}
      onDoubleClick={(e) => e.stopPropagation()}
      ref={mountRef}
      style={{
        ...(themeVars as React.CSSProperties),
        width: '100%',
        height: '100%',
        boxSizing: 'border-box',
        padding: bodyStyle?.padding ?? 8,
        display: 'flex',
        flexDirection: 'column',
        justifyContent: V_ALIGN[bodyStyle?.vAlign ?? (centered ? 'middle' : 'top')],
        fontFamily: resolveFont(bodyStyle?.fontFamily, theme),
        fontSize: bodyStyle?.fontSize ?? 18,
        fontWeight: bodyStyle?.bold ? 700 : 400,
        color: bodyStyle?.color ? resolveColor(bodyStyle.color, theme) : resolveColor({ kind: 'theme', slot: 'dk1' }, theme),
        textAlign: centered ? 'center' : 'left',
        cursor: 'text',
      }}
    />
  );
}
