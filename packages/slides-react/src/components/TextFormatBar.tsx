// TextFormatBar — inline text-formatting controls, rendered *inside* the main
// Toolbar row (not as a separate bar) so selecting a text element doesn't shift
// the layout. Routes each action through applyFormat, which targets the live
// editor (range formatting) when one is open, or the selected element's stored
// JSON (idle, whole-body) otherwise. Renders nothing unless a single text/shape
// element is selected or being edited.

import React, { useSyncExternalStore } from 'react';
import { Bold, Italic, Underline, Strikethrough, AlignLeft, AlignCenter, AlignRight, List, ListOrdered, Link, Highlighter } from 'lucide-react';
import type { TextFormatSpec, TextAlign, ListType } from '@b32nio/spindle-slides-core';
import { useDeck, useSelection, useEditingId } from '../hooks';
import { useDeckContext } from '../context/DeckContext';
import { applyFormat } from '../interactions/formatting';
import { cellsInSelection } from '../interactions/table-selection-store';
import { ToolbarButton, ToolbarDivider } from './toolbarUI';

const FONT_SIZES = [12, 14, 16, 18, 20, 24, 28, 32, 40, 48, 54, 66, 80];
// 'major'/'minor' follow the theme's heading/body fonts; the rest are literal.
const FONTS: Array<{ label: string; value: string }> = [
  { label: 'Heading font', value: 'major' },
  { label: 'Body font', value: 'minor' },
  { label: 'Arial', value: 'Arial' },
  { label: 'Georgia', value: 'Georgia' },
  { label: 'Times New Roman', value: 'Times New Roman' },
  { label: 'Courier New', value: 'Courier New' },
  { label: 'Verdana', value: 'Verdana' },
];
const LINE_SPACINGS = [1, 1.15, 1.5, 2];

const field: React.CSSProperties = {
  height: 30, border: '1px solid rgba(0,0,0,0.1)', borderRadius: 8,
  background: 'rgba(255,255,255,0.8)', color: '#334155', padding: '0 6px',
  fontSize: 13, fontFamily: '"Inter", sans-serif', cursor: 'pointer',
};

/** The single text element eligible for idle (whole-body) formatting, if any.
 *  A shape's text bar appears when its text is actually being edited, so a
 *  merely-selected shape shows its own fill/border bar instead of the text one. */
function idleTextTarget(deck: ReturnType<typeof useDeck>, ids: string[]): string | null {
  if (ids.length !== 1) return null;
  const el = deck.getElement(ids[0]);
  return el && el.type === 'text' ? ids[0] : null;
}

export function TextFormatBar(): React.ReactElement | null {
  const deck = useDeck();
  const { editing, tableSel } = useDeckContext();
  const selection = useSelection();
  const editingId = useEditingId();
  const sel = useSyncExternalStore(tableSel.subscribe, tableSel.getState);

  // A cell-range selection (not editing) formats every cell at once — e.g. bold
  // + color a whole header row in one action.
  const cellRange = !editingId && sel ? sel : null;
  const elementId = editingId ?? idleTextTarget(deck, selection.elementIds);
  if (!elementId && !cellRange) return null;

  const tableCells = cellRange ? { tableId: cellRange.tableId, cells: cellsInSelection(cellRange) } : null;
  const fmt = (spec: TextFormatSpec) => {
    applyFormat({ editingView: editing.getView(), deck, elementId, tableCells }, spec);
  };

  return (
    <>
      <ToolbarDivider />
      <ToolbarButton title="Bold (⌘B)" onClick={() => fmt({ toggleMark: 'bold' })}><Bold size={15} /></ToolbarButton>
      <ToolbarButton title="Italic (⌘I)" onClick={() => fmt({ toggleMark: 'italic' })}><Italic size={15} /></ToolbarButton>
      <ToolbarButton title="Underline (⌘U)" onClick={() => fmt({ toggleMark: 'underline' })}><Underline size={15} /></ToolbarButton>
      <ToolbarButton title="Strikethrough" onClick={() => fmt({ toggleMark: 'strikethrough' })}><Strikethrough size={15} /></ToolbarButton>

      <ToolbarDivider />

      <select
        title="Font"
        defaultValue="minor"
        onMouseDown={(e) => e.stopPropagation()}
        onChange={(e) => fmt({ setMark: { type: 'fontFamily', attrs: { family: e.target.value } } })}
        style={{ ...field, maxWidth: 130 }}
      >
        {FONTS.map((f) => (
          <option key={f.value} value={f.value} style={f.value !== 'major' && f.value !== 'minor' ? { fontFamily: f.value } : undefined}>{f.label}</option>
        ))}
      </select>
      <select
        title="Font size"
        defaultValue={18}
        onMouseDown={(e) => e.stopPropagation()}
        onChange={(e) => fmt({ setMark: { type: 'fontSize', attrs: { size: Number(e.target.value) } } })}
        style={field}
      >
        {FONT_SIZES.map((s) => (
          <option key={s} value={s}>{s}</option>
        ))}
      </select>
      <input
        type="color"
        title="Text color"
        onMouseDown={(e) => e.preventDefault()}
        onChange={(e) => fmt({ setMark: { type: 'textColor', attrs: { color: { kind: 'rgb', hex: e.target.value } } } })}
        style={{ ...field, width: 30, padding: 3, marginLeft: 2 }}
      />
      <Highlighter size={14} style={{ color: '#64748b', marginLeft: 4 }} aria-hidden />
      <input
        type="color"
        title="Highlight color"
        onMouseDown={(e) => e.preventDefault()}
        onChange={(e) => fmt({ setMark: { type: 'highlight', attrs: { color: { kind: 'rgb', hex: e.target.value } } } })}
        style={{ ...field, width: 30, padding: 3 }}
      />
      <ToolbarButton title="No highlight" onClick={() => fmt({ removeMark: 'highlight' })}>
        <span style={{ fontSize: 13, fontWeight: 700 }}>⦸</span>
      </ToolbarButton>

      <ToolbarDivider />

      {(['left', 'center', 'right'] as TextAlign[]).map((a) => (
        <ToolbarButton key={a} title={`Align ${a}`} onClick={() => fmt({ paragraph: { align: a } })}>
          {a === 'left' ? <AlignLeft size={15} /> : a === 'center' ? <AlignCenter size={15} /> : <AlignRight size={15} />}
        </ToolbarButton>
      ))}
      {(['bullet', 'number'] as ListType[]).map((l) => (
        <ToolbarButton key={l} title={`${l === 'bullet' ? 'Bulleted' : 'Numbered'} list`} onClick={() => fmt({ paragraph: { listType: l } })}>
          {l === 'bullet' ? <List size={15} /> : <ListOrdered size={15} />}
        </ToolbarButton>
      ))}
      <select
        title="Line spacing"
        defaultValue="1"
        onMouseDown={(e) => e.stopPropagation()}
        onChange={(e) => fmt({ paragraph: { lineHeight: Number(e.target.value) } })}
        style={field}
      >
        {LINE_SPACINGS.map((v) => (
          <option key={v} value={v}>{v.toFixed(2)}×</option>
        ))}
      </select>

      <ToolbarDivider />

      <ToolbarButton
        title="Link (⌘K)"
        onClick={() => {
          const url = window.prompt('Link URL', '');
          if (url === null) return;
          if (url.trim() === '') fmt({ removeMark: 'link' });
          else fmt({ setMark: { type: 'link', attrs: { href: url.trim() } } });
        }}
      >
        <Link size={15} />
      </ToolbarButton>
    </>
  );
}
