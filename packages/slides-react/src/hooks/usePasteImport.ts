// Paste handler for the editor root. Turns a spreadsheet paste (Excel, Numbers,
// Google Sheets) into a real table element on the active slide; otherwise falls
// back to the internal element clipboard (copy/paste of slide elements). Pastes
// aimed at a text field or the live cell/text editor are left alone so normal
// text pasting still works.

import { useCallback } from 'react';
import type React from 'react';
import { richTextFromPlainText, type NewElementSpec, type TableCell, type DeckImpl } from '@b32nio/spindle-slides-core';
import { useDeck } from './index';
import { useDeckContext } from '../context/DeckContext';
import { pasteElements, hasClipboardContent } from './useClipboard';
import { clipboardToGrid } from '../interactions/table-paste';

function isTextEntry(target: EventTarget | null): boolean {
  const el = target as HTMLElement | null;
  if (!el) return false;
  const tag = el.tagName;
  return tag === 'INPUT' || tag === 'TEXTAREA' || el.isContentEditable;
}

/** Drop a table built from a grid of plain-text cells onto the active slide,
 *  centred and sized to the content, and select it — a single undo entry. */
function insertTableFromGrid(deck: DeckImpl, grid: string[][]): void {
  const rows = grid.length;
  const cols = grid[0].length;
  const cells: TableCell[][] = grid.map((row) => row.map((txt) => ({ richText: richTextFromPlainText(txt) })));
  const { w: sw, h: sh } = deck.getSlideSize();
  const w = Math.min(sw * 0.92, Math.max(240, cols * 150));
  const h = Math.min(sh * 0.92, Math.max(80, rows * 44));
  const slideId = deck.getActiveSlideId();
  const el = deck.addElement(slideId, {
    type: 'table',
    rows,
    cols,
    cells,
    x: (sw - w) / 2,
    y: (sh - h) / 2,
    w,
    h,
  } as NewElementSpec);
  deck.setSelection({ slideId, elementIds: [el.id] });
}

export function usePasteImport(): { onPaste: (e: React.ClipboardEvent) => void } {
  const deck = useDeck();
  const { editing } = useDeckContext();

  const onPaste = useCallback(
    (e: React.ClipboardEvent) => {
      // Editing text (a cell, a text box, or any input) → let the editor paste.
      if (isTextEntry(e.target) || editing.getEditingId()) return;
      const cd = e.clipboardData;
      if (!cd) return;

      const grid = clipboardToGrid({ html: cd.getData('text/html'), text: cd.getData('text/plain') });
      if (grid) {
        e.preventDefault();
        insertTableFromGrid(deck, grid);
        return;
      }
      // Not tabular → paste any copied slide elements from the internal buffer.
      if (hasClipboardContent()) {
        e.preventDefault();
        pasteElements(deck);
      }
    },
    [deck, editing]
  );

  return { onPaste };
}
