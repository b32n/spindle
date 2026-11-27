import React from 'react';
import { renderToStaticMarkup } from 'react-dom/server';
import { WorkbookImpl } from '@b32nio/spindle-sheets-core';
import { WorkbookProvider } from '../context/WorkbookContext';
import { CACHED_ONLY_NOTICE, FormulaBar } from './FormulaBar';

function render(workbook: WorkbookImpl, activeCell: { row: number; col: number } | null): string {
  return renderToStaticMarkup(
    <WorkbookProvider workbook={workbook}>
      <FormulaBar activeCell={activeCell} />
    </WorkbookProvider>
  );
}

describe('FormulaBar cached-only indicator', () => {
  // legacy (numeric) cell keys: rows and columns get their ids as the cells load
  const workbook = new WorkbookImpl('wb', 'WB');
  workbook.setData({
    id: 'wb',
    name: 'WB',
    activeSheetId: 's',
    defaultRowHeight: 20,
    defaultColWidth: 100,
    stylePool: {},
    formatPool: {},
    sheets: [
      {
        id: 's',
        name: 'S',
        rowCount: 100,
        colCount: 26,
        config: {},
        cells: [
          { key: '0:0', cell: { value: 2 } },
          { key: '0:1', cell: { value: 4, formula: '=A1*2' } },
          { key: '4:4', cell: { value: 99, formula: '=_xlfn.LET(_xlpm.x,1,_xlpm.x)', cachedOnly: true } },
        ],
      },
    ],
  });

  it('shows a labelled, described note for a cell whose formula the engine cannot evaluate', () => {
    const html = render(workbook, { row: 4, col: 4 });
    expect(html).toContain('data-testid="formula-cached-only"');
    expect(html).toContain('role="note"');
    expect(html).toContain('tabindex="0"');
    expect(html).toContain('>Saved value<');
    // the staleness caveat is in the accessible name and the tooltip
    expect(CACHED_ONLY_NOTICE).toMatch(/not updated when the cells it uses change/);
    // references in the text follow inserted and deleted rows and columns now (S4b), so the note no longer says they do not
    expect(CACHED_ONLY_NOTICE).not.toMatch(/adjust/);
    expect(CACHED_ONLY_NOTICE).toMatch(/kept and saved unchanged/);
    expect(html.match(/aria-label="[^"]*not updated when the cells it uses change[^"]*"/)).not.toBeNull();
    expect(html.match(/title="[^"]*not updated when the cells it uses change[^"]*"/)).not.toBeNull();
  });

  it('takes the note colours from the host\'s CSS variables, with defaults that keep a contrast of at least 4.5:1', () => {
    const html = render(workbook, { row: 4, col: 4 });
    const fg = /color:var\(--spindle-note-fg, (#[0-9a-f]{6})\)/.exec(html)?.[1];
    const bg = /background-color:var\(--spindle-note-bg, (#[0-9a-f]{6})\)/.exec(html)?.[1];
    expect(html).toContain('border:1px solid var(--spindle-note-border, #b45309)');
    expect(fg).toBeDefined();
    expect(bg).toBeDefined();
    // WCAG 2.x contrast ratio of the default colours (a host that overrides the variables for a dark theme sets a pair of its own)
    const luminance = (hex: string): number => {
      const [r, g, b] = [1, 3, 5].map((i) => parseInt(hex.slice(i, i + 2), 16) / 255).map((c) => (c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4));
      return 0.2126 * r + 0.7152 * g + 0.0722 * b;
    };
    const contrast = (a: string, b: string): number => {
      const [hi, lo] = [luminance(a), luminance(b)].sort((x, y) => y - x);
      return (hi + 0.05) / (lo + 0.05);
    };
    expect(contrast(fg as string, bg as string)).toBeGreaterThanOrEqual(4.5);
    // about 6.4:1 for the default pair (not the 7:1 of WCAG AAA)
    expect(contrast(fg as string, bg as string)).toBeLessThan(7);
    // the border is a non-text element: WCAG 1.4.11 asks for 3:1 against what it sits on
    expect(contrast('#b45309', bg as string)).toBeGreaterThanOrEqual(3);
    // no fixed colour is left in the note's own style
    expect(html).not.toMatch(/formula-cached-only[^>]*(color|background-color):#/);
  });

  it('shows nothing for an ordinary formula, a plain value, an empty cell or no cell', () => {
    for (const cell of [{ row: 0, col: 1 }, { row: 0, col: 0 }, { row: 9, col: 9 }, null]) {
      expect(render(workbook, cell)).not.toContain('formula-cached-only');
    }
  });
});
