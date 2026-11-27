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
    expect(html.match(/aria-label="[^"]*not updated when the cells it uses change[^"]*"/)).not.toBeNull();
    expect(html.match(/title="[^"]*not updated when the cells it uses change[^"]*"/)).not.toBeNull();
  });

  it('shows nothing for an ordinary formula, a plain value, an empty cell or no cell', () => {
    for (const cell of [{ row: 0, col: 1 }, { row: 0, col: 0 }, { row: 9, col: 9 }, null]) {
      expect(render(workbook, cell)).not.toContain('formula-cached-only');
    }
  });
});
