import { docsSchema } from '@b32nio/spindle-docs-core';
import { proseMirrorToFlowBlocks } from './pm-to-blocks';
import { cellBorderCss, cellLayoutCss, tableCellPlaces } from './table-format';
import type { TableBlock, TableCell } from './flow-blocks';

const cell = (extra: Partial<TableCell> = {}): TableCell => ({ id: 'c', blocks: [], ...extra });
const table = (rows: TableCell[][], extra: Partial<TableBlock> = {}): TableBlock => ({
  kind: 'table',
  id: 't',
  rows: rows.map((cells, i) => ({ id: `r${i}`, cells })),
  ...extra,
});

describe('tableCellPlaces', () => {
  it('skips grid slots taken by a cell that spans rows and counts spanned columns', () => {
    const t = table([[cell({ colspan: 2 }), cell()], [cell({ rowspan: 2 }), cell(), cell()], [cell(), cell()]]);
    const places = tableCellPlaces(t);
    expect(places[0].map((p) => [p.col, p.colspan])).toEqual([[0, 2], [2, 1]]);
    expect(places[1].map((p) => p.col)).toEqual([0, 1, 2]);
    // Row 2: column 0 is taken by the rowspan from row 1.
    expect(places[2].map((p) => p.col)).toEqual([1, 2]);
    expect(places[0][0]).toMatchObject({ rows: 3, cols: 3 });
  });
});

describe('cellBorderCss', () => {
  it('draws the default grid when neither the table nor the cell has borders', () => {
    const t = table([[cell()]]);
    expect(cellBorderCss(t, t.rows[0].cells[0], tableCellPlaces(t)[0][0], 1)).toBe('border: 1px solid #dadce0;');
  });

  it('uses the table outer lines on the edges and the inside lines between cells', () => {
    const t = table([[cell(), cell()], [cell(), cell()]], {
      borders: { top: { style: 'single', widthPx: 2, color: '#ff0000' }, insideV: { style: 'dashed' }, insideH: { style: 'dotted', color: '#00ff00' } },
    });
    const places = tableCellPlaces(t);
    const css = (r: number, c: number) => cellBorderCss(t, t.rows[r].cells[c], places[r][c], 1);
    expect(css(0, 0)).toBe('border-top: 2px single #ff0000;border-right: 1px dashed #000000;border-bottom: 1px dotted #00ff00;border-left: none;');
    expect(css(1, 1)).toBe('border-top: 1px dotted #00ff00;border-right: none;border-bottom: none;border-left: 1px dashed #000000;');
  });

  it('draws no lines for a table whose borders object is empty', () => {
    const t = table([[cell()]], { borders: {} });
    expect(cellBorderCss(t, t.rows[0].cells[0], tableCellPlaces(t)[0][0], 1)).toBe('border-top: none;border-right: none;border-bottom: none;border-left: none;');
  });

  it("lets a cell's own side win and widens a double line so both strokes show; scales widths", () => {
    const t = table([[cell({ borders: { bottom: { style: 'double' } } })]], { borders: { bottom: { style: 'single' } } });
    expect(cellBorderCss(t, t.rows[0].cells[0], tableCellPlaces(t)[0][0], 2)).toContain('border-bottom: 3px double #000000;');
  });

  it('uses the default line for sides a cell leaves unset when the table has no borders', () => {
    const t = table([[cell({ borders: { top: { style: 'none' } } })]]);
    expect(cellBorderCss(t, t.rows[0].cells[0], tableCellPlaces(t)[0][0], 1)).toBe(
      'border-top: none;border-right: 1px solid #dadce0;border-bottom: 1px solid #dadce0;border-left: 1px solid #dadce0;'
    );
  });

  it('never emits an unsafe colour', () => {
    const t = table([[cell({ borders: { top: { style: 'single', color: 'red;background:url(x)' } } })]], { borders: {} });
    expect(cellBorderCss(t, t.rows[0].cells[0], tableCellPlaces(t)[0][0], 1)).toContain('border-top: 1px single #000000;');
  });
});

describe('cellLayoutCss', () => {
  it('aligns vertically and sums the widths of the spanned columns', () => {
    expect(cellLayoutCss(cell({ verticalAlign: 'center', colwidth: [100, 50] }), 2)).toBe('vertical-align: middle;width: 300px;');
    expect(cellLayoutCss(cell({ verticalAlign: 'bottom' }), 1)).toBe('vertical-align: bottom;');
    expect(cellLayoutCss(cell(), 1)).toBe('vertical-align: top;');
  });
});

describe('proseMirrorToFlowBlocks table formatting', () => {
  it('carries borders, vertical alignment and shading from the table nodes to the flow block', () => {
    const s = docsSchema;
    const para = () => s.nodes.paragraph.create();
    const c = s.nodes.table_cell.create({ borders: { top: { style: 'single' } }, verticalAlign: 'center', backgroundColor: '#ddeeff', colwidth: [96] }, para());
    const t = s.nodes.table.create({ borders: { left: { style: 'double' } } }, [s.nodes.table_row.create({ repeatHeader: true, cantSplit: true }, [c])]);
    const block = proseMirrorToFlowBlocks(s.nodes.doc.create(null, t))[0] as TableBlock;
    expect(block.borders).toEqual({ left: { style: 'double' } });
    expect(block.rows[0]).toMatchObject({ repeatHeader: true, cantSplit: true });
    expect(block.rows[0].cells[0]).toMatchObject({ borders: { top: { style: 'single' } }, verticalAlign: 'center', backgroundColor: '#ddeeff', colwidth: [96] });
  });
});
