import { Node as PmNode } from 'prosemirror-model';
import { EditorState } from 'prosemirror-state';
import { docsSchema } from '@b32nio/spindle-docs-core';
import {
  applyInsertColumn,
  applyDeleteColumn,
  applyInsertRow,
  applyDeleteRow,
  applyColumnResize,
  buildColumnGrid,
  logicalColumnForCell,
} from './pm-table-columns';

const schema = docsSchema;
const cell = (attrs: Record<string, unknown> = {}): PmNode =>
  schema.nodes.table_cell.create(attrs, schema.nodes.paragraph.create());
const cellText = (text: string, attrs: Record<string, unknown> = {}): PmNode =>
  schema.nodes.table_cell.create(attrs, schema.nodes.paragraph.create(null, schema.text(text)));
/** All non-empty text in the table, for content-loss checks. */
const tableText = (t: PmNode): string[] => {
  const out: string[] = [];
  t.descendants((n) => {
    if (n.isText && n.text) out.push(n.text);
  });
  return out;
};
const row = (...cells: PmNode[]): PmNode => schema.nodes.table_row.create(null, cells);
const table = (...rows: PmNode[]): PmNode => schema.nodes.table.create(null, rows);
const docWith = (t: PmNode): PmNode => schema.nodes.doc.create(null, t);

/** Independent logical-width computation to cross-check the result grid. */
function rowLogicalWidths(t: PmNode): number[] {
  const grid: boolean[][] = [];
  for (let r = 0; r < t.childCount; r++) grid[r] = grid[r] || [];
  for (let r = 0; r < t.childCount; r++) {
    let col = 0;
    const rowNode = t.child(r);
    for (let ci = 0; ci < rowNode.childCount; ci++) {
      const c = rowNode.child(ci);
      while (grid[r][col]) col++;
      const cs = (c.attrs.colspan as number) || 1;
      const rs = (c.attrs.rowspan as number) || 1;
      for (let dr = 0; dr < rs && r + dr < t.childCount; dr++) {
        grid[r + dr] = grid[r + dr] || [];
        for (let dc = 0; dc < cs; dc++) grid[r + dr][col + dc] = true;
      }
      col += cs;
    }
  }
  return grid.map((g) => g.length);
}

/** Apply a transform and return the resulting table node (doc's first child). */
function run(t: PmNode, fn: (tr: import('prosemirror-state').Transaction, table: PmNode) => boolean): PmNode {
  const state = EditorState.create({ schema, doc: docWith(t) });
  const tr = state.tr;
  // table is the doc's first child → tablePos 0
  fn(tr, tr.doc.child(0));
  return tr.doc.child(0);
}

describe('pm-table-columns — plain tables', () => {
  const plain3x3 = table(row(cell(), cell(), cell()), row(cell(), cell(), cell()), row(cell(), cell(), cell()));

  it('inserts a column, keeping every row rectangular', () => {
    const out = run(plain3x3, (tr, t) => applyInsertColumn(tr, t, 0, 1, schema));
    expect(rowLogicalWidths(out)).toEqual([4, 4, 4]);
    out.forEach((r) => expect(r.childCount).toBe(4));
  });

  it('deletes a column, keeping every row rectangular', () => {
    const out = run(plain3x3, (tr, t) => applyDeleteColumn(tr, t, 0, 1));
    expect(rowLogicalWidths(out)).toEqual([2, 2, 2]);
  });

  it('refuses to delete the last column', () => {
    const oneCol = table(row(cell()), row(cell()));
    const state = EditorState.create({ schema, doc: docWith(oneCol) });
    const tr = state.tr;
    expect(applyDeleteColumn(tr, tr.doc.child(0), 0, 0)).toBe(false);
  });
});

describe('pm-table-columns — merged tables stay rectangular (the corruption fix)', () => {
  // Row 0: a 2x2 span (cols 0-1) then a plain cell; row 1: one plain cell (cols
  // 0-1 covered by the span); row 2: three plain cells. Physical index != logical.
  const merged = table(
    row(cell({ colspan: 2, rowspan: 2 }), cell()),
    row(cell()),
    row(cell(), cell(), cell()),
  );

  it('the naive physical index is not the logical column', () => {
    const state = EditorState.create({ schema, doc: docWith(merged) });
    // Row 0's second physical cell is logical column 2, not 1.
    expect(logicalColumnForCell(state.doc.child(0), 0, 0, 1)).toBe(2);
    expect(buildColumnGrid(state.doc.child(0), 0).colCount).toBe(3);
  });

  it('insert through the span widens it once; other rows gain a real cell', () => {
    const out = run(merged, (tr, t) => applyInsertColumn(tr, t, 0, 1, schema));
    expect(rowLogicalWidths(out)).toEqual([4, 4, 4]);
    // The span went colspan 2 -> 3, still rowspan 2.
    expect(out.child(0).child(0).attrs.colspan).toBe(3);
    expect(out.child(0).child(0).attrs.rowspan).toBe(2);
    expect(out.child(0).childCount).toBe(2); // no extra cell in the merged row
    expect(out.child(2).childCount).toBe(4); // plain row gained a cell
  });

  it('append a column past the span keeps the grid rectangular', () => {
    const out = run(merged, (tr, t) => applyInsertColumn(tr, t, 0, 3, schema));
    expect(rowLogicalWidths(out)).toEqual([4, 4, 4]);
    expect(out.child(0).child(0).attrs.colspan).toBe(2); // span unchanged
  });

  it('delete the spanned column shrinks the span, not a wrong physical cell', () => {
    const out = run(merged, (tr, t) => applyDeleteColumn(tr, t, 0, 0));
    expect(rowLogicalWidths(out)).toEqual([2, 2, 2]);
    expect(out.child(0).child(0).attrs.colspan).toBe(1);
    expect(out.child(0).child(0).attrs.rowspan).toBe(2);
    expect(out.child(2).childCount).toBe(2);
  });
});

describe('pm-table-columns — span-aware column resize', () => {
  const merged = table(
    row(cell({ colspan: 2, rowspan: 2 }), cell()),
    row(cell()),
    row(cell(), cell(), cell()),
  );

  it('sets colwidth on the owning cell with an array sized to its colspan', () => {
    // Resize logical column 0, which is owned by the 2x2 span (colspan 2).
    const out = run(merged, (tr, t) => applyColumnResize(tr, t, 0, 0, 90));
    const span = out.child(0).child(0);
    expect(span.attrs.colspan).toBe(2);
    expect(span.attrs.colwidth).toEqual([90, 0]); // width set at column 0 of the span
    // Plain row's first cell (also logical col 0) gets a length-1 colwidth.
    expect(out.child(2).child(0).attrs.colwidth).toEqual([90]);
  });

  it('resizing the second logical column of a span sets the right array slot', () => {
    const out = run(merged, (tr, t) => applyColumnResize(tr, t, 0, 1, 70));
    expect(out.child(0).child(0).attrs.colwidth).toEqual([0, 70]);
  });
});

describe('pm-table-columns — span-aware row ops (no ragged grid, no content loss)', () => {
  // 2x2 span at top-left carrying text "A"; the third column of rows 0/1, and a
  // plain third row, carry identifiable text.
  const merged = table(
    row(cellText('A', { colspan: 2, rowspan: 2 }), cellText('b0')),
    row(cellText('b1')),
    row(cellText('c0'), cellText('c1'), cellText('c2')),
  );

  it('insert a row through the span widens it and keeps the grid rectangular', () => {
    const out = run(merged, (tr, t) => applyInsertRow(tr, t, 0, 0, 'after', schema));
    expect(rowLogicalWidths(out)).toEqual([3, 3, 3, 3]); // 4 rows now
    // The 2x2 span became 2x3 (covers the inserted row) — content preserved.
    expect(out.child(0).child(0).attrs.rowspan).toBe(3);
    expect(out.child(0).child(0).attrs.colspan).toBe(2);
    expect(tableText(out)).toContain('A');
  });

  it('delete a row a rowspan reaches into shrinks the span (no orphaned coverage)', () => {
    const out = run(merged, (tr, t) => applyDeleteRow(tr, t, 0, 1, schema));
    expect(rowLogicalWidths(out)).toEqual([3, 3]);
    // Span shrank from rowspan 2 to 1; all cell text survives.
    expect(out.child(0).child(0).attrs.rowspan).toBe(1);
    expect(tableText(out).sort()).toEqual(['A', 'b0', 'c0', 'c1', 'c2'].sort());
  });

  it('delete the row a rowspan ORIGINATES in relocates its content, not loses it', () => {
    const out = run(merged, (tr, t) => applyDeleteRow(tr, t, 0, 0, schema));
    expect(rowLogicalWidths(out)).toEqual([3, 3]);
    // "A" must survive (relocated down), not vanish.
    expect(tableText(out)).toContain('A');
    // b0 (rowspan-1 plain cell in the deleted row) is gone; b1 stays.
    const text = tableText(out).sort();
    expect(text).toContain('b1');
    expect(text).not.toContain('b0');
  });

  it('plain-table row insert/delete round-trips', () => {
    const plain = table(row(cell(), cell()), row(cell(), cell()));
    const grown = run(plain, (tr, t) => applyInsertRow(tr, t, 0, 0, 'after', schema));
    expect(grown.childCount).toBe(3);
    expect(rowLogicalWidths(grown)).toEqual([2, 2, 2]);
    const shrunk = run(grown, (tr, t) => applyDeleteRow(tr, t, 0, 1, schema));
    expect(shrunk.childCount).toBe(2);
  });
});
