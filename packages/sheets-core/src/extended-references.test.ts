// Workbook-level behaviour of the reference forms added to the formula engine: whole-column/row
// ranges, 3D references, range operators over computed references, union, intersection, `@`
// and the spill operator. Covers evaluation, the formula text the engine keeps, dependency
// tracking and recalculation, and what structural edits (insert/delete rows and columns,
// sheet rename/delete/reorder, fill) do to each form.

import * as fs from 'fs';
import * as path from 'path';
import { WorkbookImpl } from './workbook';
import { adjustFormula } from './formula-parser/formula-adjust';
import type { CellResolver } from './types';

function colIndex(label: string): number {
  let col = 0;
  for (const ch of label) col = col * 26 + (ch.charCodeAt(0) - 64);
  return col - 1;
}

/** "B3" -> [row 2, col 1]. */
function at(addr: string): [number, number] {
  const m = /^([A-Z]+)(\d+)$/.exec(addr) as RegExpExecArray;
  return [Number(m[2]) - 1, colIndex(m[1])];
}

class Book {
  wb = new WorkbookImpl('wb', 'WB');

  constructor(sheets = ['Sheet1', 'Sheet2', 'Sheet3']) {
    this.wb.renameSheet(this.wb.getSheet().id, sheets[0]);
    for (const name of sheets.slice(1)) this.wb.addSheet(name);
  }

  id(sheet?: string): string | undefined {
    return sheet ? this.wb.getSheetIdByName(sheet) : undefined;
  }

  set(addr: string, value: unknown, sheet?: string): this {
    const [r, c] = at(addr);
    this.wb.setCellValue(this.id(sheet), r, c, value);
    return this;
  }

  fill(col: string, values: unknown[], sheet?: string): this {
    values.forEach((v, i) => this.set(`${col}${i + 1}`, v, sheet));
    return this;
  }

  f(addr: string, formula: string, sheet?: string): this {
    const [r, c] = at(addr);
    this.wb.setFormula(this.id(sheet), r, c, formula);
    return this;
  }

  v(addr: string, sheet?: string): unknown {
    const [r, c] = at(addr);
    return this.wb.getCellCalculatedValue(this.id(sheet), r, c);
  }

  /** The value as stored on the cell (what the grid shows). */
  shown(addr: string, sheet?: string): unknown {
    const [r, c] = at(addr);
    return this.wb.getCell(this.id(sheet), r, c)?.value;
  }

  text(addr: string, sheet?: string): string | undefined {
    const [r, c] = at(addr);
    return this.wb.getCell(this.id(sheet), r, c)?.formula;
  }

  sheet(name?: string) {
    return this.wb.getSheet(this.id(name));
  }
}

describe('probe 4 oracle (LibreOffice-calculated values)', () => {
  const table = JSON.parse(fs.readFileSync(path.join(__dirname, 'formula-syntax-oracle.json'), 'utf8')) as {
    rows: Array<{ editorFormula: string; fileFormula: string; note: string; expected: unknown; cached: unknown }>;
  };

  /** The data layout of fixtures/xlsx/formula-syntax.xlsx, formulas in G2.. as the file has them. */
  function probeBook(): Book {
    const b = new Book(['Sheet1', 'Sheet2', 'Q1 Data']);
    b.fill('A', [1, 2, 3, 4, 5]).set('B1', 10).set('C1', 20);
    b.f('J1', '=SEQUENCE(3)');
    b.set('A1', 100, 'Sheet2').set('A1', 7, 'Q1 Data');
    return b;
  }

  it('covers all 23 forms of the probe', () => {
    expect(table.rows).toHaveLength(23);
  });

  const cases: Array<[number, string, string]> = table.rows.map((r, i) => [i, r.editorFormula, r.note]);

  it.each(cases)('row %i: =%s (%s)', (i, editorFormula) => {
    const b = probeBook();
    b.f(`G${i + 2}`, `=${editorFormula}`);
    const expected = table.rows[i].expected;
    const actual = b.v(`G${i + 2}`);
    if (typeof expected === 'number') expect(actual).toBeCloseTo(expected, 10);
    else expect(actual).toBe(expected);
  });

  it('keeps the formulas of the probe unchanged in the formula bar', () => {
    const b = probeBook();
    table.rows.forEach((row, i) => {
      b.f(`G${i + 2}`, `=${row.editorFormula}`);
    });
    // Canonical spellings print back exactly.
    for (const [i, row] of table.rows.entries()) {
      expect(b.text(`G${i + 2}`)).toBe(`=${row.editorFormula}`);
    }
  });
});

describe('whole-column and whole-row references in a workbook', () => {
  it('sums a column of any size and treats the other columns as separate', () => {
    const b = new Book().fill('A', [1, 2, 3]).fill('B', [10, 20]);
    b.f('D20', '=SUM(A:A)').f('D21', '=SUM(B:B)').f('D22', '=SUM(A:B)').f('D23', '=SUM(1:1)').f('D24', '=SUM(2:3)');
    expect([b.v('D20'), b.v('D21'), b.v('D22'), b.v('D23'), b.v('D24')]).toEqual([6, 30, 36, 11, 2 + 20 + 3]);
  });

  it('does not count formulas of its own column twice or style-only cells', () => {
    const b = new Book().fill('A', [1, 2]);
    b.wb.setCell(undefined, 50, 0, { styleId: undefined }); // an empty cell far down
    b.f('B1', '=ROWS(A:A)').f('C1', '=SUM(A:A)');
    expect(b.v('B1')).toBe(1048576);
    expect(b.v('C1')).toBe(3);
  });

  it('includes formula results and spilled values in the used range', () => {
    const b = new Book().f('A1', '=SEQUENCE(4)').f('C1', '=SUM(A:A)').f('C2', '=COUNT(A:A)');
    expect(b.v('C1')).toBe(10);
    expect(b.v('C2')).toBe(4);
    b.f('D10', '=SUM(1:3)');
    expect(b.v('D10')).toBe(1 + 2 + 3 + 10 + 4); // rows 1-3: A1:A3 spill, C1 (10), C2 (4)
  });

  it('spills a whole column as an array and reads other sheets', () => {
    const b = new Book().fill('A', [5, 6, 7]).set('A1', 1, 'Sheet2').set('A2', 2, 'Sheet2');
    b.f('C1', '=A:A');
    expect([b.v('C1'), b.v('C2'), b.v('C3')]).toEqual([5, 6, 7]);
    b.f('E1', '=SUM(Sheet2!A:A)').f('E2', "=SUM('Sheet2'!1:1)");
    expect([b.v('E1'), b.v('E2')]).toEqual([3, 1]);
  });

  it('keeps one range rectangle as the dependency, never a cell per row', () => {
    const b = new Book().fill('A', [1, 2, 3]);
    b.f('D1', '=SUM(A:A)');
    const [r, c] = at('D1');
    const node = b.wb.getFormulaGraph().nodes.get(
      (b.sheet().getRowId(r) as string) + ':' + (b.sheet().getColId(c) as string)
    );
    expect(node?.dependencies.size).toBe(0);
    expect(node?.rangeRects).toHaveLength(1);
    expect(node?.rangeRects[0]).toMatchObject({ axis: 'column' });
    expect(node?.rangeRects[0].startKey.startsWith('*:')).toBe(true);
    expect(b.wb.getCell(undefined, r, c)?.formulaAst?.rangeRef).toBeUndefined();
  });

  it('stores the range unbounded in the AST', () => {
    const b = new Book().fill('A', [1]).f('D1', '=SUM(B:C)').f('D2', '=SUM($2:$3)');
    const col = b.wb.getCell(undefined, 0, 3)?.formulaAst?.args?.[0].rangeRef;
    expect(col?.whole).toBe('column');
    expect(col?.start.rowId).toBe('*');
    expect(col?.end.rowId).toBe('*');
    expect(col?.start.colId).toBe(b.sheet().getColId(1));
    expect(col?.end.colId).toBe(b.sheet().getColId(2));
    const row = b.wb.getCell(undefined, 1, 3)?.formulaAst?.args?.[0].rangeRef;
    expect(row?.whole).toBe('row');
    expect(row?.start.colId).toBe('*');
    expect(row?.start.rowAbsolute).toBe(true);
  });

  it('prints #REF! for a whole-column or whole-row reference to a deleted sheet', () => {
    const b = new Book().set('A1', 5, 'Sheet2');
    b.f('A1', '=SUM(Sheet2!A:A)').f('A2', '=SUM(Sheet2!1:1)');
    expect(b.v('A1')).toBe(5);
    b.wb.deleteSheet(b.id('Sheet2') as string);
    expect([b.shown('A1'), b.shown('A2')]).toEqual(['#REF!', '#REF!']);
    expect([b.text('A1'), b.text('A2')]).toEqual(['=SUM(#REF!)', '=SUM(#REF!)']);
  });

  it('follows a renamed sheet in whole ranges and plain ranges alike', () => {
    const b = new Book().set('A1', 5, 'Sheet2').set('B2', 6, 'Sheet2');
    b.f('A1', '=SUM(Sheet2!A:A)+SUM(Sheet2!A1:B2)+SUM(Sheet2!2:2)');
    b.wb.renameSheet(b.id('Sheet2') as string, 'Data 2');
    expect(b.text('A1')).toBe("=SUM('Data 2'!A:A)+SUM('Data 2'!A1:B2)+SUM('Data 2'!2:2)");
    expect(b.v('A1')).toBe(5 + 11 + 6);
  });

  it('reads a spill that straddles the columns or rows of a whole range', () => {
    const b = new Book().f('A1', '={1,2,3;4,5,6}');
    b.f('E1', '=SUM(B:B)').f('E2', '=SUM(C:D)').f('E3', '=SUM(2:2)').f('E4', '=SUM(1:1)');
    expect(b.v('E1')).toBe(2 + 5);
    expect(b.v('E2')).toBe(3 + 6);
    expect(b.v('E3')).toBe(4 + 5 + 6 + 9); // row 2: the spill (A2:C2) and E2 itself holds 9
    expect(b.v('E4')).toBe(1 + 2 + 3 + 7); // row 1: the spill (A1:C1) and E1 holds 7
  });

  describe('recalculation', () => {
    it('recomputes when a cell in the column changes, and only then', () => {
      const b = new Book().fill('A', [1, 2, 3]).fill('B', [10, 20, 30]);
      b.f('D1', '=SUM(A:A)');
      const evaluate = jest.spyOn((b.wb as unknown as { formulaParser: { evaluate: () => unknown } }).formulaParser, 'evaluate');
      b.set('A9', 100);
      expect(b.v('D1')).toBe(106);
      expect(evaluate).toHaveBeenCalledTimes(1);
      evaluate.mockClear();
      b.set('B2', 99); // another column: nothing to recompute
      expect(evaluate).not.toHaveBeenCalled();
      expect(b.v('D1')).toBe(106);
      b.set('A9', null);
      expect(b.v('D1')).toBe(6);
    });

    it('recomputes a whole-row sum when a cell in the row changes, and only then', () => {
      const b = new Book().fill('A', [1, 2, 3]).set('B1', 10);
      b.f('D5', '=SUM(1:1)');
      expect(b.v('D5')).toBe(11);
      const evaluate = jest.spyOn((b.wb as unknown as { formulaParser: { evaluate: () => unknown } }).formulaParser, 'evaluate');
      b.set('Z1', 5);
      expect(b.v('D5')).toBe(16);
      evaluate.mockClear();
      b.set('B2', 99);
      expect(evaluate).not.toHaveBeenCalled();
    });

    it('follows edits on another sheet', () => {
      const b = new Book().set('A1', 1, 'Sheet2');
      b.f('A1', '=SUM(Sheet2!A:A)');
      expect(b.v('A1')).toBe(1);
      b.set('A5', 4, 'Sheet2');
      expect(b.v('A1')).toBe(5);
      b.set('A5', 4, 'Sheet3'); // a same-position cell on a third sheet does not count
      expect(b.v('A1')).toBe(5);
    });

    it('chains through formulas', () => {
      const b = new Book().fill('A', [1, 2]);
      b.f('B1', '=SUM(A:A)').f('C1', '=B1*2').f('A3', '=C1');
      // A3 reads C1, which reads B1, which reads column A including A3: a cycle (0 in Excel's default mode).
      expect(typeof b.v('A3')).toBe('number');
    });
  });

  describe('structural edits', () => {
    it('keeps a whole-column reference valid when rows are inserted or deleted', () => {
      const b = new Book().fill('A', [1, 2, 3, 4]);
      b.f('C1', '=SUM(A:A)').f('C2', '=SUM(A:A)');
      b.sheet().insertRows(0, 2);
      expect(b.text('C3')).toBe('=SUM(A:A)');
      expect(b.v('C3')).toBe(10);
      b.sheet().insertRows(3, 1);
      expect(b.v('C3')).toBe(10);
      b.sheet().deleteRows(2, 1); // row 3: holds A1's old row? deletes the first data row
      expect(b.text('C4')).toBe('=SUM(A:A)');
      expect(b.v('C4')).toBe(9);
      b.sheet().deleteRows(0, 2);
      expect(b.v('C2')).toBeDefined();
      expect(b.text('C2')).toBe('=SUM(A:A)');
    });

    it('follows its column when columns are inserted before it', () => {
      const b = new Book().fill('A', [1, 2, 3]);
      b.f('D1', '=SUM(A:A)').f('D2', '=SUM($A:$A)').f('D3', '=SUM(A:C)');
      b.sheet().insertCols(0, 2);
      expect(b.text('F1')).toBe('=SUM(C:C)');
      expect(b.text('F2')).toBe('=SUM($C:$C)');
      expect(b.text('F3')).toBe('=SUM(C:E)');
      expect(b.v('F1')).toBe(6);
      expect(b.v('F3')).toBe(6);
    });

    it('grows when a column is inserted inside the span', () => {
      const b = new Book().fill('A', [1]).fill('B', [10]).fill('C', [100]);
      b.f('E1', '=SUM(A:C)');
      b.sheet().insertCols(1, 1);
      expect(b.text('F1')).toBe('=SUM(A:D)');
      b.set('B1', 1000);
      expect(b.v('F1')).toBe(1111);
    });

    it('becomes #REF! when its column is deleted', () => {
      const b = new Book().fill('A', [1, 2]).fill('B', [10, 20]);
      b.f('D1', '=SUM(A:A)').f('D2', '=SUM(B:B)').f('D3', '=SUM(A:B)');
      b.sheet().deleteCols(0, 1);
      expect(b.shown('C1')).toBe('#REF!');
      expect(b.text('C1')).toBe('=SUM(#REF!)');
      expect(b.v('C2')).toBe(30);
      expect(b.text('C2')).toBe('=SUM(A:A)');
      expect(b.shown('C3')).toBe(30); // the span keeps its surviving column (Excel: =SUM(A:B) without A is =SUM(A:A))
      expect(b.text('C3')).toBe('=SUM(A:A)');
    });

    it('is unaffected by inserted or deleted columns for a whole-row reference, and follows its row', () => {
      const b = new Book().fill('A', [1, 2, 3]).set('B1', 10);
      b.f('E3', '=SUM(1:1)').f('E4', '=SUM(2:3)').f('E5', '=SUM($1:$1)');
      b.sheet().insertCols(0, 1);
      expect(b.text('F3')).toBe('=SUM(1:1)');
      expect(b.v('F3')).toBe(11);
      b.sheet().insertRows(0, 1);
      expect(b.text('F4')).toBe('=SUM(2:2)');
      expect(b.text('F5')).toBe('=SUM(3:4)');
      expect(b.text('F6')).toBe('=SUM($2:$2)');
      expect(b.v('F4')).toBe(11);
    });

    it('becomes #REF! when its row is deleted', () => {
      const b = new Book().fill('A', [1, 2, 3]);
      b.f('D4', '=SUM(1:1)').f('D5', '=SUM(2:2)');
      b.sheet().deleteRows(0, 1);
      expect(b.shown('D3')).toBe('#REF!');
      expect(b.text('D3')).toBe('=SUM(#REF!)');
      expect(b.v('D4')).toBe(2);
      expect(b.text('D4')).toBe('=SUM(1:1)');
    });

    it('keeps working after a reload of the data', () => {
      const b = new Book().fill('A', [1, 2, 3]).set('Z2', 7);
      b.f('C9', '=SUM(A:A)').f('C10', '=SUM(2:2)');
      b.sheet().insertRows(0, 1);
      const reloaded = new WorkbookImpl('copy', 'Copy');
      reloaded.setData(JSON.parse(JSON.stringify(b.wb.getData())));
      expect(reloaded.getCellCalculatedValue(undefined, 9, 2)).toBe(6);
      expect(reloaded.getCell(undefined, 9, 2)?.formula).toBe('=SUM(A:A)');
      expect(reloaded.getCell(undefined, 10, 2)?.formula).toBe('=SUM(3:3)');
      // Row 3 is the row that held row 2 before the insert: A2 (2) and Z2 (7).
      expect(reloaded.getCellCalculatedValue(undefined, 10, 2)).toBe(9);
    });
  });

  describe('copy and fill', () => {
    const rebase = (b: Book, formula: string, from: string, to: string) => {
      const [fr, fc] = at(from);
      const [tr, tc] = at(to);
      return adjustFormula(formula, b.wb, undefined, fr, fc, tr, tc);
    };

    it('moves a relative whole-column reference along columns only', () => {
      const b = new Book();
      expect(rebase(b, '=SUM(A:A)', 'B1', 'C1')).toBe('=SUM(B:B)');
      expect(rebase(b, '=SUM(A:A)', 'B1', 'B9')).toBe('=SUM(A:A)');
      expect(rebase(b, '=SUM($A:$A)', 'B1', 'F9')).toBe('=SUM($A:$A)');
      expect(rebase(b, '=SUM(A:C)', 'E1', 'F1')).toBe('=SUM(B:D)');
      expect(rebase(b, '=SUM(B:B)', 'C1', 'B1')).toBe('=SUM(A:A)');
    });

    it('moves a relative whole-row reference along rows only', () => {
      const b = new Book();
      expect(rebase(b, '=SUM(1:1)', 'A2', 'A3')).toBe('=SUM(2:2)');
      expect(rebase(b, '=SUM(1:1)', 'A2', 'F2')).toBe('=SUM(1:1)');
      expect(rebase(b, '=SUM($1:$1)', 'A2', 'A9')).toBe('=SUM($1:$1)');
      expect(rebase(b, '=SUM(1:3)', 'A4', 'A5')).toBe('=SUM(2:4)');
    });

    it('is #REF! when a copy moves the reference off the sheet', () => {
      const b = new Book();
      expect(rebase(b, '=SUM(B:B)', 'C1', 'A1')).toBe('=SUM(#REF!)');
      expect(rebase(b, '=SUM(2:2)', 'A3', 'A1')).toBe('=SUM(#REF!)');
    });

    it('rebases percent, array constants, @ and # operands', () => {
      const b = new Book();
      expect(rebase(b, '=A1%', 'B1', 'B2')).toBe('=A2%');
      expect(rebase(b, '=SUM({1,2;3,4})+A1', 'B1', 'C1')).toBe('=SUM({1,2;3,4})+B1');
      expect(rebase(b, '=@A1:A3', 'B1', 'C1')).toBe('=@B1:B3');
      expect(rebase(b, '=SUM(A1#)', 'B1', 'B2')).toBe('=SUM(A2#)');
      expect(rebase(b, '=SUM((A1:A2,C1))', 'E1', 'E2')).toBe('=SUM((A2:A3,C2))');
      expect(rebase(b, '=SUM(A1:A3 A2:A4)', 'E1', 'F1')).toBe('=SUM(B1:B3 B2:B4)');
      expect(rebase(b, '=SUM(A1:INDEX(A1:A5,3))', 'E1', 'E2')).toBe('=SUM(A2:INDEX(A2:A6,3))');
    });

    it('keeps a 3D reference at the same positions unless the cell part is relative', () => {
      const b = new Book();
      expect(rebase(b, '=SUM(Sheet1:Sheet3!A1)', 'B1', 'B2')).toBe('=SUM(Sheet1:Sheet3!A2)');
      expect(rebase(b, '=SUM(Sheet1:Sheet3!A1)', 'B1', 'C1')).toBe('=SUM(Sheet1:Sheet3!B1)');
      expect(rebase(b, '=SUM(Sheet1:Sheet3!$A$1)', 'B1', 'C5')).toBe('=SUM(Sheet1:Sheet3!$A$1)');
      expect(rebase(b, '=SUM(Sheet1:Sheet3!A1:B2)', 'D1', 'D3')).toBe('=SUM(Sheet1:Sheet3!A3:B4)');
      expect(rebase(b, '=SUM(Sheet1:Sheet3!A:A)', 'D1', 'E1')).toBe('=SUM(Sheet1:Sheet3!B:B)');
      expect(rebase(b, '=SUM(Sheet1:Sheet3!A1)', 'B2', 'B1')).toBe('=SUM(#REF!)');
      expect(rebase(b, '=SUM(Sheet1:Sheet3!B:B)', 'C1', 'A1')).toBe('=SUM(#REF!)');
      expect(rebase(b, '=SUM(Sheet1:Sheet3!2:2)', 'A3', 'A1')).toBe('=SUM(#REF!)');
    });
  });
});

describe('percent in cells', () => {
  it('evaluates and keeps its text', () => {
    const b = new Book().set('A1', 200);
    b.f('B1', '=A1*50%').f('B2', '=A1+10%').f('B3', '=10%^2').f('B4', '=-A1%');
    expect(b.v('B1')).toBe(100);
    expect(b.v('B2')).toBeCloseTo(200.1);
    expect(b.v('B3')).toBeCloseTo(0.01);
    expect(b.v('B4')).toBe(-2);
    expect(['B1', 'B2', 'B3', 'B4'].map((a) => b.text(a))).toEqual(['=A1*50%', '=A1+10%', '=10%^2', '=-A1%']);
  });

  it('applies to a range and spills', () => {
    const b = new Book().fill('A', [100, 200]);
    b.f('B1', '=A1:A2%');
    expect([b.v('B1'), b.v('B2')]).toEqual([1, 2]);
  });

  it('is #VALUE! for text', () => {
    const b = new Book().set('A1', 'abc');
    b.f('B1', '=A1%');
    expect(b.shown('B1')).toBe('#VALUE!');
  });
});

describe('array constants in cells', () => {
  it('spills rows across and down', () => {
    const b = new Book();
    b.f('A1', '={1,2,3}').f('A3', '={1,2;3,4}').f('A6', '={1;2;3}');
    expect([b.v('A1'), b.v('B1'), b.v('C1')]).toEqual([1, 2, 3]);
    expect([b.v('A3'), b.v('B3'), b.v('A4'), b.v('B4')]).toEqual([1, 2, 3, 4]);
    expect([b.v('A6'), b.v('A7'), b.v('A8')]).toEqual([1, 2, 3]);
  });

  it('is #SPILL! when something is in the way', () => {
    const b = new Book().set('B1', 'blocker');
    b.f('A1', '={1,2,3}');
    expect(b.shown('A1')).toBe('#SPILL!');
  });

  it('works in functions and keeps its text', () => {
    const b = new Book().set('A1', 5);
    b.f('C1', '=SUM({1,2;3,4})').f('C2', '=A1*{1,2,3}').f('C3', '=SUMPRODUCT({1,2,3},{4,5,6})').f('C4', '={-1.5,"a b",TRUE,#N/A}');
    expect(b.v('C1')).toBe(10);
    expect([b.v('C2'), b.v('D2'), b.v('E2')]).toEqual([5, 10, 15]);
    expect(b.v('C3')).toBe(32);
    expect(b.text('C4')).toBe('={-1.5,"a b",TRUE,#N/A}');
    expect([b.v('C4'), b.v('D4'), b.v('E4'), b.v('F4')]).toEqual([-1.5, 'a b', true, '#N/A']);
  });
});

describe('3D references in a workbook', () => {
  function threeSheets(): Book {
    const b = new Book();
    b.set('A1', 1, 'Sheet1').set('B2', 10, 'Sheet1');
    b.set('A1', 2, 'Sheet2').set('B2', 20, 'Sheet2');
    b.set('A1', 3, 'Sheet3').set('B2', 30, 'Sheet3');
    return b;
  }

  it('sums across the span and keeps its text', () => {
    const b = threeSheets();
    b.f('D1', '=SUM(Sheet1:Sheet3!A1)').f('D2', '=SUM(Sheet1:Sheet3!A1:B2)').f('D3', '=AVERAGE(Sheet1:Sheet2!B2)');
    expect([b.v('D1'), b.v('D2'), b.v('D3')]).toEqual([6, 66, 15]);
    expect([b.text('D1'), b.text('D2'), b.text('D3')]).toEqual([
      '=SUM(Sheet1:Sheet3!A1)',
      '=SUM(Sheet1:Sheet3!A1:B2)',
      '=AVERAGE(Sheet1:Sheet2!B2)',
    ]);
  });

  it('recomputes when a cell on any sheet of the span changes', () => {
    const b = threeSheets();
    b.f('D1', '=SUM(Sheet1:Sheet3!A1)');
    b.set('A1', 100, 'Sheet2');
    expect(b.v('D1')).toBe(104);
    b.set('A1', 7, 'Sheet3');
    expect(b.v('D1')).toBe(108);
    b.set('A1', 5, 'Sheet1');
    expect(b.v('D1')).toBe(112);
  });

  it('is evaluated on every pass, being recalculated whatever changes', () => {
    const b = threeSheets();
    b.f('D1', '=SUM(Sheet1:Sheet3!A1)');
    const key = (b.sheet().getRowId(0) as string) + ':' + (b.sheet().getColId(3) as string);
    expect(b.wb.getFormulaGraph().nodes.get(key)?.volatile).toBe(true);
    expect(b.wb.getFormulaGraph().nodes.get(key)?.dependencies.size).toBe(0);
    expect(b.wb.getFormulaGraph().nodes.get(key)?.rangeRects).toHaveLength(0);
  });

  it('is a positional reference: structural edits on its sheets do not move it', () => {
    const b = threeSheets();
    b.f('D1', '=SUM(Sheet1:Sheet3!A1)');
    b.sheet('Sheet2').insertRows(0, 1);
    expect(b.v('D1')).toBe(1 + 0 + 3); // A1 of Sheet2 is now empty (its value moved to A2)
    expect(b.text('D1')).toBe('=SUM(Sheet1:Sheet3!A1)');
  });

  it('follows a renamed end sheet and quotes the span when needed', () => {
    const b = threeSheets();
    b.f('D1', '=SUM(Sheet1:Sheet3!A1)');
    b.wb.renameSheet(b.id('Sheet3') as string, 'Last');
    expect(b.text('D1')).toBe('=SUM(Sheet1:Last!A1)');
    expect(b.v('D1')).toBe(6);
    b.wb.renameSheet(b.id('Sheet1') as string, 'First one');
    expect(b.text('D1', 'First one')).toBe("=SUM('First one:Last'!A1)");
    expect(b.v('D1', 'First one')).toBe(6);
    b.wb.renameSheet(b.id('Sheet2') as string, 'Mid'); // an inside sheet: no change to the span
    expect(b.text('D1', 'First one')).toBe("=SUM('First one:Last'!A1)");
    expect(b.v('D1', 'First one')).toBe(6);
  });

  it('shrinks inward when an end sheet is deleted', () => {
    const b = threeSheets();
    b.f('D1', '=SUM(Sheet1:Sheet3!A1)', 'Sheet2');
    b.wb.deleteSheet(b.id('Sheet3') as string);
    expect(b.text('D1', 'Sheet2')).toBe('=SUM(Sheet1:Sheet2!A1)');
    expect(b.v('D1', 'Sheet2')).toBe(3);

    const c = threeSheets();
    c.f('D1', '=SUM(Sheet1:Sheet3!A1)', 'Sheet2');
    c.wb.deleteSheet(c.id('Sheet1') as string);
    expect(c.text('D1', 'Sheet2')).toBe('=SUM(Sheet2:Sheet3!A1)');
    expect(c.v('D1', 'Sheet2')).toBe(5);
  });

  it('shrinks inward for a reversed span too', () => {
    const b = threeSheets();
    b.f('D1', '=SUM(Sheet3:Sheet1!A1)', 'Sheet2');
    b.wb.deleteSheet(b.id('Sheet3') as string);
    expect(b.text('D1', 'Sheet2')).toBe('=SUM(Sheet2:Sheet1!A1)');
    expect(b.v('D1', 'Sheet2')).toBe(3);
  });

  it('drops a middle sheet from the total when it is deleted', () => {
    const b = threeSheets();
    b.f('D1', '=SUM(Sheet1:Sheet3!A1)');
    b.wb.deleteSheet(b.id('Sheet2') as string);
    expect(b.text('D1')).toBe('=SUM(Sheet1:Sheet3!A1)');
    expect(b.v('D1')).toBe(4);
  });

  it('is #REF! when the only sheet of the span is deleted', () => {
    const b = new Book(['Sheet1', 'Sheet2']);
    b.f('A1', '=SUM(Sheet2:Sheet2!A1)');
    b.wb.deleteSheet(b.id('Sheet2') as string);
    expect(b.shown('A1')).toBe('#REF!');
  });

  it('covers a new sheet only if it lands between the end sheets', () => {
    const b = threeSheets();
    b.f('D1', '=SUM(Sheet1:Sheet3!A1)');
    b.wb.addSheet('Sheet4'); // appended after the span
    b.set('A1', 1000, 'Sheet4');
    expect(b.v('D1')).toBe(6);
  });

  it('follows the workbook sheet order when sheets are reordered', () => {
    const b = threeSheets();
    b.f('D1', '=SUM(Sheet1:Sheet2!A1)');
    expect(b.v('D1')).toBe(3);
    const data = JSON.parse(JSON.stringify(b.wb.getData()));
    data.sheets = [data.sheets[0], data.sheets[2], data.sheets[1]]; // Sheet1, Sheet3, Sheet2
    b.wb.setData(data);
    expect(b.v('D1')).toBe(1 + 3 + 2); // Sheet3 now sits between Sheet1 and Sheet2
    expect(b.text('D1')).toBe('=SUM(Sheet1:Sheet2!A1)');
  });

  it('is #VALUE! when used outside a function argument', () => {
    const b = threeSheets();
    b.f('D1', '=Sheet1:Sheet3!A1');
    expect(b.shown('D1')).toBe('#VALUE!');
  });

  it('is #REF! when an end sheet is unknown', () => {
    const b = threeSheets();
    b.f('D1', '=SUM(Sheet1:Nowhere!A1)');
    expect(b.shown('D1')).toBe('#REF!');
    b.f('D2', '=SUM(Nowhere:Sheet3!A1)');
    expect(b.shown('D2')).toBe('#REF!');
  });

  it('binds again when the missing end sheet is created', () => {
    const b = new Book(['Sheet1', 'Sheet2']);
    b.set('A1', 1, 'Sheet1').set('A1', 2, 'Sheet2');
    b.f('D1', '=SUM(Sheet1:Later!A1)');
    expect(b.shown('D1')).toBe('#REF!');
    b.wb.addSheet('Later');
    b.set('A1', 4, 'Later');
    expect(b.v('D1')).toBe(7);
  });

  it('reads whole columns and rows across the span', () => {
    const b = threeSheets();
    b.f('D10', '=SUM(Sheet1:Sheet3!A:A)').f('D11', '=SUM(Sheet1:Sheet3!2:2)');
    expect([b.v('D10'), b.v('D11')]).toEqual([6, 60]);
    expect(b.text('D10')).toBe('=SUM(Sheet1:Sheet3!A:A)');
    expect(b.text('D11')).toBe('=SUM(Sheet1:Sheet3!2:2)');
  });
});

describe('ranges built from reference-returning functions in a workbook', () => {
  it('ends a range at INDEX, OFFSET and INDIRECT and tracks changes', () => {
    const b = new Book().fill('A', [1, 2, 3, 4, 5]);
    b.f('C1', '=SUM(A1:INDEX(A1:A5,3))').f('C2', '=SUM(A1:OFFSET(A1,2,0))').f('C3', '=SUM(INDIRECT("A1"):INDIRECT("A3"))');
    expect([b.v('C1'), b.v('C2'), b.v('C3')]).toEqual([6, 6, 6]);
    b.set('A2', 20);
    expect([b.v('C1'), b.v('C2'), b.v('C3')]).toEqual([24, 24, 24]);
    expect(['C1', 'C2', 'C3'].map((a) => b.text(a))).toEqual([
      '=SUM(A1:INDEX(A1:A5,3))',
      '=SUM(A1:OFFSET(A1,2,0))',
      '=SUM(INDIRECT("A1"):INDIRECT("A3"))',
    ]);
  });

  it('recomputes a range between computed references when a cell inside it changes', () => {
    const b = new Book().fill('A', [1, 2, 3, 4, 5]).fill('B', [10, 20, 30, 40, 50]);
    b.f('D1', '=SUM(A1:INDEX(B1:B5,2))');
    expect(b.v('D1')).toBe(1 + 2 + 10 + 20);
    b.set('B2', 200); // the box is A1:B2, which no single listed reference covers
    expect(b.v('D1')).toBe(1 + 2 + 10 + 200);
    b.set('A2', 7);
    expect(b.v('D1')).toBe(1 + 7 + 10 + 200);
  });

  it('feeds ROW, COLUMN and the shape functions', () => {
    const b = new Book().fill('A', [1, 2, 3]);
    b.f('C1', '=ROW(1:1)').f('C2', '=COLUMN(A:A)').f('C3', '=ROWS(A:A)').f('C4', '=COLUMNS(1:1)');
    expect([b.v('C1'), b.v('C2'), b.v('C3'), b.v('C4')]).toEqual([1, 1, 1048576, 16384]);
  });

  it('does not make ROW(1:1) depend on row 1 (no false cycle)', () => {
    const b = new Book();
    b.f('A1', '=ROW(1:1)').f('B1', '=COLUMN(A:A)+ROWS(1:1)');
    expect([b.shown('A1'), b.shown('B1')]).toEqual([1, 2]);
  });
});

describe('union and intersection in a workbook', () => {
  it('sums a union and an intersection and tracks edits', () => {
    const b = new Book().fill('A', [1, 2, 3, 4, 5]).set('C1', 100);
    b.f('E1', '=SUM((A1:A2,A4,C1))').f('E2', '=SUM(A1:A3 A2:A4)').f('E10', '=SUM(A:A 3:3)');
    expect([b.v('E1'), b.v('E2'), b.v('E10')]).toEqual([107, 5, 3]);
    b.set('A2', 20);
    b.set('A3', 30);
    expect([b.v('E1'), b.v('E2'), b.v('E10')]).toEqual([125, 50, 30]);
    expect(['E1', 'E2', 'E10'].map((a) => b.text(a))).toEqual(['=SUM((A1:A2,A4,C1))', '=SUM(A1:A3 A2:A4)', '=SUM(A:A 3:3)']);
  });

  it('is #NULL! for an empty intersection and #VALUE! for a union as a result', () => {
    const b = new Book().fill('A', [1, 2]);
    b.f('C1', '=A1:A2 B1:B2').f('C2', '=(A1,A2)');
    expect(b.shown('C1')).toBe('#NULL!');
    expect(b.shown('C2')).toBe('#VALUE!');
  });

  it('follows rows inserted inside an intersected range', () => {
    const b = new Book().fill('A', [1, 2, 3]);
    b.f('C1', '=SUM(A1:A3 A2:A3)');
    b.sheet().insertRows(2, 1);
    expect(b.text('C1')).toBe('=SUM(A1:A4 A2:A4)');
    expect(b.v('C1')).toBe(5);
  });
});

describe('implicit intersection in a workbook', () => {
  it('reads the cell in the formula row and tracks edits', () => {
    const b = new Book().fill('A', [1, 2, 3]);
    b.f('C2', '=@A1:A3').f('C3', '=@A:A*10').f('C9', '=@A1:A3');
    expect(b.v('C2')).toBe(2);
    expect(b.v('C3')).toBe(30);
    expect(b.shown('C9')).toBe('#VALUE!');
    b.set('A2', 50);
    expect(b.v('C2')).toBe(50);
    expect(b.text('C2')).toBe('=@A1:A3');
  });
});

describe('the spill operator in a workbook', () => {
  it('reads the spill range of an anchor and follows its size', () => {
    const b = new Book().set('E1', 3);
    b.f('A1', '=SEQUENCE(E1)').f('G1', '=SUM(A1#)').f('G2', '=ROWS(A1#)').f('G3', '=A1#');
    expect([b.v('G1'), b.v('G2')]).toEqual([6, 3]);
    expect([b.v('G3'), b.v('G4'), b.v('G5')]).toEqual([1, 2, 3]);
    b.set('E1', 5);
    expect([b.v('G1'), b.v('G2')]).toEqual([15, 5]);
    b.set('E1', 2);
    expect([b.v('G1'), b.v('G2')]).toEqual([3, 2]);
    expect(b.text('G1')).toBe('=SUM(A1#)');
  });

  it('is #REF! for a cell that anchors no spill, and recovers when it starts to', () => {
    const b = new Book().set('A1', 1);
    b.f('C1', '=SUM(A1#)');
    expect(b.shown('C1')).toBe('#REF!');
    b.f('A1', '=SEQUENCE(3)');
    expect(b.v('C1')).toBe(6);
    b.set('A1', 9);
    expect(b.shown('C1')).toBe('#REF!');
  });

  it('is #REF! for a spilled cell that is not the anchor', () => {
    const b = new Book().f('A1', '=SEQUENCE(3)');
    b.f('C1', '=SUM(A2#)');
    expect(b.shown('C1')).toBe('#REF!');
  });

  it('reads an anchor on another sheet', () => {
    const b = new Book().f('B2', '=SEQUENCE(4)', 'Sheet2');
    b.f('A1', '=SUM(Sheet2!B2#)');
    expect(b.v('A1')).toBe(10);
    expect(b.text('A1')).toBe('=SUM(Sheet2!B2#)');
    b.f('A2', '=SUM(Nowhere!B2#)');
    expect(b.shown('A2')).toBe('#REF!');
  });

  it('reads the anchor before it has been evaluated during a full recalculation', () => {
    const b = new Book().f('B1', '=SUM(D1#)').f('D1', '=SEQUENCE(3)');
    // Renaming a sheet re-evaluates every formula in sheet order (B1 before D1).
    b.wb.renameSheet(b.id() as string, 'Renamed');
    expect(b.v('B1', 'Renamed')).toBe(6);
  });

  it('works as a range endpoint and in whole-column sums', () => {
    const b = new Book().f('A1', '=SEQUENCE(3)');
    b.f('C1', '=SUM(A1#:B1)').f('C2', '=SUM(A:A)');
    expect(b.v('C1')).toBe(6);
    expect(b.v('C2')).toBe(6);
  });
});

describe('error literals in a workbook', () => {
  it('shows and propagates #NULL!', () => {
    const b = new Book();
    b.f('A1', '=#NULL!').f('A2', '=ISERROR(A1)').f('A3', '=IFERROR(#NULL!,"x")').f('A4', '=SUM(1,#NULL!)');
    expect(b.shown('A1')).toBe('#NULL!');
    expect(b.v('A2')).toBe(true);
    expect(b.v('A3')).toBe('x');
    expect(b.shown('A4')).toBe('#NULL!');
    expect(b.text('A1')).toBe('=#NULL!');
  });
});

describe('formulas that do not parse', () => {
  it('show an #ERROR! and stop being the previous formula', () => {
    const b = new Book().fill('A', [1, 2]);
    b.f('C1', '=SUM(A1:A2)');
    expect(b.v('C1')).toBe(3);
    b.f('C1', '=SUM(A1:A2');
    expect(b.shown('C1')).toBe('#ERROR! Unexpected end of formula (expected "," or ")")');
    expect(b.text('C1')).toBe('=SUM(A1:A2');
    b.set('A1', 10); // a recalculation must not bring the old formula back
    expect(b.shown('C1')).toContain('#ERROR!');
    expect(b.text('C1')).toBe('=SUM(A1:A2');
  });

  it('release a spill the cell anchored', () => {
    const b = new Book().f('A1', '=SEQUENCE(3)');
    expect(b.shown('A2')).toBe(undefined);
    expect(b.wb.isSpilledCell(undefined, 1, 0)).toBe(true);
    b.f('A1', '=SEQUENCE(');
    expect(b.wb.isSpilledCell(undefined, 1, 0)).toBe(false);
  });

  it('reach dependents', () => {
    const b = new Book().set('A1', 1);
    b.f('B1', '=A1+1').f('C1', '=B1*2');
    expect(b.v('C1')).toBe(4);
    b.f('B1', '=A1+');
    expect(b.shown('B1')).toContain('#ERROR!');
    expect(b.shown('C1')).not.toBe(4);
  });

  it('are reported for the malformed forms of the new syntax', () => {
    const b = new Book();
    for (const bad of ['=A:', '=1:A', '=SUM({1,2;3})', '=Sheet1:!A1', '=SUM({})', '=(A1,2)', '=A1:', '=@']) {
      b.f('A1', bad);
      expect(b.shown('A1')).toMatch(/^#ERROR! /);
    }
    b.f('A1', '=SUM({1,2;3})');
    expect(b.shown('A1')).toBe('#ERROR! Array constant at position 4 is ragged: row 2 has 1 elements, row 1 has 2');
  });

  it('keep the file value of an imported formula the engine cannot parse until it is recalculated', () => {
    const b = new Book();
    const data = b.wb.getData();
    delete data.sheets[0].rowOrder; // the legacy wire form: cells keyed "row:col"
    delete data.sheets[0].colOrder;
    data.sheets[0].cells = [{ key: '0:0', cell: { value: 23, formula: '=SUM(Sales[Q1])' } }] as never;
    b.wb.setData(data);
    expect(b.wb.getCell(undefined, 0, 0)?.value).toBe(23);
    expect(b.wb.getCell(undefined, 0, 0)?.formula).toBe('=SUM(Sales[Q1])');
  });
});

describe('formula text round trips through the engine', () => {
  const formulas = [
    '=SUM(A:A)', '=SUM($A:$A)', '=SUM(A:C)', '=SUM(1:1)', '=SUM($1:$3)', '=SUM(Sheet2!A:B)', '=SUM(Sheet2!2:3)',
    '=A1+50%', '=10%^2', '=-A1%', '=(A1+1)%', '=5%%', '=-(10%)', '=1+2*3%', '=A1^B1%',
    '={1,2,3}', '={1;2;3}', '=SUM({1,2;3,4})', '={-1,"a""b",TRUE,FALSE,#N/A}', '={1.5,2.25}',
    '=SUM(Sheet1:Sheet3!A1)', '=SUM(Sheet1:Sheet3!A1:B2)', "=SUM('My Sheet:Other Sheet'!A1)", '=SUM(Sheet1:Sheet3!A:A)', '=SUM(Sheet1:Sheet3!1:2)',
    '=SUM(A1:INDEX(A1:A5,3))', '=INDEX(A1:A5,1):INDEX(A1:A5,3)', '=SUM(A1:OFFSET(A1,2,0))', '=SUM(INDIRECT("A1"):INDIRECT("A3"))',
    '=SUM(A1:A3 A2:A4)', '=SUM((A1:A2,C1))', '=SUM((A1,B1,C1))', '=SUM((A1:B2 B1:C3,D1))', '=SUM((A1:A2,C1),5)',
    '=@A1:A3', '=@A1', '=@INDEX(A1:A3,1)+1', '=SUM(A1#)', '=SUM(Sheet2!B2#)', '=A1#*2',
    '=#NULL!', '=#N/A', '=IFERROR(#DIV/0!,0)',
    '=SUM(A:A)/COUNT(A:A)', '=IF(A1>0,SUM(1:1),0)', '=A1&"x"&B1', '=SUM(A1,B1)', '=1-(2-3)', '=-(1+2)^2', '=2^(3^2)',
  ];

  it.each(formulas)('%s prints back unchanged', (formula) => {
    const b = new Book(['Sheet1', 'Sheet2', 'Sheet3', 'My Sheet', 'Other Sheet']);
    b.f('H9', formula);
    expect(b.text('H9')).toBe(formula);
  });

  it.each([
    ['=sum(a:a)', '=SUM(A:A)'],
    ['= SUM( A1 : B2 )', '=SUM(A1:B2)'],
    ['=a1+ 50 %', '=A1+50%'],
    ['={ 1 , 2 ; 3 , 4 }', '={1,2;3,4}'],
    ['=SUM(Sheet1:Sheet3!$a$1)', '=SUM(Sheet1:Sheet3!$A$1)'],
    ['=SUM((A1:A2 , C1))', '=SUM((A1:A2,C1))'],
    ['=#n/a', '=#N/A'],
    ["=SUM('Sheet1:Sheet3'!A1)", '=SUM(Sheet1:Sheet3!A1)'],
    ['=SUM(((A1,B1),C1))', '=SUM((A1,B1,C1))'],
    ['=SUM(A1:A3  A2:A4)', '=SUM(A1:A3 A2:A4)'],
    ['=1.50+2.0', '=1.5+2'],
  ])('%s is normalised to %s', (typed, shown) => {
    const b = new Book();
    b.f('H9', typed);
    expect(b.text('H9')).toBe(shown);
  });

  it('quotes sheet names that need it', () => {
    const b = new Book(['Sheet1', 'My Sheet', 'A1', 'it\'s', 'Q-1', '2024', 'TRUE', 'Über']);
    b.f('H9', "=SUM('My Sheet'!A1)+SUM('A1'!B2)+SUM('it''s'!C3)+SUM('Q-1'!D4)+SUM('2024'!E5)+SUM('TRUE'!F6)+SUM(Über!G7)");
    expect(b.text('H9')).toBe("=SUM('My Sheet'!A1)+SUM('A1'!B2)+SUM('it''s'!C3)+SUM('Q-1'!D4)+SUM('2024'!E5)+SUM('TRUE'!F6)+SUM(Über!G7)");
  });
});

describe('SUM(A:A) on a column of a million rows', () => {
  it('is fast, and recalculates on an edit', () => {
    const b = new Book();
    const sheet = b.sheet();
    const rows = 1_000_000;
    let expected = 0;
    for (let i = 0; i < rows; i++) {
      sheet.setCell(i, 0, { value: i % 7 });
      expected += i % 7;
    }
    const t0 = Date.now();
    b.f('B1', '=SUM(A:A)');
    const first = Date.now() - t0;
    expect(b.v('B1')).toBe(expected);
    const t1 = Date.now();
    b.set('A500000', 1000);
    const edit = Date.now() - t1;
    expect(b.v('B1')).toBe(expected - (499999 % 7) + 1000);
    // Measured at about 2 s each on a laptop; the bound is generous so a slow CI runner passes
    // while a per-row or per-cell dependency expansion (minutes) would not.
    expect(first).toBeLessThan(20_000);
    expect(edit).toBeLessThan(20_000);
    const node = b.wb.getFormulaGraph().nodes.get(
      (sheet.getRowId(0) as string) + ':' + (sheet.getColId(1) as string)
    );
    expect(node?.dependencies.size).toBe(0);
    expect(node?.rangeRects).toHaveLength(1);
  }, 120_000);
});

describe('dependency graph resolver for whole-axis keys', () => {
  it('resolves the bounded axis of a whole-column and whole-row key', () => {
    const b = new Book();
    const resolve = (key: string): ReturnType<CellResolver> =>
      (b.wb as unknown as { resolveGlobalCell: CellResolver }).resolveGlobalCell(key);
    b.set('C3', 1);
    const colId = b.sheet().getColId(2) as string;
    const rowId = b.sheet().getRowId(2) as string;
    expect(resolve(`*:${colId}`)).toMatchObject({ row: 0, col: 2 });
    expect(resolve(`${rowId}:*`)).toMatchObject({ row: 2, col: 0 });
    expect(resolve('*:NOSUCHCOL')).toBeUndefined();
    expect(resolve('NOSUCHROW:*')).toBeUndefined();
  });
});

describe('collaboration: the new AST kinds travel as plain JSON', () => {
  it('a peer receives whole-column, 3D, array, percent, @ and # formulas and computes the same values', async () => {
    const { InMemoryProvider, __resetInMemoryRooms } = await import('@b32nio/spindle-shared');
    __resetInMemoryRooms();
    const identity = (name: string) => ({ userId: name, displayName: name, color: '#123456' });
    const a = new Book(['Sheet1']);
    a.fill('A', [1, 2, 3]);
    const ha = await a.wb.attachCollab(new InMemoryProvider(), identity('A'), { roomId: 'room' });
    const b = new Book(['Sheet1']);
    const hb = await b.wb.attachCollab(new InMemoryProvider(), identity('B'), { roomId: 'room' });
    a.wb.addSheet('Sheet2');
    a.set('A1', 10, 'Sheet2');

    try {
      const formulas: Array<[string, string, number]> = [
        ['D1', '=SUM(A:A)', 6],
        ['D2', '=SUM(Sheet1:Sheet2!A1)', 11],
        ['D3', '=SUM({1,2;3,4})+A1%', 10.01],
        ['H2', '=@A1:A3', 2],
        ['D5', '=SUM(A1:INDEX(A1:A3,2))', 3],
        ['D6', '=SUM((A1,A3))', 4],
      ];
      for (const [addr, formula] of formulas) a.f(addr, formula);
      a.f('F1', '=SEQUENCE(3)').f('G1', '=SUM(F1#)');

      for (const [addr, formula, value] of formulas) {
        expect(b.text(addr)).toBe(formula);
        expect([addr, b.v(addr)]).toEqual([addr, expect.closeTo(value)]);
      }
      expect(b.text('G1')).toBe('=SUM(F1#)');
      expect(b.v('G1')).toBe(6);
      // The peer edits the data; the formula's owner sees the whole-column sum change.
      b.set('A9', 100);
      expect(a.v('D1')).toBe(106);

    } finally {
      ha.detach();
      hb.detach();
      __resetInMemoryRooms();
    }
  });
});

describe('whole ranges on one sheet line up (review fixes)', () => {
  it('gives two whole columns of different lengths the same height', () => {
    const b = new Book().fill('A', [1, 2, 3]).fill('B', [10, 20, 30, 40, 50]);
    b.f('D1', '=SUMPRODUCT(A:A*B:B)').f('D2', '=SUM(A:A+B:B)').f('D3', '=SUMPRODUCT(A:A,B:B)').f('D4', '=SUMPRODUCT((A:A>1)*(B:B))');
    expect(b.v('D1')).toBe(10 + 40 + 90);
    expect(b.v('D2')).toBe(6 + 150);
    expect(b.v('D3')).toBe(140);
    expect(b.v('D4')).toBe(20 + 30 + 0 + 0 + 0 + 0 + 0); // rows 4-5 have no A: blank > 1 is false
  });

  it('gives two whole rows of different widths the same width', () => {
    const b = new Book().fill('A', [1, 2]).set('B1', 2).set('C1', 3).set('B2', 5);
    b.f('A10', '=SUMPRODUCT(1:1*2:2)');
    expect(b.v('A10')).toBe(1 * 2 + 2 * 5 + 0);
  });

  it('trims a very wide or very tall span to the used range instead of allocating it', () => {
    const b = new Book(['Sheet1', 'Sheet2']);
    for (let r = 0; r < 200; r++) b.wb.setCellValue(undefined, r, r % 5, 1);
    b.f('A1', '=SUM(Sheet1!A:XFD)', 'Sheet2').f('A2', '=COUNTA(Sheet1!1:1048576)', 'Sheet2');
    b.f('A3', '=COLUMNS(Sheet1!A:XFD)', 'Sheet2').f('A4', '=ROWS(Sheet1!1:1048576)', 'Sheet2');
    const t = Date.now();
    expect([b.v('A1', 'Sheet2'), b.v('A2', 'Sheet2')]).toEqual([200, 200]);
    expect(Date.now() - t).toBeLessThan(5000);
    expect([b.v('A3', 'Sheet2'), b.v('A4', 'Sheet2')]).toEqual([16384, 1048576]);
    // An empty sheet still has one cell.
    const empty = new Book(['Sheet1', 'Sheet2']);
    empty.f('A1', '=COUNTA(Sheet1!B:XFD)', 'Sheet2').f('A2', '=COUNTA(Sheet1!5:1048576)', 'Sheet2');
    expect([empty.v('A1', 'Sheet2'), empty.v('A2', 'Sheet2')]).toEqual([0, 0]);
  });

  it('keeps a modest span at its nominal width, so VLOOKUP into a sparse A:C still works', () => {
    const b = new Book().set('A1', 'k').set('A2', 'j');
    b.f('E1', '=IFERROR(VLOOKUP("k",A:C,3,FALSE),"none")');
    expect(b.v('E1')).toBe(0); // the blank C1: column 3 exists
  });
});

describe('formulas the engine cannot parse keep their saved value (review fixes)', () => {
  function importedBook(): Book {
    const b = new Book(['Sheet1']);
    const data = b.wb.getData();
    delete data.sheets[0].rowOrder;
    delete data.sheets[0].colOrder;
    data.sheets[0].cells = [
      { key: '0:0', cell: { value: 23, formula: '=SUM(Sales[Q1])' } },
      { key: '1:0', cell: { value: 5 } },
    ] as never;
    b.wb.setData(data);
    return b;
  }

  it('keeps the value and the text through a full recalculation and an unrelated edit', () => {
    const b = importedBook();
    b.set('A2', 6);
    b.wb.renameSheet(b.id() as string, 'Renamed');
    expect(b.shown('A1', 'Renamed')).toBe(23);
    expect(b.v('A1', 'Renamed')).toBe(23);
    expect(b.text('A1', 'Renamed')).toBe('=SUM(Sales[Q1])');
    b.f('B1', '=A1+1', 'Renamed');
    expect(b.v('B1', 'Renamed')).toBe(24);
  });

  it('writes the saved value back on export, never an error string (an error result is an error value)', async () => {
    const b = importedBook();
    b.set('A2', 6);
    b.f('C1', '=SUM(');
    b.f('C2', '=Undefined1');
    const { exportToXlsx } = await import('./export/xlsx');
    const ExcelJS = (await import('exceljs')).default;
    const bytes = await exportToXlsx(b.wb);
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(bytes as unknown as ArrayBuffer);
    const ws = wb.worksheets[0];
    expect((ws.getCell('A1').value as { result?: unknown }).result).toBe(23);
    expect(JSON.stringify(ws.getCell('C1').value)).not.toContain('ERROR');
    // an error result is an error value (t="e"), read back by exceljs as { error }
    expect((ws.getCell('C2').value as { result?: unknown }).result).toEqual({ error: '#NAME?' });
  });
});

describe('array-constant errors are tagged (review fixes)', () => {
  it('keeps {#N/A} an error and {"#N/A"} text through printing', () => {
    const b = new Book();
    b.f('A1', '={#N/A,"#N/A","x"}').f('A2', '={"#N/A"}');
    expect(b.text('A1')).toBe('={#N/A,"#N/A","x"}');
    expect(b.text('A2')).toBe('={"#N/A"}');
    expect(b.wb.getCell(undefined, 0, 0)?.formulaAst?.value).toEqual([[{ error: '#N/A' }, '#N/A', 'x']]);
    expect([b.v('A1'), b.v('B1'), b.v('C1')]).toEqual(['#N/A', '#N/A', 'x']);
  });
});

describe('unquoted sheet names that earlier versions accepted', () => {
  it('still read, and print back quoted', () => {
    const b = new Book(['Sheet1', 'My Sheet', 'Q1 Data', '2024']);
    b.set('A1', 7, 'My Sheet').set('A1', 9, 'Q1 Data').set('A1', 8, '2024');
    b.f('C1', '=My Sheet!A1+Q1 Data!A1+2024!A1');
    expect(b.v('C1')).toBe(24);
    expect(b.text('C1')).toBe("='My Sheet'!A1+'Q1 Data'!A1+'2024'!A1");
  });
});
