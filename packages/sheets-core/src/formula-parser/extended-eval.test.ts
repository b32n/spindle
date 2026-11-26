// Evaluation of the reference forms and operators added to the parser: whole-column/row
// ranges, postfix percent, array constants, 3D references, reference-returning functions
// as range endpoints, union, intersection, `@`, the spill operator and error literals.
// A mock sheet store stands in for the workbook; the workbook-level behaviour (dependencies,
// structural edits, spills) is in extended-references.test.ts.

import { FormulaParser } from './parser';
import type { EvaluationContext, RangeReference } from './types';

type Grid = Record<string, unknown>; // "row:col" -> value (0-based)

/** Build a grid from A1-style addresses: { A1: 1, B2: 'x' }. */
function grid(cells: Record<string, unknown>): Grid {
  const out: Grid = {};
  for (const [addr, value] of Object.entries(cells)) {
    const m = /^([A-Z]+)(\d+)$/.exec(addr) as RegExpExecArray;
    let col = 0;
    for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
    out[`${Number(m[2]) - 1}:${col - 1}`] = value;
  }
  return out;
}

const parser = new FormulaParser();

interface Setup {
  sheets?: Record<string, Grid>;
  spills?: Record<string, { rows: number; cols: number; values: unknown[][] }>; // "sheet|row:col"
  at?: [number, number];
}

function run(formula: string, setup: Setup = {}): unknown {
  const sheets = setup.sheets ?? { S1: {} };
  const names = Object.keys(sheets);
  const [row, col] = setup.at ?? [0, 0];
  const value = (r: number, c: number, sheet?: string): unknown => {
    const g = sheets[sheet ?? names[0]];
    if (!g) return '#REF!';
    return g[`${r}:${c}`] ?? null;
  };
  const abs = (ref: { row: number; col: number; rowAbsolute: boolean; colAbsolute: boolean }) => ({
    row: ref.rowAbsolute ? ref.row : row + ref.row,
    col: ref.colAbsolute ? ref.col : col + ref.col,
  });
  const ctx: EvaluationContext = {
    getCellValue: (r, c, _id, sheet) => value(r, c, sheet),
    getRangeValues: (range: RangeReference, _id, sheet) => {
      const [a, b] = [abs(range.start), abs(range.end)];
      const [r1, r2] = [Math.min(a.row, b.row), Math.max(a.row, b.row)];
      const [c1, c2] = [Math.min(a.col, b.col), Math.max(a.col, b.col)];
      const g = sheets[sheet ?? names[0]] ?? {};
      let [lastRow, lastCol] = [r2, c2];
      if (range.whole) {
        // The workbook clamps the unbounded axis to the used range; so does the mock.
        const used = Object.keys(g).map((k) => k.split(':').map(Number));
        if (range.whole === 'column') lastRow = Math.max(r1, ...used.filter(([, c]) => c >= c1 && c <= c2).map(([r]) => r));
        else lastCol = Math.max(c1, ...used.filter(([r]) => r >= r1 && r <= r2).map(([, c]) => c));
      }
      const out: unknown[][] = [];
      for (let r = r1; r <= lastRow; r++) {
        const rowValues: unknown[] = [];
        for (let c = c1; c <= lastCol; c++) rowValues.push(value(r, c, sheet));
        out.push(rowValues);
      }
      return out;
    },
    getSheetNames: () => names,
    getSpillRegion: (r, c, sheet) => setup.spills?.[`${sheet ?? names[0]}|${r}:${c}`],
  };
  const { ast, error } = parser.parse(formula, row, col);
  if (error) throw new Error(error);
  return parser.evaluate(ast, ctx, row, col);
}

const data = grid({ A1: 1, A2: 2, A3: 3, A4: 4, A5: 5, B1: 10, B2: 20, C1: 100, J1: 7 });

describe('whole-column and whole-row ranges', () => {
  const sheets = { S1: data };

  it('sums a column and a row, clamped to the used range', () => {
    expect(run('=SUM(A:A)', { sheets, at: [0, 5] })).toBe(15);
    expect(run('=SUM(1:1)', { sheets, at: [5, 0] })).toBe(1 + 10 + 100 + 7);
    expect(run('=SUM(A:C)', { sheets, at: [0, 5] })).toBe(15 + 30 + 100);
    expect(run('=SUM(1:2)', { sheets, at: [9, 0] })).toBe(1 + 10 + 100 + 7 + 2 + 20);
  });

  it('is empty (not an error) over an empty column', () => {
    expect(run('=SUM(Z:Z)', { sheets })).toBe(0);
    expect(run('=COUNT(Z:Z)', { sheets })).toBe(0);
    expect(run('=SUM(50:50)', { sheets })).toBe(0);
  });

  it('keeps the nominal size for ROWS, COLUMNS and COUNTBLANK', () => {
    expect(run('=ROWS(A:A)', { sheets })).toBe(1048576);
    expect(run('=COLUMNS(A:C)', { sheets })).toBe(3);
    expect(run('=ROWS(1:3)', { sheets })).toBe(3);
    expect(run('=COLUMNS(1:1)', { sheets })).toBe(16384);
    expect(run('=COUNTBLANK(A:A)', { sheets, at: [0, 5] })).toBe(1048576 - 5);
  });

  it('reports the first row or column of the range for ROW and COLUMN', () => {
    expect(run('=ROW(1:1)', { sheets })).toBe(1);
    expect(run('=ROW(3:5)', { sheets })).toBe(3);
    expect(run('=COLUMN(A:A)', { sheets })).toBe(1);
    expect(run('=COLUMN(C:D)', { sheets })).toBe(3);
    expect(run('=ROW(B:B)', { sheets })).toBe(1);
    expect(run('=COLUMN(2:2)', { sheets })).toBe(1);
  });

  it('reads relative whole ranges relative to the formula cell', () => {
    // From G1, column offset -6 reaches A.
    expect(run('=SUM(A:A)', { sheets, at: [0, 6] })).toBe(15);
    // $A:$A is absolute.
    expect(run('=SUM($A:$A)', { sheets, at: [3, 9] })).toBe(15);
    expect(run('=SUM(2:2)', { sheets, at: [9, 9] })).toBe(2 + 20);
  });

  it('reads another sheet', () => {
    const two = { S1: data, Other: grid({ A1: 100, A2: 200 }) };
    expect(run('=SUM(Other!A:A)', { sheets: two })).toBe(300);
    expect(run('=SUM(Other!1:1)', { sheets: two })).toBe(100);
  });

  it('indexes the nominal column', () => {
    expect(run('=INDEX(A:A,3)', { sheets })).toBe(3);
    expect(run('=INDEX(A:A,500)', { sheets })).toBeNull(); // past the used range, still inside the sheet
    expect(run('=INDEX(A:C,2,2)', { sheets })).toBe(20);
    expect(run('=INDEX(1:1,2)', { sheets })).toBe(10);
    expect(() => run('=INDEX(A:A,1048577)', { sheets })).toThrow('#REF!');
  });

  it('works with MATCH, VLOOKUP and array arithmetic over the used range', () => {
    expect(run('=MATCH(3,A:A,0)', { sheets, at: [0, 5] })).toBe(3);
    expect(run('=VLOOKUP(2,A:B,2,FALSE)', { sheets, at: [0, 5] })).toBe(20);
    expect(run('=SUMPRODUCT(A:A*2)', { sheets, at: [0, 5] })).toBe(30);
  });
});

describe('percent', () => {
  it('divides by 100 as a postfix operator', () => {
    expect(run('=50%')).toBe(0.5);
    expect(run('=A1+50%', { sheets: { S1: data } })).toBe(1.5);
    expect(run('=200*10%')).toBeCloseTo(20);
    expect(run('=5%%')).toBeCloseTo(0.0005);
    expect(run('=-50%')).toBe(-0.5);
    expect(run('=(1+1)%')).toBe(0.02);
    expect(run('=1.5e2%')).toBe(1.5);
  });

  it('binds tighter than ^ and looser than a leading sign', () => {
    expect(run('=10%^2')).toBeCloseTo(0.01);
    expect(run('=2^50%')).toBeCloseTo(Math.SQRT2);
    expect(run('=-2^2')).toBe(4);
    expect(run('=-10%^2')).toBeCloseTo(0.01);
  });

  it('applies element-wise to arrays and to ranges', () => {
    expect(run('={50,100}%')).toEqual([[0.5, 1]]);
    expect(run('=A1:A2%', { sheets: { S1: data } })).toEqual([[0.01], [0.02]]);
  });

  it('reads a blank as 0 and a boolean as 1 or 0', () => {
    expect(run('=Z9%')).toBe(0);
    expect(run('=TRUE%')).toBe(0.01);
  });

  it('is #VALUE! over text and keeps an error element an error in an array', () => {
    expect(() => run('="abc"%')).toThrow('#VALUE!');
    expect(run('={50,"x"}%')).toEqual([[0.5, '#VALUE!']]);
  });
});

describe('array constants', () => {
  it('evaluates to a fresh matrix every time', () => {
    const { ast } = parser.parse('={1,2;3,4}');
    const ctx = { getCellValue: () => null, getRangeValues: () => [] };
    const first = parser.evaluate(ast, ctx) as number[][];
    expect(first).toEqual([[1, 2], [3, 4]]);
    first[0][0] = 99;
    expect(parser.evaluate(ast, ctx)).toEqual([[1, 2], [3, 4]]);
  });

  it('orders rows and columns correctly: commas across, semicolons down', () => {
    expect(run('={1,2,3}')).toEqual([[1, 2, 3]]);
    expect(run('={1;2;3}')).toEqual([[1], [2], [3]]);
    expect(run('={1,2;3,4}')).toEqual([[1, 2], [3, 4]]);
    expect(run('=ROWS({1,2;3,4;5,6})')).toBe(3);
    expect(run('=COLUMNS({1,2;3,4;5,6})')).toBe(2);
    expect(run('=INDEX({1,2;3,4},2,1)')).toBe(3);
    expect(run('=INDEX({1,2;3,4},1,2)')).toBe(2);
    expect(run('=INDEX({10,20,30},2)')).toBe(20);
    expect(run('=INDEX({10;20;30},3)')).toBe(30);
  });

  it('feeds aggregate functions', () => {
    expect(run('=SUM({1,2,3})')).toBe(6);
    expect(run('=SUM({1,2;3,4})')).toBe(10);
    expect(run('=SUM({-1,-2,-3})')).toBe(-6);
    expect(run('=MAX({1,5;3,4})')).toBe(5);
    expect(run('=COUNT({1,"a",TRUE,2})')).toBe(2);
    expect(run('=COUNTA({1,"a",TRUE})')).toBe(3);
    expect(run('=PRODUCT({2,3;4,5})')).toBe(120);
    expect(run('=SUMPRODUCT({1,2,3},{4,5,6})')).toBe(32);
    expect(run('=AVERAGE({2,4;6,8})')).toBe(5);
  });

  it('holds text, booleans and errors', () => {
    expect(run('={"a","b"}&"x"')).toEqual([['ax', 'bx']]);
    expect(run('=TEXTJOIN("-",TRUE,{"a","b";"c","d"})')).toBe('a-b-c-d');
    expect(run('=AND({TRUE,TRUE})')).toBe(true);
    expect(run('=OR({FALSE,FALSE})')).toBe(false);
    expect(run('={1,#N/A}')).toEqual([[1, '#N/A']]);
    expect(() => run('=SUM({1,#N/A})')).toThrow('#N/A');
    expect(run('=IFERROR({1,#DIV/0!},0)')).toEqual(expect.anything());
  });

  it('broadcasts against scalars, ranges and other arrays', () => {
    expect(run('={1,2,3}*2')).toEqual([[2, 4, 6]]);
    expect(run('={1,2,3}+{10,20,30}')).toEqual([[11, 22, 33]]);
    expect(run('={1;2}*{10,20}')).toEqual([[10, 20], [20, 40]]);
    expect(run('=A1:A3*{1,2}', { sheets: { S1: data } })).toEqual([[1, 2], [2, 4], [3, 6]]);
    expect(run('={1,2,3}>=2')).toEqual([[false, true, true]]);
  });

  it('keeps an array constant out of the dependency set', () => {
    expect(parser.parse('=SUM({1,2;3,4})').dependencies.size).toBe(0);
  });
});

describe('3D references', () => {
  const sheets = {
    Sheet1: grid({ A1: 1, B2: 10 }),
    Sheet2: grid({ A1: 2, B2: 20 }),
    Sheet3: grid({ A1: 3, B2: 30 }),
    'Q1 Data': grid({ A1: 100, B2: 1000 }),
  };

  it('aggregates the same cell across the span', () => {
    expect(run('=SUM(Sheet1:Sheet3!A1)', { sheets })).toBe(6);
    expect(run('=SUM(Sheet1:Sheet2!A1)', { sheets })).toBe(3);
    expect(run('=SUM(Sheet2:Sheet2!A1)', { sheets })).toBe(2);
    expect(run('=MAX(Sheet1:Sheet3!A1)', { sheets })).toBe(3);
    expect(run('=COUNT(Sheet1:Sheet3!A1)', { sheets })).toBe(3);
    expect(run('=AVERAGE(Sheet1:Sheet3!B2)', { sheets })).toBe(20);
  });

  it('aggregates a range, a whole column and a whole row across the span', () => {
    expect(run('=SUM(Sheet1:Sheet3!A1:B2)', { sheets })).toBe(6 + 60);
    expect(run('=SUM(Sheet1:Sheet3!A:A)', { sheets })).toBe(6);
    expect(run('=SUM(Sheet1:Sheet3!2:2)', { sheets })).toBe(60);
  });

  it('takes the span from workbook order, whichever way it is written', () => {
    expect(run('=SUM(Sheet3:Sheet1!A1)', { sheets })).toBe(6);
    const reordered = { Sheet3: sheets.Sheet3, Sheet1: sheets.Sheet1, Sheet2: sheets.Sheet2 };
    expect(run('=SUM(Sheet1:Sheet2!A1)', { sheets: reordered })).toBe(1 + 2);
    expect(run('=SUM(Sheet3:Sheet1!A1)', { sheets: reordered })).toBe(3 + 1);
  });

  it('covers quoted sheet names', () => {
    expect(run("=SUM('Sheet3:Q1 Data'!A1)", { sheets })).toBe(103);
  });

  it('combines with other arguments', () => {
    expect(run('=SUM(Sheet1:Sheet3!A1,10)', { sheets })).toBe(16);
    expect(run('=SUM(Sheet1:Sheet3!A1)*2', { sheets })).toBe(12);
  });

  it('resolves a relative reference against the formula cell', () => {
    expect(run('=SUM(Sheet1:Sheet3!A1)', { sheets, at: [0, 3] })).toBe(6);
    expect(run('=SUM(Sheet1:Sheet3!B2)', { sheets, at: [1, 1] })).toBe(60);
  });

  it('is #REF! when an end sheet is missing', () => {
    expect(() => run('=SUM(Sheet1:Gone!A1)', { sheets })).toThrow('#REF!');
    expect(() => run('=SUM(Gone:Sheet3!A1)', { sheets })).toThrow('#REF!');
    expect(() => run('=SUM(Gone:Sheet3!A1:B2)', { sheets })).toThrow('#REF!');
  });

  it('is #REF! without a sheet list in the context', () => {
    const { ast } = parser.parse('=SUM(Sheet1:Sheet3!A1)');
    expect(() => parser.evaluate(ast, { getCellValue: () => 1, getRangeValues: () => [] })).toThrow('#REF!');
  });

  it('is #VALUE! anywhere but a function argument', () => {
    expect(() => run('=Sheet1:Sheet3!A1', { sheets })).toThrow('#VALUE!');
    expect(() => run('=Sheet1:Sheet3!A1+1', { sheets })).toThrow('#VALUE!');
    expect(() => run('=Sheet1:Sheet3!A1:B2', { sheets })).toThrow('#VALUE!');
  });
});

describe('ranges built from reference-returning functions', () => {
  const sheets = { S1: grid({ A1: 1, A2: 2, A3: 3, A4: 4, A5: 5, B1: 10, B2: 20, B3: 30, C3: 300 }) };

  it('ends a range at INDEX', () => {
    expect(run('=SUM(A1:INDEX(A1:A5,3))', { sheets })).toBe(6);
    expect(run('=SUM(INDEX(A1:A5,2):A5)', { sheets })).toBe(2 + 3 + 4 + 5);
    expect(run('=SUM(INDEX(A1:A5,2):INDEX(A1:A5,4))', { sheets })).toBe(9);
    expect(run('=SUM(A1:INDEX(A1:C5,3,2))', { sheets })).toBe(1 + 2 + 3 + 10 + 20 + 30);
    expect(run('=INDEX(A1:A5,3):INDEX(A1:A5,3)', { sheets })).toBe(3);
  });

  it('ends a range at OFFSET', () => {
    expect(run('=SUM(A1:OFFSET(A1,2,0))', { sheets })).toBe(6);
    expect(run('=SUM(OFFSET(A1,1,0):A3)', { sheets })).toBe(5);
    expect(run('=SUM(OFFSET(A1,0,0,3,1))', { sheets })).toBe(6);
    expect(run('=SUM(OFFSET(A1:A2,1,0))', { sheets })).toBe(2 + 3);
    expect(run('=SUM(OFFSET(A:A,0,1))', { sheets, at: [0, 6] })).toBe(60);
    expect(() => run('=SUM(A1:OFFSET(A1,-5,0))', { sheets, at: [0, 6] })).toThrow('#REF!');
    expect(() => run('=OFFSET(1+1,0,0)', { sheets })).toThrow('#REF!');
    expect(() => run('=OFFSET()', { sheets })).toThrow('#REF!');
  });

  it('ends a range at INDIRECT', () => {
    expect(run('=SUM(A1:INDIRECT("A3"))', { sheets })).toBe(6);
    expect(run('=SUM(INDIRECT("A1"):INDIRECT("A3"))', { sheets })).toBe(6);
    expect(run('=SUM(INDIRECT("A1:A3"))', { sheets })).toBe(6);
    expect(run('=SUM(INDIRECT("A:A"))', { sheets })).toBe(15);
    expect(run('=SUM(INDIRECT("1:1"))', { sheets })).toBe(11);
    expect(run('=INDIRECT("B2")', { sheets })).toBe(20);
    expect(() => run('=INDIRECT("not a ref")', { sheets })).toThrow('#REF!');
  });

  it('takes the bounding box of the two ends, in either order', () => {
    expect(run('=SUM(C3:A1)', { sheets })).toBe(1 + 2 + 3 + 10 + 20 + 30 + 300);
    expect(run('=SUM(A1:INDEX(A1:C5,3,3))', { sheets })).toBe(1 + 2 + 3 + 10 + 20 + 30 + 300);
    expect(run('=SUM(INDEX(A1:C5,3,3):A1)', { sheets })).toBe(1 + 2 + 3 + 10 + 20 + 30 + 300);
  });

  it('refuses ends on different sheets', () => {
    const two = { S1: sheets.S1, S2: grid({ A1: 1 }) };
    expect(() => run('=SUM(A1:S2!A3)', { sheets: two })).toThrow('#VALUE!');
    expect(() => run('=SUM(S2!A1:INDEX(A1:A5,3))', { sheets: two })).toThrow('#VALUE!');
    expect(run('=SUM(S2!A1:S2!A1)', { sheets: two })).toBe(1);
    expect(run('=SUM(S2!A1:INDEX(S2!A1:A5,1))', { sheets: two })).toBe(1);
  });

  it('refuses non-reference ends', () => {
    expect(() => run('=SUM(A1:SEQUENCE(3))', { sheets })).toThrow('#VALUE!');
    expect(() => run('=SUM(A1:foo)', { sheets })).toThrow('#VALUE!');
  });

  it('feeds ROW, COLUMN, ROWS, COLUMNS and ISREF', () => {
    expect(run('=ROW(INDEX(A1:A5,4))', { sheets })).toBe(4);
    expect(run('=ROW(A2:INDEX(A1:A5,4))', { sheets })).toBe(2);
    expect(run('=ROWS(A1:INDEX(A1:A5,4))', { sheets })).toBe(4);
    expect(run('=COLUMNS(A1:INDEX(A1:C5,1,3))', { sheets })).toBe(3);
    expect(run('=COLUMN(OFFSET(A1,0,2))', { sheets })).toBe(3);
    expect(run('=ISREF(A1:INDEX(A1:A5,3))', { sheets })).toBe(true);
    expect(run('=ISREF(INDEX(A1:A5,3))', { sheets })).toBe(true);
    expect(run('=ISREF(INDEX({1,2},1))', { sheets })).toBe(false);
    expect(run('=ISREF(INDEX())', { sheets })).toBe(false);
    expect(run('=ISREF(A:A)', { sheets })).toBe(true);
    expect(run('=ISREF(1:1)', { sheets })).toBe(true);
    expect(run('=ISREF(Sheet1:Sheet3!A1)', { sheets })).toBe(true);
    expect(run('=ISREF((A1,B1))', { sheets })).toBe(true);
    expect(run('=ISREF(A1 B1)', { sheets })).toBe(true);
    expect(run('=ISREF(A1#)', { sheets })).toBe(true);
    expect(run('=ISREF(OFFSET(A1,1,1))', { sheets })).toBe(true);
    expect(run('=ISREF(INDIRECT("A1"))', { sheets })).toBe(true);
    expect(run('=ISREF(5)', { sheets })).toBe(false);
    expect(run('=ISREF("A1")', { sheets })).toBe(false);
    expect(run('=ISREF({1,2})', { sheets })).toBe(false);
    expect(run('=ISREF(SUM(A1))', { sheets })).toBe(false);
    expect(run('=ISREF()', { sheets })).toBe(false);
  });

  it('keeps ROW, COLUMN, ROWS and COLUMNS strict about non-references', () => {
    expect(() => run('=ROW(5)', { sheets })).toThrow('#VALUE!');
    expect(() => run('=COLUMN("x")', { sheets })).toThrow('#VALUE!');
    expect(() => run('=ROWS()', { sheets })).toThrow('#VALUE!');
    expect(() => run('=COLUMNS()', { sheets })).toThrow('#VALUE!');
    expect(run('=ROWS({1;2;3})', { sheets })).toBe(3);
    expect(run('=COLUMNS({1,2;3,4})', { sheets })).toBe(2);
    expect(run('=ROWS(5)', { sheets })).toBe(1);
  });

  it('indexes arrays and rejects bad INDEX positions', () => {
    expect(run('=INDEX(SEQUENCE(5),4)', { sheets })).toBe(4);
    expect(() => run('=INDEX(A1:A5,6)', { sheets })).toThrow('#REF!');
    expect(() => run('=INDEX(A1:A5,0)', { sheets })).toThrow('#REF!');
    expect(() => run('=INDEX(A1:C5,2)', { sheets })).toThrow('#REF!'); // 2-D range needs a column
    expect(run('=INDEX(B1:B3,2)', { sheets })).toBe(20);
    expect(run('=INDEX(A1:C1,3)', { sheets })).toBe(null);
    expect(run('=INDEX(A1:A5)', { sheets })).toBe(1);
  });
});

describe('CELL with the new reference forms', () => {
  const sheets = { S1: grid({ B3: 5 }) };
  it('reads the top-left of any reference', () => {
    expect(run('=CELL("address",B3:C4)', { sheets })).toBe('$B$3');
    expect(run('=CELL("row",B:B)', { sheets })).toBe(1);
    expect(run('=CELL("col",3:3)', { sheets })).toBe(1);
    expect(run('=CELL("address",INDEX(A1:C5,3,2))', { sheets })).toBe('$B$3');
    expect(run('=CELL("contents",B3)', { sheets })).toBe(5);
  });
  it('is #VALUE! for a non-reference', () => {
    expect(() => run('=CELL("row",5)', { sheets })).toThrow('#VALUE!');
  });
});

describe('union', () => {
  const sheets = { S1: grid({ A1: 1, A2: 2, A3: 3, A4: 4, A5: 5, C1: 100 }) };

  it('stacks the areas of a parenthesised union for aggregate functions', () => {
    expect(run('=SUM((A1:A2,A4))', { sheets })).toBe(7);
    expect(run('=SUM((A1,A2,A3))', { sheets })).toBe(6);
    expect(run('=COUNT((A1:A5,C1))', { sheets })).toBe(6);
    expect(run('=MAX((A1:A2,C1))', { sheets })).toBe(100);
    expect(run('=SUM((A1:A2,C1),A5)', { sheets })).toBe(108);
    expect(run('=SUM(((A1,A2),A3))', { sheets })).toBe(6);
    expect(run('=SUM((A1,(A2,A3)))', { sheets })).toBe(6);
  });

  it('reaches ref functions and lazy functions through their arguments', () => {
    expect(run('=COUNTBLANK((A1:A2,Z1))', { sheets })).toBe(1);
    expect(run('=IF(TRUE,SUM((A1,A2)),0)', { sheets })).toBe(3);
  });

  it('is #VALUE! outside a function argument', () => {
    expect(() => run('=(A1,A2)', { sheets })).toThrow('#VALUE!');
    expect(() => run('=(A1,A2)+1', { sheets })).toThrow('#VALUE!');
  });
});

describe('intersection', () => {
  const sheets = { S1: grid({ A1: 1, A2: 2, A3: 3, A4: 4, B2: 20, B3: 30, B4: 40, C3: 300 }) };

  it('returns the overlap of two ranges', () => {
    expect(run('=SUM(A1:A3 A2:A4)', { sheets })).toBe(5);
    expect(run('=SUM(A1:C3 B2:D4)', { sheets })).toBe(20 + 30 + 300);
    expect(run('=SUM(A:A 2:2)', { sheets })).toBe(2);
    expect(run('=SUM(A:B 3:3)', { sheets })).toBe(3 + 30);
    expect(run('=A1:B3 B2:C4', { sheets })).toEqual([[20], [30]]);
  });

  it('collapses a single-cell overlap to a scalar', () => {
    expect(run('=A1:A3 A2:C2', { sheets })).toBe(2);
    expect(run('=B3 A3:C3', { sheets })).toBe(30);
  });

  it('is #NULL! when the ranges do not overlap', () => {
    expect(() => run('=SUM(A1:A2 A3:A4)', { sheets })).toThrow('#NULL!');
    expect(() => run('=A1 B1', { sheets })).toThrow('#NULL!');
  });

  it('is #VALUE! across sheets or with a non-reference', () => {
    const two = { S1: sheets.S1, S2: grid({ A1: 1 }) };
    expect(() => run('=S1!A1:A3 S2!A1:A3', { sheets: two })).toThrow('#VALUE!');
    expect(() => run('=A1:A3 SEQUENCE(3)', { sheets })).toThrow('#VALUE!');
    expect(() => run('=SUM(S2!A1:A3 A1:A2)', { sheets: two })).toThrow('#VALUE!');
    expect(run('=SUM(S2!A1:A3 S2!A1:A2)', { sheets: two })).toBe(1);
  });

  it('chains', () => {
    expect(run('=A1:C4 A2:B4 B1:C3', { sheets })).toEqual([[20], [30]]);
  });
});

describe('implicit intersection (@)', () => {
  const sheets = { S1: grid({ A1: 1, A2: 2, A3: 3, B1: 10, C1: 100, B2: 20, A5: 5 }) };

  it('picks the cell of a vertical range in the formula row', () => {
    expect(run('=@A1:A3', { sheets, at: [1, 5] })).toBe(2);
    expect(run('=@A1:A3', { sheets, at: [2, 5] })).toBe(3);
    expect(run('=@A:A', { sheets, at: [4, 5] })).toBe(5);
    expect(run('=@A:A', { sheets, at: [9, 5] })).toBe(null);
  });

  it('picks the cell of a horizontal range in the formula column', () => {
    expect(run('=@A1:C1', { sheets, at: [7, 2] })).toBe(100);
    expect(run('=@1:1', { sheets, at: [7, 1] })).toBe(10);
  });

  it('picks the cell of a two-dimensional range in the formula row and column', () => {
    expect(run('=@A1:C3', { sheets, at: [1, 1] })).toBe(20);
  });

  it('is #VALUE! outside the range', () => {
    expect(() => run('=@A1:A3', { sheets, at: [7, 5] })).toThrow('#VALUE!');
    expect(() => run('=@A1:C1', { sheets, at: [7, 9] })).toThrow('#VALUE!');
    expect(() => run('=@A1:A3', { sheets, at: [0, 0] })).not.toThrow(); // formula row 1 is inside A1:A3
  });

  it('passes a single cell and a scalar through', () => {
    expect(run('=@A1', { sheets })).toBe(1);
    expect(run('=@5')).toBe(5);
    expect(run('=@"x"')).toBe('x');
    expect(run('=@(1+2)')).toBe(3);
  });

  it('takes the top-left value of a computed array', () => {
    expect(run('=@SEQUENCE(3)')).toBe(1);
    expect(run('=@{7,8;9,10}')).toBe(7);
    expect(run('=@{7,8}')).toBe(7);
    expect(run('=@(A1:A3*2)', { sheets })).toBe(2);
  });

  it('keeps a falsy top-left value', () => {
    expect(run('=@{0}')).toBe(0);
    expect(run('=@{FALSE,1}')).toBe(false);
  });

  it('composes with arithmetic', () => {
    expect(run('=@A1:A3+1', { sheets, at: [1, 5] })).toBe(3);
    expect(run('=1+@A1:A3', { sheets, at: [2, 5] })).toBe(4);
    expect(run('=@A1:A3*@A1:A3', { sheets, at: [1, 5] })).toBe(4);
  });
});

describe('the spill operator (#)', () => {
  const spills = { 'S1|0:0': { rows: 3, cols: 1, values: [[1], [2], [3]] } };
  const sheets = { S1: grid({ A1: 1 }) };

  it('returns the values of the spill range anchored at the cell', () => {
    expect(run('=A1#', { sheets, spills, at: [0, 3] })).toEqual([[1], [2], [3]]);
    expect(run('=SUM(A1#)', { sheets, spills, at: [0, 3] })).toBe(6);
    expect(run('=ROWS(A1#)', { sheets, spills, at: [0, 3] })).toBe(3);
    expect(run('=COLUMNS(A1#)', { sheets, spills, at: [0, 3] })).toBe(1);
    expect(run('=A1#*2', { sheets, spills, at: [0, 3] })).toEqual([[2], [4], [6]]);
    expect(run('=SUM($A$1#)', { sheets, spills, at: [5, 5] })).toBe(6);
  });

  it('returns a copy, so callers cannot edit the spill', () => {
    const values = run('=A1#', { sheets, spills, at: [0, 3] }) as number[][];
    values[0][0] = 99;
    expect(spills['S1|0:0'].values[0][0]).toBe(1);
  });

  it('reads another sheet', () => {
    const two = { S1: sheets.S1, S2: grid({}) };
    expect(run('=SUM(S2!B2#)', { sheets: two, spills: { 'S2|1:1': { rows: 2, cols: 1, values: [[5], [6]] } } })).toBe(11);
  });

  it('is #REF! when the cell anchors no spill', () => {
    expect(() => run('=A2#', { sheets, spills, at: [0, 3] })).toThrow('#REF!');
    expect(() => run('=SUM(A2#)', { sheets, spills, at: [0, 3] })).toThrow('#REF!');
    expect(() => run('=A1#', { sheets })).toThrow('#REF!');
  });

  it('is #REF! without spill support in the context', () => {
    const { ast } = parser.parse('=A1#');
    expect(() => parser.evaluate(ast, { getCellValue: () => 1, getRangeValues: () => [] })).toThrow('#REF!');
  });

  it('is a reference for ROW, COLUMN and the range operator', () => {
    expect(run('=ROW(A1#)', { sheets, spills, at: [0, 3] })).toBe(1);
    // The mock grid holds only A1, so the box A1:B3 sums to 1 (the workbook overlays the spilled values).
    expect(run('=SUM(A1#:B1)', { sheets, spills, at: [0, 3] })).toBe(1);
  });
});

describe('error literals', () => {
  it('evaluate to the error value', () => {
    expect(run('=#NULL!')).toBe('#NULL!');
    expect(run('=#N/A')).toBe('#N/A');
    expect(run('=#DIV/0!')).toBe('#DIV/0!');
    expect(run('=#REF!')).toBe('#REF!');
    expect(run('=#NAME?')).toBe('#NAME?');
    expect(run('=#VALUE!')).toBe('#VALUE!');
    expect(run('=#NUM!')).toBe('#NUM!');
  });

  it('are caught by the error functions', () => {
    expect(run('=ISERROR(#NULL!)')).toBe(true);
    expect(run('=ISNA(#N/A)')).toBe(true);
    expect(run('=ISNA(#NULL!)')).toBe(false);
    expect(run('=IFERROR(#NULL!,"fallback")')).toBe('fallback');
    expect(run('=IF(ISERROR(#REF!),"bad","ok")')).toBe('bad');
    expect(run('=ERROR.TYPE(#NULL!)')).toBe(1);
    expect(run('=ERROR.TYPE(#DIV/0!)')).toBe(2);
    expect(run('=ERROR.TYPE(#N/A)')).toBe(7);
  });

  it('propagate through aggregates', () => {
    expect(() => run('=SUM(1,#NULL!)')).toThrow('#NULL!');
    expect(() => run('=TEXT(#NUM!,"0")')).toThrow('#NUM!');
  });
});

describe('names', () => {
  it('evaluate to #NAME? until defined names are supported', () => {
    expect(() => run('=Quarter1')).toThrow('#NAME? Name Quarter1 not found');
    expect(run('=IFERROR(Quarter1,"none")')).toBe('none');
  });
});

describe('best-effort dependency set', () => {
  it('lists cells and small ranges, but never whole-column, whole-row or 3D references', () => {
    expect([...parser.parse('=A1+B2').dependencies].sort()).toEqual(['0:0', '1:1']);
    expect([...parser.parse('=SUM(A1:A3)').dependencies].sort()).toEqual(['0:0', '1:0', '2:0']);
    expect(parser.parse('=SUM(A:A)').dependencies.size).toBe(0);
    expect(parser.parse('=SUM(1:1)').dependencies.size).toBe(0);
    expect(parser.parse('=SUM(Sheet1:Sheet3!A1)').dependencies.size).toBe(0);
    expect([...parser.parse('=Other!B2').dependencies]).toEqual(['Other!1:1']);
    expect([...parser.parse('=SUM(Other!A1:A2)').dependencies].sort()).toEqual(['Other!0:0', 'Other!1:0']);
    expect(parser.parse('=SUM(A1:A1048576)').dependencies.size).toBe(0);
    expect([...parser.parse('=SUM((A1,B1))').dependencies].sort()).toEqual(['0:0', '0:1']);
  });

  it('reports a syntax error with its message and an empty AST', () => {
    const result = parser.parse('=SUM(1,');
    expect(result.error).toBe('#ERROR! Unexpected end of formula');
    expect(result.ast).toEqual({ type: 'string', value: '' });
    expect(parser.parse('=' + '1+'.repeat(6000) + '1').error).toBe('#ERROR! Formula too long');
  });
});

describe('edge cases of reference resolution', () => {
  const sheets = { S1: grid({ A1: 1, A2: 2, B1: 10, XFD1: 5 }) };

  it('is #REF! for the shape of a spill that does not exist', () => {
    expect(() => run('=ROWS(A2#)', { sheets })).toThrow('#REF!');
  });

  it('reads a range spanning every column as a whole row', () => {
    expect(run('=SUM(A1:INDEX(1:1,16384))', { sheets })).toBe(1 + 10 + 5);
    expect(run('=SUM(A1:INDEX(A:A,1048576))', { sheets })).toBe(1 + 2);
  });

  it('defaults the omitted OFFSET arguments', () => {
    expect(run('=OFFSET(A1)', { sheets })).toBe(1);
    expect(run('=OFFSET(A1,1)', { sheets })).toBe(2);
    expect(run('=OFFSET(A1,0,1)', { sheets })).toBe(10);
    expect(run('=SUM(OFFSET(A1,0,0,2))', { sheets })).toBe(3);
  });

  it('records absolute range corners in the best-effort dependency set', () => {
    expect([...parser.parse('=SUM($A$1:$A$2)').dependencies].sort()).toEqual(['0:0', '1:0']);
    expect([...parser.parse('=SUM($A1:B$2)', 5, 5).dependencies].sort()).toEqual(['0:0', '0:1', '1:0', '1:1']);
  });

  it('copes with a context that returns flat arrays', () => {
    const flat = (values: unknown[]): EvaluationContext => ({
      getCellValue: () => null,
      getRangeValues: () => values as unknown[][],
    });
    const eval1 = (formula: string, values: unknown[]) => {
      const { ast } = parser.parse(formula);
      return parser.evaluate(ast, flat(values));
    };
    expect(eval1('=SUM((A1:A2,B1))', [1, 2])).toBe(3);
    expect(eval1('=@IFERROR(A1:A2,0)', [7, 8])).toBe(7);
    expect(eval1('=@IFERROR(A1:A2,0)', [])).toBeNull();
    expect(eval1('=@IFERROR(A1:A2,0)', [[], []])).toBeNull();
  });
});

describe('error operands (a deleted reference)', () => {
  const sheets = { S1: grid({ A1: 1 }) };
  it('propagate their error through the reference operators and functions', () => {
    expect(() => run('=SUM(#REF!:A1)', { sheets })).toThrow('#REF!');
    expect(() => run('=SUM(A1:#NAME?)', { sheets })).toThrow('#NAME?');
    expect(() => run('=#REF! #REF!', { sheets })).toThrow('#REF!');
    expect(() => run('=SUM((#REF!,A1))', { sheets })).toThrow('#REF!');
    expect(() => run('=ROW(#REF!)', { sheets })).toThrow('#REF!');
    expect(() => run('=OFFSET(#REF!,1,1)', { sheets })).toThrow('#REF!');
    expect(run('=IFERROR(SUM(#REF!:A1),"gone")', { sheets })).toBe('gone');
    expect(() => run('=INDEX(#REF!,3)', { sheets })).toThrow('#REF!');
    expect(run('=ISREF(#REF!)', { sheets })).toBe(false);
  });
});

describe('shape and information functions over unusual values', () => {
  const sheets = { S1: grid({ A1: 1 }) };
  const flatCtx = (values: unknown): EvaluationContext => ({ getCellValue: () => null, getRangeValues: () => values as unknown[][] });
  const withCtx = (formula: string, ctx: EvaluationContext) => parser.evaluate(parser.parse(formula).ast, ctx);

  it('measures a flat, an empty and a non-array result for ROWS and COLUMNS', () => {
    expect(withCtx('=ROWS(IFERROR(A1:A3,0))', flatCtx([1, 2, 3]))).toBe(1);
    expect(withCtx('=COLUMNS(IFERROR(A1:A3,0))', flatCtx([1, 2, 3]))).toBe(3);
    expect(withCtx('=ROWS(IFERROR(A1:A3,0))', flatCtx([]))).toBe(1);
    expect(withCtx('=COLUMNS(IFERROR(A1:A3,0))', flatCtx([]))).toBe(0);
    expect(run('=COLUMNS(5)', { sheets })).toBe(1);
    expect(run('=COLUMNS({1,2,3;4,5,6})', { sheets })).toBe(3);
  });

  it('ISFORMULA needs a cell reference and a context that can answer', () => {
    expect(() => run('=ISFORMULA(5)', { sheets })).toThrow('#VALUE!');
    expect(withCtx('=ISFORMULA(A1)', { getCellValue: () => null, getRangeValues: () => [] })).toBe(false);
  });

  it('CELL("contents") shows a blank as 0', () => {
    expect(run('=CELL("contents",Z9)', { sheets })).toBe(0);
  });

  it('refuses an operator or a node type it does not know', () => {
    const internal = parser as unknown as { applyOperator(op: string, a: unknown, b: unknown): unknown };
    expect(() => internal.applyOperator('?', 1, 2)).toThrow('#ERROR! Unknown operator: ?');
    const ctx = { getCellValue: () => null, getRangeValues: () => [] };
    expect(parser.evaluate({ type: 'bogus' } as never, ctx)).toBeNull();
  });
});

describe('array operators', () => {
  const flat = (values: unknown): EvaluationContext => ({ getCellValue: () => null, getRangeValues: () => values as unknown[][] });
  const evalWith = (formula: string, ctx: EvaluationContext) => parser.evaluate(parser.parse(formula).ast, ctx);

  it('broadcasts over a flat array returned by the context', () => {
    expect(evalWith('=A1:A3+1', flat([1, 2, 3]))).toEqual([[2], [3], [4]]); // a flat array reads as a column
  });

  it('turns an operator failure that is not a spreadsheet error into #VALUE! for that element', () => {
    const internal = parser as unknown as { applyOperator: (...args: unknown[]) => unknown };
    const original = internal.applyOperator;
    internal.applyOperator = () => {
      throw new Error('boom');
    };
    try {
      expect(run('={1,2}+1')).toEqual([['#VALUE!', '#VALUE!']]);
    } finally {
      internal.applyOperator = original;
    }
  });

  it('evaluates a formula written without the leading =', () => {
    expect(parser.evaluate(parser.parse('1+1').ast, flat([]))).toBe(2);
  });
});
