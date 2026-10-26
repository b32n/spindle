import { WorkbookImpl } from './workbook';

// Table-driven Excel-conformance checks, run through the real parser/evaluator.
// Each row: formula, cells it reads (A1 notation -> value), expected result.
// The formula is placed in AA1, away from the data.

type Cells = Record<string, unknown>;
type Row = [formula: string, cells: Cells, expected: unknown];

function a1(ref: string): { row: number; col: number } {
  const m = /^([A-Z]+)(\d+)$/.exec(ref)!;
  const col = [...m[1]].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
  return { row: Number(m[2]) - 1, col };
}

function evalFormula(formula: string, cells: Cells = {}): unknown {
  const w = new WorkbookImpl('wb', 'WB');
  for (const [ref, v] of Object.entries(cells)) {
    const { row, col } = a1(ref);
    if (typeof v === 'string' && v.startsWith('=')) w.setFormula(undefined, row, col, v);
    else w.setCellValue(undefined, row, col, v);
  }
  w.setFormula(undefined, 0, 26, formula);
  return w.getCellValue(undefined, 0, 26);
}

/** The block a formula in AA1 produces (its value plus any spill), `rows` x `cols`. */
function evalGrid(formula: string, cells: Cells, rows: number, cols: number): unknown[][] {
  const w = new WorkbookImpl('wb', 'WB');
  for (const [ref, v] of Object.entries(cells)) {
    const { row, col } = a1(ref);
    w.setCellValue(undefined, row, col, v);
  }
  w.setFormula(undefined, 0, 26, formula);
  const out: unknown[][] = [];
  for (let r = 0; r < rows; r++) {
    out.push(Array.from({ length: cols }, (_, c) => w.getCellCalculatedValue(undefined, r, 26 + c)));
  }
  return out;
}

function run(rows: Row[]): void {
  it.each(rows)('%s  %j', (formula, cells, expected) => {
    expect(evalFormula(formula, cells)).toEqual(expected);
  });
}

describe('conformance: blank cells (B1 is empty unless set)', () => {
  run([
    ['=ISBLANK(B1)', {}, true],
    ['=ISBLANK(B1)', { B1: 0 }, false],
    ['=ISBLANK(B1)', { B1: '' }, false],
    ['=ISBLANK(B1)', { B1: ' ' }, false],
    ['=ISBLANK(B1)', { B1: '="" ' }, false],
    ['=ISBLANK(B1)', { B1: 5 }, false],
    ['=IF(ISBLANK(B1),"empty","full")', {}, 'empty'],
    ['=ISNUMBER(B1)', {}, false],
    ['=ISNUMBER(B1)', { B1: 0 }, true],
    ['=ISTEXT(B1)', {}, false],
    ['=B1', {}, 0],
    ['=B1+1', {}, 1],
    ['=B1*5', {}, 0],
    ['=-B1', {}, 0],
    ['=B1&"x"', {}, 'x'],
    ['=B1&B2', {}, ''],
    ['=LEN(B1)', {}, 0],
    ['=B1=0', {}, true],
    ['=B1=""', {}, true],
    ['=B1=FALSE', {}, true],
    ['=B1<>0', {}, false],
    ['=B1="a"', {}, false],
    ['=B1>-1', {}, true],
    ['=B1<1', {}, true],
    ['=B1<"a"', {}, true],
    ['=B1>=0', {}, true],
    ['=COUNT(B1)', {}, 0],
    ['=COUNTA(B1)', {}, 0],
    ['=COUNT(B1:B3)', { B2: 4 }, 1],
    ['=COUNTBLANK(B1:B3)', { B2: 4 }, 2],
    ['=SUM(B1,2)', {}, 2],
    ['=AVERAGE(B1:B3)', { B2: 4 }, 4],
    ['=IF(B1,"t","f")', {}, 'f'],
    ['=N(B1)', {}, 0],
    ['=B1&B1=""', {}, true],
    ['=ISBLANK(A1)', { A1: '=B1' }, false], // a formula cell is never blank, even if it shows 0
    ['=A1', { A1: '=B1' }, 0],
    ['=ISBLANK(B1)', { B1: null }, true], // a cleared cell is blank again
    ['=ISNUMBER(A1)', { A1: '=B1' }, true],
  ]);

  it('re-evaluates when a value is typed into and cleared from the referenced cell', () => {
    const w = new WorkbookImpl('wb', 'WB');
    w.setFormula(undefined, 0, 0, '=ISBLANK(B1)');
    expect(w.getCellValue(undefined, 0, 0)).toBe(true);
    w.setCellValue(undefined, 0, 1, 'x');
    expect(w.getCellValue(undefined, 0, 0)).toBe(false);
    w.setCellValue(undefined, 0, 1, null);
    expect(w.getCellValue(undefined, 0, 0)).toBe(true);
  });
});

describe('conformance: operator precedence and unary signs', () => {
  run([
    ['=-2^2', {}, 4],
    ['=2^-2', {}, 0.25],
    ['=-2^-2', {}, 0.25],
    ['=-A1^2', { A1: 3 }, 9],
    ['=-(2)^2', {}, 4],
    ['=-(1+1)^2', {}, 4],
    ['=--2', {}, 2],
    ['=+2^2', {}, 4],
    ['=-2+3', {}, 1],
    ['=-2*3', {}, -6],
    ['=-6/2', {}, -3],
    ['=2*-3', {}, -6],
    ['=3-2-1', {}, 0],
    ['=2^3^2', {}, 64], // Excel is left-associative: (2^3)^2
    ['=2+3*4', {}, 14],
    ['=(2+3)*4', {}, 20],
    ['=1+2&3', {}, '33'],
    ['=-2^2+1', {}, 5],
    ['=1-2^2', {}, -3],
    ['=-1E-2*100', {}, -1],
    ['=1=1+0', {}, true],
  ]);

  it('keeps the formula text when re-rendered', () => {
    const w = new WorkbookImpl('wb', 'WB');
    w.setFormula(undefined, 0, 0, '=-2^2');
    expect(w.getCell(undefined, 0, 0)?.formula).toBe('=-2^2');
    w.setFormula(undefined, 0, 2, '=-A5^2');
    expect(w.getCell(undefined, 0, 2)?.formula).toBe('=-A5^2');
  });

  it.each([
    '=(1+2)*3', '=2*(3+4)', '=-(1+2)', '=A5-(B5-C5)', '=A5-B5-C5', '=A5/(B5*C5)', '=2^(3^2)', '=(2^3)^2',
    '=-(2^2)', '=(A5+B5)&"x"', '=(1+2)=3', '=1+(2=3)', '=2--2', '=2*-3', '=-A5*B5', '=(-A5)^2',
  ])('%s survives a render/re-parse round trip with the same value', (formula) => {
    const w = new WorkbookImpl('wb', 'WB');
    w.setCellValue(undefined, 4, 0, 3);
    w.setCellValue(undefined, 4, 1, 4);
    w.setCellValue(undefined, 4, 2, 5);
    w.setFormula(undefined, 0, 0, formula);
    const first = w.getCellValue(undefined, 0, 0);
    const text = w.getCell(undefined, 0, 0)!.formula!;
    w.setFormula(undefined, 1, 0, text);
    expect(w.getCellValue(undefined, 1, 0)).toEqual(first);
    // and the rendered text is stable (idempotent)
    expect(w.getCell(undefined, 1, 0)!.formula).toBe(text);
  });
});

describe('conformance: array arithmetic and broadcasting', () => {
  const data = { A2: 1, A3: 2, A4: 3, B2: 4, B3: 5, B4: 6 };
  run([
    ['=SUMPRODUCT(A2:A4*B2:B4)', data, 32],
    ['=SUMPRODUCT(A2:A4,B2:B4)', data, 32],
    ['=SUMPRODUCT(A2:A4*B2:B4*2)', data, 64],
    ['=SUMPRODUCT((A2:A4>1)*1)', data, 2],
    ['=SUMPRODUCT(--(A2:A4>1))', data, 2],
    ['=SUMPRODUCT((A2:A4>1)*(B2:B4<6))', data, 1],
    ['=SUMPRODUCT(A2:A4+1)', data, 9],
    ['=SUMPRODUCT(A2:A4-B2:B4)', data, -9],
    ['=SUMPRODUCT(B2:B4/A2:A4)', data, 4 + 2.5 + 2],
    ['=SUMPRODUCT(A2:A4^2)', data, 14],
    ['=SUM(A2:A4*B2:B4)', data, 32],
    ['=MAX(A2:A4*B2:B4)', data, 18],
    ['=SUMPRODUCT(-A2:A4)', data, -6],
    ['=SUMPRODUCT(1-(A2:A4>1))', data, 1],
    ['=SUMPRODUCT((A2:A4>1)*B2:B4)', data, 11],
    ['=COUNTIF(A2:A4,">1")', data, 2],
    ['=SUMPRODUCT(A2:A4*B2:B4)', { A2: 1, A3: 'x', A4: 3, B2: 4, B3: 5, B4: 6 }, '#VALUE!'], // text element -> #VALUE! propagates
    ['=SUMPRODUCT(A2:A4*C2:C4)', data, 0], // blank column multiplies as 0
    ['=1+A2:A4', data, 2], // a top-level array formula shows its first element
  ]);

  it.each<[string, Cells, number, number, unknown[][]]>([
    ['=A2:A4*2', data, 3, 1, [[2], [4], [6]]],
    ['=A2:A4+B2:B4', data, 3, 1, [[5], [7], [9]]],
    ['=A2:A4>1', data, 3, 1, [[false], [true], [true]]],
    ['=A2:A4&"x"', data, 3, 1, [['1x'], ['2x'], ['3x']]],
    ['=-A2:A4', data, 3, 1, [[-1], [-2], [-3]]],
    ['=2*A2:B2', data, 1, 2, [[2, 8]]],
    ['=A2:A4*E6:G6', { ...data, E6: 1, F6: 2, G6: 3 }, 3, 3, [[1, 2, 3], [2, 4, 6], [3, 6, 9]]],
    ['=A2:A4+B2:B3', data, 3, 1, [[5], [7], ['#N/A']]], // size mismatch -> #N/A past the shorter operand
    ['=B2:B4/(A2:A4-2)', data, 3, 1, [[-4], ['#DIV/0!'], [6]]], // a failing element does not poison the rest
    ['=SEQUENCE(3)*10', data, 3, 1, [[10], [20], [30]]],
  ])('%s spills element-wise', (formula, cells, rows, cols, expected) => {
    expect(evalGrid(formula, cells, rows, cols)).toEqual(expected);
  });

  it('re-evaluates the array when an input cell changes', () => {
    const w = new WorkbookImpl('wb', 'WB');
    w.setCellValue(undefined, 1, 0, 1);
    w.setCellValue(undefined, 2, 0, 2);
    w.setFormula(undefined, 0, 2, '=SUMPRODUCT(A2:A3*A2:A3)');
    expect(w.getCellValue(undefined, 0, 2)).toBe(5);
    w.setCellValue(undefined, 2, 0, 3);
    expect(w.getCellValue(undefined, 0, 2)).toBe(10);
  });
});
