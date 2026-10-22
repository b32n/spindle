import { WorkbookImpl } from './workbook';

// Regression tests for known formula-evaluation bugs, run through the real
// parser/evaluator on a sheet. Each uses `test.failing`: it passes while the bug
// exists and starts FAILING once the bug is fixed — at which point flip it to a
// plain `it` (and un-fixme the matching E2E test in e2e/sheets/formulas-and-fill.spec.ts).

describe('known formula bugs (expected failures)', () => {
  // Fixed: https://github.com/b32n/spindle/issues/66
  it('ISBLANK(ref to an empty cell) is TRUE', () => {
    const w = new WorkbookImpl('wb', 'WB');
    w.setCellValue(undefined, 0, 0, 5); // A1 = 5; B1 stays empty
    w.setFormula(undefined, 0, 2, '=ISBLANK(B1)'); // C1
    expect(w.getCellValue(undefined, 0, 2)).toBe(true);
  });

  // https://ISSUE_SUMPRODUCT
  test.failing('SUMPRODUCT(A2:A4*B2:B4) multiplies the arrays element-wise', () => {
    const w = new WorkbookImpl('wb', 'WB');
    [1, 2, 3].forEach((v, i) => w.setCellValue(undefined, i + 1, 0, v)); // A2:A4
    [4, 5, 6].forEach((v, i) => w.setCellValue(undefined, i + 1, 1, v)); // B2:B4
    w.setFormula(undefined, 0, 3, '=SUMPRODUCT(A2:A4*B2:B4)'); // D1
    expect(w.getCellValue(undefined, 0, 3)).toBe(32); // 1*4 + 2*5 + 3*6
  });

  // Fixed: https://github.com/b32n/spindle/issues/68
  it('=-2^2 is 4 (unary minus binds tighter than ^, as in Excel)', () => {
    const w = new WorkbookImpl('wb', 'WB');
    w.setFormula(undefined, 0, 0, '=-2^2');
    expect(w.getCellValue(undefined, 0, 0)).toBe(4);
  });
});
