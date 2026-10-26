import { WorkbookImpl } from './workbook';
import { FormulaParser } from './formula-parser/parser';
import type { EvaluationContext } from './formula-parser/types';
import { formatNumber } from './utils/format-utils';

const parser = new FormulaParser();
const ctx: EvaluationContext = { getCellValue: () => null, getRangeValues: () => [] };
const ev = (f: string): unknown => {
  const { ast, error } = parser.parse(f);
  if (error) throw new Error(error);
  return parser.evaluate(ast, ctx);
};

describe('sort recalculates dependent formulas', () => {
  it('a position-based formula recomputes after a sort', () => {
    const wb = new WorkbookImpl('wb', 'WB');
    wb.setCellValue(undefined, 0, 0, 'Header'); // A1 header (kept in place by sort)
    [5, 3, 1].forEach((v, r) => wb.setCellValue(undefined, r + 1, 0, v)); // A2..A4 data
    // A header-row (unsorted) formula that reads by POSITION — its value changes
    // when the data reorders, so it must recompute after sort.
    wb.setFormula(undefined, 0, 2, '=INDEX(A2:A4,1)'); // C1 = value at position 1
    expect(wb.getCellValue(undefined, 0, 2)).toBe(5);

    wb.setSortOrder([{ column: 0, direction: 'asc' }]);
    wb.sortSheet();
    expect(wb.getCellValue(undefined, 1, 0)).toBe(1); // A2 now the smallest
    expect(wb.getCellValue(undefined, 0, 2)).toBe(1); // INDEX recomputed (was stale at 5)
  });
});

describe('spill overlay is rebuilt on structural edits (no ghosts / spurious #SPILL!)', () => {
  const calc = (wb: WorkbookImpl, r: number, c: number) => wb.getCellCalculatedValue(undefined, r, c);

  it('inserting a row above a spill anchor re-spills cleanly', () => {
    const wb = new WorkbookImpl('wb', 'WB');
    wb.setFormula(undefined, 2, 0, '=SEQUENCE(3)'); // A3 spills A3:A5 = 1,2,3
    expect(calc(wb, 2, 0)).toBe(1);
    expect(calc(wb, 4, 0)).toBe(3);

    wb.getSheet().insertRows(0, 1); // anchor shifts to A4
    expect(calc(wb, 3, 0)).toBe(1); // anchor, not #SPILL!
    expect(calc(wb, 5, 0)).toBe(3); // spilled tail
    expect(calc(wb, 3, 0)).not.toBe('#SPILL!');
  });

  it('deleting the anchor row removes the spilled ghosts', () => {
    const wb = new WorkbookImpl('wb', 'WB');
    wb.setFormula(undefined, 1, 0, '=SEQUENCE(3)'); // A2 spills A2:A4 = 1,2,3
    expect(calc(wb, 3, 0)).toBe(3);
    wb.getSheet().deleteRows(1, 1); // remove the anchor row
    // The formula is gone; no spilled ghost should remain at the old positions.
    expect(calc(wb, 1, 0)).toBe(0); // empty
    expect(calc(wb, 2, 0)).toBe(0);
  });
});

describe('#CIRCULAR! is a recognized error value', () => {
  it('propagates through aggregations and is caught by ISERROR/IFERROR', () => {
    const wb = new WorkbookImpl('wb', 'WB');
    wb.setCircularReferenceMode('error'); // the diagnostic token (default mode evaluates a cycle to 0)
    wb.setFormula(undefined, 0, 0, '=A1'); // A1 self-reference → #CIRCULAR!
    expect(wb.getCellValue(undefined, 0, 0)).toBe('#CIRCULAR!');
    wb.setCellValue(undefined, 1, 0, 5); // A2
    wb.setFormula(undefined, 0, 1, '=SUM(A1:A2)'); // includes the circular cell
    wb.setFormula(undefined, 0, 2, '=ISERROR(A1)');
    wb.setFormula(undefined, 0, 3, '=IFERROR(A1,99)');
    expect(wb.getCellValue(undefined, 0, 1)).toBe('#CIRCULAR!');
    expect(wb.getCellValue(undefined, 0, 2)).toBe(true);
    expect(wb.getCellValue(undefined, 0, 3)).toBe(99);
  });
});

describe('AND/OR/XOR propagate a range error', () => {
  it('AND over a range containing #DIV/0! returns the error', () => {
    const wb = new WorkbookImpl('wb', 'WB');
    wb.setCellValue(undefined, 0, 0, true);
    wb.setFormula(undefined, 1, 0, '=1/0'); // A2 = #DIV/0!
    wb.setFormula(undefined, 0, 2, '=AND(A1:A2)');
    wb.setFormula(undefined, 1, 2, '=OR(A1:A2)');
    expect(wb.getCellValue(undefined, 0, 2)).toBe('#DIV/0!');
    expect(wb.getCellValue(undefined, 1, 2)).toBe('#DIV/0!');
  });
  it('AND/OR without an error still work', () => {
    expect(ev('=AND(TRUE,1,TRUE)')).toBe(true);
    expect(ev('=OR(FALSE,0)')).toBe(false);
  });
});

describe('reference to a not-yet-existing sheet is #REF! and recovers when the sheet appears', () => {
  it('a same-position ghost ref is #REF!, not a false #CIRCULAR!, and recovers', () => {
    const wb = new WorkbookImpl('wb', 'WB');
    // A1 = =Ghost!A1 — same position on a non-existent sheet.
    wb.setFormula(undefined, 0, 0, '=Ghost!A1');
    expect(wb.getCellValue(undefined, 0, 0)).toBe('#REF!'); // not #CIRCULAR!

    const ghost = wb.addSheet('Ghost');
    wb.setCellValue(ghost.id, 0, 0, 42); // Ghost!A1 = 42
    // The formula recovers once the sheet exists.
    expect(wb.getCellValue(undefined, 0, 0)).toBe(42);
  });

  it('does not create a spurious dependency on the local same-index cell', () => {
    const wb = new WorkbookImpl('wb', 'WB');
    wb.setFormula(undefined, 0, 1, '=Ghost!A1'); // B1 → Ghost!A1
    expect(wb.getCellValue(undefined, 0, 1)).toBe('#REF!');
    // Editing the LOCAL A1 must not turn B1 into anything but #REF!.
    wb.setCellValue(undefined, 0, 0, 7);
    expect(wb.getCellValue(undefined, 0, 1)).toBe('#REF!');
  });
});

describe('number format decimalPlaces is clamped (no RangeError DoS)', () => {
  it('an out-of-range decimalPlaces does not throw in any formatter', () => {
    for (const type of ['number', 'percentage', 'scientific', 'currency'] as const) {
      expect(() => formatNumber(5, { type, decimalPlaces: 500 } as never)).not.toThrow();
      expect(() => formatNumber(5, { type, decimalPlaces: -3 } as never)).not.toThrow();
    }
  });
  it('a normal decimalPlaces is unaffected', () => {
    expect(formatNumber(1.5, { type: 'number', decimalPlaces: 2 } as never)).toBe('1.50');
  });
});
