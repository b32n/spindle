import { WorkbookImpl } from './workbook';

// A blank cell inside a range must read as empty (null), not 0 — otherwise every
// blank-aware aggregate is wrong. Formulas are placed outside the ranges they
// read so they don't self-reference.
describe('blank cells inside a range are empty, not 0', () => {
  function wb(): WorkbookImpl {
    const w = new WorkbookImpl('wb', 'WB');
    w.setCellValue(undefined, 0, 0, 10); // A1 = 10
    // A2, A3 left blank
    return w;
  }

  it('AVERAGE ignores blank cells', () => {
    const w = wb();
    w.setFormula(undefined, 0, 2, '=AVERAGE(A1:A3)'); // C1
    expect(w.getCellValue(undefined, 0, 2)).toBe(10);
  });

  it('COUNT / COUNTA count only populated cells', () => {
    const w = wb();
    w.setFormula(undefined, 0, 2, '=COUNT(A1:A3)');
    w.setFormula(undefined, 1, 2, '=COUNTA(A1:A3)');
    expect(w.getCellValue(undefined, 0, 2)).toBe(1);
    expect(w.getCellValue(undefined, 1, 2)).toBe(1);
  });

  it('MIN / MEDIAN skip blanks (not treated as 0)', () => {
    const w = wb();
    w.setFormula(undefined, 0, 2, '=MIN(A1:A3)');
    w.setFormula(undefined, 1, 2, '=MEDIAN(A1:A3)');
    expect(w.getCellValue(undefined, 0, 2)).toBe(10);
    expect(w.getCellValue(undefined, 1, 2)).toBe(10);
  });

  it('SUM is unaffected (blank coerces to 0 numerically)', () => {
    const w = wb();
    w.setFormula(undefined, 0, 2, '=SUM(A1:A3)');
    expect(w.getCellValue(undefined, 0, 2)).toBe(10);
  });

  it('a real 0 is still counted (distinct from blank)', () => {
    const w = new WorkbookImpl('wb', 'WB');
    w.setCellValue(undefined, 0, 0, 10);
    w.setCellValue(undefined, 1, 0, 0); // A2 = real 0
    w.setFormula(undefined, 0, 2, '=COUNT(A1:A3)'); // A1,A2 populated, A3 blank
    w.setFormula(undefined, 1, 2, '=COUNTBLANK(A1:A3)');
    expect(w.getCellValue(undefined, 0, 2)).toBe(2);
    expect(w.getCellValue(undefined, 1, 2)).toBe(1);
  });

  it('COUNTBLANK counts blank cells in a normal range', () => {
    const w = wb();
    w.setFormula(undefined, 0, 2, '=COUNTBLANK(A1:A3)');
    expect(w.getCellValue(undefined, 0, 2)).toBe(2);
  });
});

describe('ROWS/COLUMNS do not falsely self-reference', () => {
  it('=ROWS of a range containing the formula cell is not #CIRCULAR!', () => {
    const w = new WorkbookImpl('wb', 'WB');
    w.setFormula(undefined, 2, 0, '=ROWS(A1:A5)'); // A3 is inside A1:A5
    expect(w.getCellValue(undefined, 2, 0)).toBe(5);
  });

  it('=COLUMNS of a range containing the formula cell is not #CIRCULAR!', () => {
    const w = new WorkbookImpl('wb', 'WB');
    w.setFormula(undefined, 0, 2, '=COLUMNS(A1:E1)'); // C1 is inside A1:E1
    expect(w.getCellValue(undefined, 0, 2)).toBe(5);
  });

  it('a genuine value self-reference (=SUM inside its own range) is still circular (0, as in Excel)', () => {
    const w = new WorkbookImpl('wb', 'WB');
    w.setFormula(undefined, 2, 0, '=SUM(A1:A5)'); // A3 inside A1:A5, reads values
    expect(w.getCellValue(undefined, 2, 0)).toBe(0);
  });
});
