import { WorkbookImpl } from './workbook';

describe('cross-sheet dependency tracking', () => {
  it('recomputes a formula when the cell it references on another sheet changes', () => {
    const wb = new WorkbookImpl('wb', 'WB');
    const s2 = wb.addSheet('Sheet2');
    wb.setCellValue(s2.id, 0, 0, 10); // Sheet2!A1 = 10
    wb.setFormula(undefined, 0, 0, '=Sheet2!A1'); // Sheet1!A1
    expect(wb.getCellValue(undefined, 0, 0)).toBe(10);

    wb.setCellValue(s2.id, 0, 0, 20); // edit Sheet2!A1
    expect(wb.getCellValue(undefined, 0, 0)).toBe(20); // was stale (10) before this fix
  });

  it('recomputes a cross-sheet range aggregate when a cell inside it changes', () => {
    const wb = new WorkbookImpl('wb', 'WB');
    const s2 = wb.addSheet('Sheet2');
    wb.setCellValue(s2.id, 0, 0, 1);
    wb.setCellValue(s2.id, 1, 0, 2);
    wb.setCellValue(s2.id, 2, 0, 3);
    wb.setFormula(undefined, 0, 0, '=SUM(Sheet2!A1:A3)');
    expect(wb.getCellValue(undefined, 0, 0)).toBe(6);

    wb.setCellValue(s2.id, 1, 0, 10); // Sheet2!A2
    expect(wb.getCellValue(undefined, 0, 0)).toBe(14);
  });

  it('detects a circular reference that spans two sheets', () => {
    const wb = new WorkbookImpl('wb', 'WB');
    const s2 = wb.addSheet('Sheet2');
    wb.setFormula(undefined, 0, 0, '=Sheet2!A1'); // Sheet1!A1 -> Sheet2!A1
    wb.setFormula(s2.id, 0, 0, '=Sheet1!A1'); // Sheet2!A1 -> Sheet1!A1 (cycle)
    expect(wb.getCellValue(undefined, 0, 0)).toBe(0); // Excel: a circular reference evaluates to 0
    expect(wb.getCellValue(s2.id, 0, 0)).toBe(0);
  });

  it('recomputes to #REF! when a referenced sheet is deleted', () => {
    const wb = new WorkbookImpl('wb', 'WB');
    const s2 = wb.addSheet('Sheet2');
    wb.setCellValue(s2.id, 0, 0, 10);
    wb.setFormula(undefined, 0, 0, '=Sheet2!A1');
    expect(wb.getCellValue(undefined, 0, 0)).toBe(10);

    wb.deleteSheet(s2.id);
    expect(wb.getCellValue(undefined, 0, 0)).toBe('#REF!');
  });

  it('propagates #REF! through a chain of local formulas after a sheet delete (no stale cache)', () => {
    const wb = new WorkbookImpl('wb', 'WB');
    const s2 = wb.addSheet('Sheet2');
    wb.setCellValue(s2.id, 0, 0, 10);
    // A1 = B1, B1 = Sheet2!A1. A1 forward-references B1, which is the cell that
    // actually goes stale on delete.
    wb.setFormula(undefined, 0, 0, '=B1'); // Sheet1!A1
    wb.setFormula(undefined, 0, 1, '=Sheet2!A1'); // Sheet1!B1
    expect(wb.getCellValue(undefined, 0, 0)).toBe(10);
    expect(wb.getCellValue(undefined, 0, 1)).toBe(10);

    wb.deleteSheet(s2.id);
    // Both must recompute; A1 must NOT keep its pre-delete cached 10.
    expect(wb.getCellValue(undefined, 0, 1)).toBe('#REF!');
    expect(wb.getCellValue(undefined, 0, 0)).toBe('#REF!');
  });

  it('follows a renamed sheet instead of breaking the reference', () => {
    const wb = new WorkbookImpl('wb', 'WB');
    const s2 = wb.addSheet('Sheet2');
    wb.setCellValue(s2.id, 0, 0, 10);
    wb.setFormula(undefined, 0, 0, '=Sheet2!A1');
    expect(wb.getCellValue(undefined, 0, 0)).toBe(10);

    wb.renameSheet(s2.id, 'Renamed');
    wb.setCellValue(s2.id, 0, 0, 20); // edit via the renamed sheet
    expect(wb.getCellValue(undefined, 0, 0)).toBe(20); // reference followed the rename
  });

  it('does not falsely recompute for a same-index cell on another sheet', () => {
    const wb = new WorkbookImpl('wb', 'WB');
    const s2 = wb.addSheet('Sheet2');
    wb.setCellValue(s2.id, 0, 0, 1);
    wb.setCellValue(s2.id, 1, 0, 2);
    wb.setFormula(undefined, 0, 0, '=SUM(Sheet2!A1:A2)');
    expect(wb.getCellValue(undefined, 0, 0)).toBe(3);

    // Sheet1!A2 shares row/col indices with a cell inside the Sheet2 range but
    // must not dirty the aggregate — that would be a cross-sheet false positive.
    wb.setCellValue(undefined, 1, 0, 999);
    expect(wb.getCellValue(undefined, 0, 0)).toBe(3);
  });
});
