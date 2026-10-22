import { exportToCSV, importFromCSV } from './csv';
import { WorkbookImpl } from '../workbook';

function freshSheet() {
  const wb = new WorkbookImpl('wb', 'WB');
  return { wb, sheet: wb.getSheet() };
}

describe('CSV import', () => {
  it('grows the sheet so rows and columns beyond the 1000x100 default stay reachable', () => {
    const { sheet } = freshSheet();
    const rows = Array.from({ length: 1500 }, (_, i) => `${i},x`);
    importFromCSV(rows.join('\n') + '\n' + ','.repeat(119) + 'z', sheet);
    expect(sheet.rowCount).toBe(1501);
    expect(sheet.colCount).toBe(120);
  });

  it('handles CRLF line endings without a trailing \\r on the last column', () => {
    const { wb, sheet } = freshSheet();
    importFromCSV('a,b\r\nc,d\r\n', sheet);
    expect(wb.getCellValue(undefined, 0, 1)).toBe('b'); // not "b\r"
    expect(wb.getCellValue(undefined, 1, 1)).toBe('d');
  });

  it('keeps a quoted field that contains a newline as one cell', () => {
    const { wb, sheet } = freshSheet();
    importFromCSV('"line1\nline2",z\n', sheet);
    expect(wb.getCellValue(undefined, 0, 0)).toBe('line1\nline2');
    expect(wb.getCellValue(undefined, 0, 1)).toBe('z');
    // No phantom extra row from the embedded newline.
    expect(wb.getCellValue(undefined, 1, 0)).toBeNull();
  });

  it('handles quoted commas and escaped quotes', () => {
    const { wb, sheet } = freshSheet();
    importFromCSV('"a,b","c""d"\n', sheet);
    expect(wb.getCellValue(undefined, 0, 0)).toBe('a,b');
    expect(wb.getCellValue(undefined, 0, 1)).toBe('c"d');
  });

  it('preserves a whitespace-only field between commas', () => {
    const { wb, sheet } = freshSheet();
    importFromCSV('a, ,c\n', sheet);
    expect(wb.getCellValue(undefined, 0, 1)).toBe(' ');
    expect(wb.getCellValue(undefined, 0, 2)).toBe('c');
  });
});

describe('CSV export → import round-trips', () => {
  it('round-trips a cell containing commas and newlines', () => {
    const { wb } = freshSheet();
    wb.setCellValue(undefined, 0, 0, 'x\ny');
    wb.setCellValue(undefined, 0, 1, 'a,b');
    wb.setCellValue(undefined, 1, 0, 'plain');
    const csv = exportToCSV(wb);

    const { wb: wb2, sheet: sheet2 } = freshSheet();
    importFromCSV(csv, sheet2);
    expect(wb2.getCellValue(undefined, 0, 0)).toBe('x\ny');
    expect(wb2.getCellValue(undefined, 0, 1)).toBe('a,b');
    expect(wb2.getCellValue(undefined, 1, 0)).toBe('plain');
  });
});

describe('CSV export of formulas (#77)', () => {
  it('writes computed values, not formula text', () => {
    const { wb } = freshSheet();
    wb.setCellValue(undefined, 0, 0, 10);
    wb.setCellValue(undefined, 0, 1, 2.5);
    wb.setFormula(undefined, 0, 2, '=A1*B1');
    wb.setFormula(undefined, 0, 3, '=A1&"x"');
    wb.setFormula(undefined, 0, 4, '=A1>5');
    wb.setFormula(undefined, 0, 5, '=1/0');
    expect(exportToCSV(wb).trim()).toBe('10,2.5,25,10x,true,#DIV/0!');
  });

  it('includes spilled cells and round-trips as numbers', () => {
    const { wb } = freshSheet();
    wb.setFormula(undefined, 0, 0, '=SEQUENCE(3)');
    const csv = exportToCSV(wb);
    expect(csv.trim().split(/\r?\n/)).toEqual(['1', '2', '3']);
    const { wb: wb2, sheet } = freshSheet();
    importFromCSV(csv, sheet);
    expect(wb2.getCellValue(undefined, 2, 0)).toBe(3);
  });
});
