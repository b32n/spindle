import { withImplicitIntersection } from './xlsx-formulas';
import JSZip from 'jszip';
import { exportToXlsx, importFromXlsx } from './xlsx';
import { fcell, ncell, xlsxWith } from './test-xlsx';
import { WorkbookImpl } from '../workbook';

// SA-gap: a formula a file saved without the array mark is a legacy one, and Excel intersects a range it uses as an operand with the
// formula's own row (a column range) or column (a row range): the cell, or #VALUE! when the formula is outside the range. Excel 365 shows
// that as `@` ("Implicit intersection operator: @", Microsoft Support). The tables below are written from that rule, and the values of
// the import table were also cross-checked by recalculating the same formulas in LibreOffice 25.2.3.2 (an openpyxl file, no saved results,
// converted to CSV): every value agrees. They are not produced by the engine.

describe('withImplicitIntersection: where the `@` goes', () => {
  const table: Array<[string, string]> = [
    ['A1:A3*2', '@A1:A3*2'],
    ['A1:A3', '@A1:A3'],
    ['2*A1:A3', '2*@A1:A3'],
    ['A1:A3+B1:B3', '@A1:A3+@B1:B3'],
    ['-A1:A3', '-@A1:A3'],
    ['A1:A3%', '@A1:A3%'],
    ['A1:A3&"x"', '@A1:A3&"x"'],
    ['A1:A3>1', '@A1:A3>1'],
    ['$A$1:$A$3*2', '@$A$1:$A$3*2'],
    ['A:A*2', '@A:A*2'],
    ['1:1*2', '@1:1*2'],
    ['Sheet2!A1:A3*2', '@Sheet2!A1:A3*2'],
    ["'My Sheet'!A1:A3*2", "@'My Sheet'!A1:A3*2"],
    ['A1:A3 * 2', '@A1:A3 * 2'],
    // one cell, or no range: nothing to intersect
    ['A1*2', 'A1*2'],
    ['A1:A1*2', 'A1:A1*2'],
    ['1+2', '1+2'],
    ['"A1:A3"', '"A1:A3"'],
    // a call's arguments, a group and an array constant are left alone (what is done with a range depends on the function)
    ['SUM(A1:A3)', 'SUM(A1:A3)'],
    ['SUM(A1:A3*2)', 'SUM(A1:A3*2)'],
    ['SUMPRODUCT((A1:A10>3)*1)', 'SUMPRODUCT((A1:A10>3)*1)'],
    ['(A1:A3)*2', '(A1:A3)*2'],
    ['SUM({1,2;3,4})+A1:A3', 'SUM({1,2;3,4})+@A1:A3'],
    // already explicit, a spill reference, a union, an intersection, a 3D reference, a range of ranges
    ['@A1:A3*2', '@A1:A3*2'],
    ['A1#*2', 'A1#*2'],
    ['(A1:A2,C1)', '(A1:A2,C1)'],
    ['A1:A3 B2:C2', 'A1:A3 B2:C2'],
    ['Sheet1:Sheet3!A1:A3*2', 'Sheet1:Sheet3!A1:A3*2'],
    ['A1:A3:B2*2', 'A1:A3:B2*2'],
    ['A1:A3*B1:C1:D1', 'A1:A3*B1:C1:D1'],
    // text the tokenizer rejects comes back as it was
    ['A1:A3*"x', 'A1:A3*"x'],
    ['', ''],
  ];
  it.each(table)('%s -> %s', (input, expected) => {
    expect(withImplicitIntersection(input)).toBe(expected);
  });

  it('is idempotent', () => {
    for (const [, out] of table) expect(withImplicitIntersection(out)).toBe(out);
  });
});

describe('a legacy formula over a multi-cell range is intersected with its own row or column on load', () => {
  // data: A1:A3 = 1, 2, 3 and B1:D1 = 10, 20, 30; the formula sits at `at`
  const data = [ncell('A1', 1), ncell('B1', 10), ncell('C1', 20), ncell('D1', 30)];
  const rows = (at: string, f: string, attrs = ''): Record<number, string[]> => {
    const row = Number(/\d+/.exec(at)?.[0]);
    const cells: Record<number, string[]> = { 1: [...data], 2: [ncell('A2', 2)], 3: [ncell('A3', 3)] };
    (cells[row] ??= []).push(fcell(at, f, 0, attrs));
    return cells;
  };
  async function valueAt(at: string, f: string, attrs = ''): Promise<unknown> {
    const wb = new WorkbookImpl('wb', 'WB');
    await importFromXlsx(await xlsxWith(rows(at, f, attrs)), wb, { recalculateOnLoad: true });
    const [, letters, digits] = /^([A-Z]+)(\d+)$/.exec(at) as RegExpExecArray;
    let col = 0;
    for (const ch of letters) col = col * 26 + ch.charCodeAt(0) - 64;
    return wb.getCellValue(undefined, Number(digits) - 1, col - 1);
  }

  const table: Array<[string, string, unknown]> = [
    // a column range takes the cell in the formula's row
    ['F2', 'A1:A3*2', 4],
    ['F3', 'A1:A3*2', 6],
    ['F1', 'A1:A3&"x"', '1x'],
    ['F3', 'A1:A3', 3],
    ['F2', '-A1:A3', -2],
    ['F2', 'A:A*2', 4],
    // outside the range's rows: #VALUE!
    ['F5', 'A1:A3*2', '#VALUE!'],
    ['F5', 'A1:A3', '#VALUE!'],
    ['F7', 'A1:A3%', '#VALUE!'],
    // a row range takes the cell in the formula's column
    ['C5', 'B1:D1*2', 40],
    ['D5', 'B1:D1', 30],
    ['E5', 'B1:D1*2', '#VALUE!'],
    ['A5', 'B1:D1*2', '#VALUE!'],
    ['C5', '1:1*2', 40],
    // a single cell, and a call's arguments, are as they were
    ['F9', 'A1:A1*2', 2],
    ['F9', 'SUM(A1:A3*2)', 12],
    ['F9', 'SUMPRODUCT(A1:A3,A1:A3)', 14],
  ];
  it.each(table)('%s: =%s gives %p', async (at, f, expected) => {
    expect(await valueAt(at, f)).toEqual(expected);
  });

  it('shows the `@` in the formula bar and writes it back as SINGLE, which Excel reads as the same formula', async () => {
    const wb = new WorkbookImpl('wb', 'WB');
    await importFromXlsx(await xlsxWith(rows('F5', 'A1:A3*2')), wb);
    expect(wb.getCell(undefined, 4, 5)?.formula).toBe('=@A1:A3*2');
    // the saved result is kept (the default import does not recalculate)
    expect(wb.getCellValue(undefined, 4, 5)).toBe(0);
    const xml = await (await JSZip.loadAsync(await exportToXlsx(wb))).file('xl/worksheets/sheet1.xml')!.async('string');
    expect(xml).toMatch(/<f>_xlfn\.SINGLE\(A1:A3\)\*2<\/f>/);
  });

  it('leaves an array formula (legacy or dynamic) as a range operation', async () => {
    const legacy = await xlsxWith({ 1: [ncell('A1', 1)], 2: [ncell('A2', 2)], 3: [ncell('A3', 3)], 9: [fcell('F9', 'A1:A3*2', 2, 't="array" ref="F9:F11"')], 10: [ncell('F10', 4)], 11: [ncell('F11', 6)] });
    const wb = new WorkbookImpl('wb', 'WB');
    await importFromXlsx(legacy, wb, { recalculateOnLoad: true });
    expect(wb.getCell(undefined, 8, 5)?.formula).toBe('=A1:A3*2');
    expect([8, 9, 10].map((r) => wb.getCellCalculatedValue(undefined, r, 5))).toEqual([2, 4, 6]);
  });

  it('a formula typed in the editor still spills', () => {
    const wb = new WorkbookImpl('wb', 'WB');
    wb.setCellValue(undefined, 0, 0, 1);
    wb.setCellValue(undefined, 1, 0, 2);
    wb.setFormula(undefined, 0, 2, '=A1:A2*2');
    expect([0, 1].map((r) => wb.getCellCalculatedValue(undefined, r, 2))).toEqual([2, 4]);
  });
});
