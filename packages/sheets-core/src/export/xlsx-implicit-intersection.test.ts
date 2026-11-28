import { withImplicitIntersection, withoutImplicitIntersection } from './xlsx-formulas';
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
    ['A1:A3*B1#', 'A1:A3*B1#'],
    ['A1:A3,B1', 'A1:A3,B1'],
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

  // KNOWN DIVERGENCE (documentation/sheets/TODO.md, Open items: unowned, no later plan entry): in a legacy formula Excel also intersects a range given to a scalar parameter
  // or used inside the arithmetic of a call's argument; the engine has no table of which parameters are scalar, so it leaves the calls alone
  // and returns a value where Excel gives the intersected cell or #VALUE!. The Excel column is from Excel's documented behaviour (and
  // a LibreOffice recalculation of the same formulas agrees with every row of the Excel column); these cases pin what the engine does today, so a fix shows up as a failing test to update.
  const divergent: Array<[string, string, unknown, unknown]> = [
    // at, formula, Excel, engine today
    ['F5', 'LEN(A1:A3)', '#VALUE!', 5],
    ['F5', 'IF(A1:A3>1,1,0)', '#VALUE!', 1],
    ['F2', 'LEN(A1:A3)', 1, 5],
    ['F2', 'IF(A1:A3>1,1,0)', 1, 1],
    ['F3', 'IF(A1:A3>1,1,0)', 1, 1],
    ['F1', 'IF(A1:A3>1,1,0)', 0, 1],
    ['F2', 'SUM(A1:A3*2)', 4, 12],
    ['F9', 'SUM(A1:A3*2)', '#VALUE!', 12],
    ['F2', 'ABS(A1:A3)', 2, '#VALUE!'], // the engine does not intersect a range for a scalar math parameter either: an error where Excel has the cell
    ['F5', 'ABS(A1:A3)', '#VALUE!', '#VALUE!'],
  ];
  it.each(divergent)('known divergence: %s =%s is %p in Excel, the engine gives %p', async (at, f, _excel, engine) => {
    expect(await valueAt(at, f)).toEqual(engine);
  });

  it('shows the `@` in the formula bar and writes the bare range back, as the file had it', async () => {
    const wb = new WorkbookImpl('wb', 'WB');
    await importFromXlsx(await xlsxWith(rows('F5', 'A1:A3*2')), wb);
    expect(wb.getCell(undefined, 4, 5)?.formula).toBe('=@A1:A3*2');
    // the saved result is kept (the default import does not recalculate)
    expect(wb.getCellValue(undefined, 4, 5)).toBe(0);
    expect(await formulaTexts(await exportToXlsx(wb))).toEqual({ F5: 'A1:A3*2' });
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

describe('withoutImplicitIntersection: the `@` the export drops', () => {
  const table: Array<[string, string]> = [
    ['@A1:A3*2', 'A1:A3*2'],
    ['@A1:A3', 'A1:A3'],
    ['-@A1:A3', '-A1:A3'],
    ['@A1:A3%', 'A1:A3%'],
    ['@A1:A3+@B1:B3', 'A1:A3+B1:B3'],
    ['@$A$1:$A$3*2', '$A$1:$A$3*2'],
    ['@A:A*2', 'A:A*2'],
    ['@1:1*2', '1:1*2'],
    ['@Sheet2!A1:A3*2', 'Sheet2!A1:A3*2'],
    ["@'S s'!A1:B2*2", "'S s'!A1:B2*2"],
    ['@A1:A3&"@A1:A3"', 'A1:A3&"@A1:A3"'],
    // an `@` before anything else stays, and so does one inside a call or a group (SUM(@A1:A3) is not SUM(A1:A3))
    ['@INDEX(A1:A3,2)*2', '@INDEX(A1:A3,2)*2'],
    ['@A1#*2', '@A1#*2'],
    ['@A1:A3*B1#', '@A1:A3*B1#'],
    ['@A1*2', '@A1*2'],
    ['@Name*2', '@Name*2'],
    ['@Table1[Col]*2', '@Table1[Col]*2'],
    ['SUM(@A1:A3)', 'SUM(@A1:A3)'],
    ['(@A1:A3)*2', '(@A1:A3)*2'],
    ['@A1:A3*2+@INDEX(B1:B3,1)', 'A1:A3*2+@INDEX(B1:B3,1)'],
    // nothing to drop
    ['A1:A3*2', 'A1:A3*2'],
    ['@A1:A3:B2', '@A1:A3:B2'],
    ['@A1:A3 B2:C2', '@A1:A3 B2:C2'],
    ['@Sheet1:Sheet3!A1:A3', '@Sheet1:Sheet3!A1:A3'],
    ['@A1:A3*"x', '@A1:A3*"x'],
    ['@A1+2', '@A1+2'],
    ['', ''],
  ];
  it.each(table)('%s -> %s', (input, expected) => {
    expect(withoutImplicitIntersection(input)).toBe(expected);
  });
});

/** The text of every `<f>` of the first sheet, by cell (XML entities undone). */
async function formulaTexts(file: Uint8Array): Promise<Record<string, string>> {
  const xml = await (await JSZip.loadAsync(file)).file('xl/worksheets/sheet1.xml')!.async('string');
  const out: Record<string, string> = {};
  for (const m of xml.matchAll(/<c r="([A-Z]+\d+)"[^>]*>(?:<f[^>]*>([^<]*)<\/f>|<f[^>]*\/>)/g)) {
    out[m[1]] = (m[2] ?? '').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
  }
  return out;
}

describe('import -> export -> import: the text of a legacy formula the importer touched comes back as the file had it', () => {
  const data = [ncell('A1', 1), ncell('B1', 10), ncell('C1', 20)];
  const base = (): Record<number, string[]> => ({ 1: [...data], 2: [ncell('A2', 2)], 3: [ncell('A3', 3)] });
  // the formulas of a legacy file (no `cm`, no array mark); each is exported as written
  const legacy = [
    'A1:A3*2', 'A1:A3', '2*A1:A3', 'A1:A3+B1:B3', 'A1:A3-B1:B3', '-A1:A3', 'A1:A3%', 'A1:A3^2', 'A1:A3*2+1', 'A1:A3&"x"', 'A1:A3&B1:B3',
    'A1:A3>1', 'A1:A3=A1:A3', 'A1:A3<>2', 'A1:A3 * 2', '$A$1:$A$3*2', 'A$1:A$3+$B1:$B3', 'A:A*2', '1:1*2', '$A:$A*2', 'Sheet2!A1:A3*2',
    "'My Sheet'!A1:A3*2", "'S s'!A1:B2*2", "Sheet2!A:A*2", 'Sheet2!1:1*2', 'A1:A3*"@A1:A3"', '"a@b"&A1:A3', 'A1*2', 'SUM(A1:A3)', 'SUM(A1:A3*2)',
    'IF(A1:A3>1,1,0)', '(A1:A3)*2', 'A1:A3*B1',
    // a SINGLE the file already had stays where it is not a bare range operand: before a call, inside a call
    '_xlfn.SINGLE(INDEX(A1:A3,2))*2', 'SUM(_xlfn.SINGLE(A1:A3))',
  ];
  // the one that is written differently: a SINGLE over a range operand of a formula without the array mark is the same as the bare range
  const normalised: Record<string, string> = { '_xlfn.SINGLE(A1:A3)+B1': 'A1:A3+B1' };

  it.each([...legacy, ...Object.keys(normalised)])('=%s', async (f) => {
    const wb = new WorkbookImpl('wb', 'WB');
    await importFromXlsx(await xlsxWith({ ...base(), 5: [fcell('F5', f, 0)] }), wb);
    const out = await exportToXlsx(wb);
    const expected = normalised[f] ?? f;
    expect((await formulaTexts(out)).F5).toBe(expected);
    // and a second round changes nothing
    const again = new WorkbookImpl('wb2', 'WB');
    await importFromXlsx(out, again);
    expect(again.getCell(undefined, 4, 5)?.formula).toBe(wb.getCell(undefined, 4, 5)?.formula);
    expect((await formulaTexts(await exportToXlsx(again))).F5).toBe(expected);
  });

  it('a shared formula: the master and the cells that share it come back as plain formulas, each as the file had it', async () => {
    const file = await xlsxWith({
      ...base(),
      2: [ncell('A2', 2), '<c r="G2"><f t="shared" ref="G2:G3" si="0">A1:A3</f><v>0</v></c>'],
      3: [ncell('A3', 3), '<c r="G3"><f t="shared" si="0"/><v>0</v></c>'],
      4: [ncell('A4', 4)],
    });
    const wb = new WorkbookImpl('wb', 'WB');
    await importFromXlsx(file, wb, { recalculateOnLoad: true });
    expect(wb.getCell(undefined, 1, 6)?.formula).toBe('=@A1:A3');
    expect(wb.getCell(undefined, 2, 6)?.formula).toBe('=@A2:A4');
    // G2 meets A2 (2) in A1:A3, G3 meets A3 (3) in A2:A4
    expect([wb.getCellCalculatedValue(undefined, 1, 6), wb.getCellCalculatedValue(undefined, 2, 6)]).toEqual([2, 3]);
    expect(await formulaTexts(await exportToXlsx(wb))).toMatchObject({ G2: 'A1:A3', G3: 'A2:A4' });
  });

  it('a legacy array formula (t="array") and a dynamic one (cm) keep their range operands bare, and the dynamic one its SINGLE', async () => {
    const file = await xlsxWith({
      ...base(),
      9: [fcell('F9', 'A1:A3*2', 2, 't="array" ref="F9:F11"')],
      10: [ncell('F10', 4)],
      11: [ncell('F11', 6)],
      13: [fcell('H13', '_xlfn.SINGLE(A1:A3)+A1:A3', 2, 't="array" ref="H13:H15"', ' cm="1"')],
      14: [ncell('H14', 3)],
      15: [ncell('H15', 5)],
    });
    const wb = new WorkbookImpl('wb', 'WB');
    await importFromXlsx(file, wb);
    expect(wb.getCell(undefined, 12, 7)?.formula).toBe('=@A1:A3+A1:A3');
    const texts = await formulaTexts(await exportToXlsx(wb));
    expect(texts.F9).toBe('A1:A3*2');
    expect(texts.H13).toBe('_xlfn.SINGLE(A1:A3)+A1:A3');
  });

  it('a formula the user types with `@` is written bare too (same meaning without the array mark); a spilling one keeps SINGLE', async () => {
    const wb = new WorkbookImpl('wb', 'WB');
    for (let r = 0; r < 3; r++) wb.setCellValue(undefined, r, 0, r + 1);
    wb.setFormula(undefined, 1, 5, '=@A1:A3*2');
    wb.setFormula(undefined, 4, 5, '=@INDEX(A1:A3,2)*2');
    wb.setFormula(undefined, 0, 7, '=@A1:A3+SEQUENCE(3)');
    const texts = await formulaTexts(await exportToXlsx(wb));
    expect(texts.F2).toBe('A1:A3*2');
    expect(texts.F5).toBe('_xlfn.SINGLE(INDEX(A1:A3,2))*2');
    expect(texts.H1).toBe('_xlfn.SINGLE(A1:A3)+_xlfn.SEQUENCE(3)');
  });

  it('a formula the engine cannot evaluate (cached-only) is written as the file had it, SINGLE included', async () => {
    const file = await xlsxWith({ ...base(), 5: [fcell('F5', '_xlfn.SINGLE(A1:A3)*_xlfn.NOSUCHFN(B1)', 7)] });
    const wb = new WorkbookImpl('wb', 'WB');
    await importFromXlsx(file, wb);
    expect(wb.getCell(undefined, 4, 5)?.cachedOnly).toBe(true);
    expect((await formulaTexts(await exportToXlsx(wb))).F5).toBe('_xlfn.SINGLE(A1:A3)*_xlfn.NOSUCHFN(B1)');
  });

  it('structural edits shift the references inside `@A1:A3*2`, and the export is still the bare range', async () => {
    const wb = new WorkbookImpl('wb', 'WB');
    await importFromXlsx(await xlsxWith({ ...base(), 5: [fcell('F5', 'A1:A3*2', 0)] }), wb, { recalculateOnLoad: true });
    const sheet = wb.getSheet();
    sheet.insertRows(0, 2); // the formula moves to row 7; A1:A3 becomes A3:A5
    expect(wb.getCell(undefined, 6, 5)?.formula).toBe('=@A3:A5*2');
    sheet.insertCols(0, 1); // and to column G; the range to B3:B5
    expect(wb.getCell(undefined, 6, 6)?.formula).toBe('=@B3:B5*2');
    sheet.deleteRows(0, 1);
    expect(wb.getCell(undefined, 5, 6)?.formula).toBe('=@B2:B4*2');
    expect(await formulaTexts(await exportToXlsx(wb))).toEqual({ G6: 'B2:B4*2' });
  });
});
