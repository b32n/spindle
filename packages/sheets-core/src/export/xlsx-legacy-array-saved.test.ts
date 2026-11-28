import JSZip from 'jszip';
import { exportToXlsx, importFromXlsx } from './xlsx';
import { fcell, ncell, xlsxWith } from './test-xlsx';
import { WorkbookImpl } from '../workbook';
import type { Cell } from '../types';

// S4 (SD4): a legacy array formula whose result the engine cannot reproduce keeps the file's saved results, whatever the engine says.

async function open(bytes: Uint8Array) {
  const wb = new WorkbookImpl('wb', 'WB');
  const result = await importFromXlsx(bytes, wb);
  return { wb, ...result };
}
const cellAt = (wb: WorkbookImpl, row: number, col: number): Cell | undefined => wb.getCell(undefined, row, col);
const values = (wb: WorkbookImpl, rows: number, cols: number, row = 0, col = 0) =>
  Array.from({ length: rows }, (_, r) => Array.from({ length: cols }, (_, c) => wb.getCellValue(undefined, row + r, col + c)));
const array = (ref: string) => `t="array" ref="${ref}"`;
const xmlOf = async (bytes: Uint8Array) => (await JSZip.loadAsync(bytes)).file('xl/worksheets/sheet1.xml')!.async('string');
const cachedOnlyItem = (loss: { items: Array<{ id: string; label: string; count: number; where?: string[]; effect: string }> }) => loss.items.find((i) => i.id === 'xlsx.formulas.cachedOnly');

describe('a legacy array formula the engine evaluates only as a scalar keeps the saved results', () => {
  it('{=ROW(A1:A3)} over A1:A3 with saved 1, 2, 3 stays 1, 2, 3 (not 1, 1, 1), cached-only, reported', async () => {
    const { wb, loss } = await open(await xlsxWith({ 1: [fcell('A1', 'ROW(A1:A3)', 1, array('A1:A3'))], 2: [ncell('A2', 2)], 3: [ncell('A3', 3)] }));
    expect(values(wb, 3, 1)).toEqual([[1], [2], [3]]);
    expect(cellAt(wb, 0, 0)).toMatchObject({ formula: '=ROW(A1:A3)', value: 1, cachedOnly: true, arraySpan: { rows: 3, cols: 1 } });
    expect(cellAt(wb, 0, 0)?.arraySpan?.spill).toBeUndefined();
    expect(wb.getSpillRegion(undefined, 0, 0)).toBeUndefined();
    expect(cachedOnlyItem(loss)).toMatchObject({ count: 1, where: ['Sheet1'], state: 'kept' });
    expect(cachedOnlyItem(loss)?.effect).toMatch(/^kept: shown from the value saved in the file; the saved value and the formula are kept and written back, and the formula's cell references follow row and column insert and delete/);
  });

  it('writes the saved results back, and a second import reads the same', async () => {
    const { wb } = await open(await xlsxWith({ 1: [fcell('A1', 'ROW(A1:A3)', 1, array('A1:A3'))], 2: [ncell('A2', 2)], 3: [ncell('A3', 3)] }));
    const out = await exportToXlsx(wb);
    const xml = await xmlOf(out);
    expect(xml).toMatch(/<c r="A1"><f t="array" ref="A1:A3">ROW\(A1:A3\)<\/f><v>1<\/v><\/c>/);
    expect(xml).toMatch(/<c r="A2"><v>2<\/v><\/c>/);
    expect(xml).toMatch(/<c r="A3"><v>3<\/v><\/c>/);
    expect(xml).not.toContain('cm=');
    const again = await open(out);
    expect(values(again.wb, 3, 1)).toEqual([[1], [2], [3]]);
    expect(cellAt(again.wb, 0, 0)).toMatchObject({ cachedOnly: true, arraySpan: { rows: 3, cols: 1 } });
    expect(await xmlOf(await exportToXlsx(again.wb))).toBe(xml);
  });

  it('does not recalculate it when a precedent changes (the saved values are all there is)', async () => {
    const { wb } = await open(await xlsxWith({ 1: [ncell('A1', 5), fcell('C1', 'ROW(A1:A3)', 1, array('C1:C3'))], 2: [ncell('C2', 2)], 3: [ncell('C3', 3)] }));
    wb.setCellValue(undefined, 0, 0, 6);
    expect(values(wb, 3, 1, 0, 2)).toEqual([[1], [2], [3]]);
  });

  it('{=COLUMN(A1:C1)} over B2:D2 keeps its saved 1, 2, 3', async () => {
    const { wb, loss } = await open(await xlsxWith({ 2: [fcell('B2', 'COLUMN(A1:C1)', 1, array('B2:D2')), ncell('C2', 2), ncell('D2', 3)] }));
    expect(values(wb, 1, 3, 1, 1)).toEqual([[1, 2, 3]]);
    expect(cellAt(wb, 1, 1)?.cachedOnly).toBe(true);
    expect(cachedOnlyItem(loss)?.count).toBe(1);
  });

  it('counts every such array in the report, per sheet name (ROW and COLUMN)', async () => {
    const { loss } = await open(
      await xlsxWith({
        1: [fcell('A1', 'ROW(A1:A3)', 1, array('A1:A3')), fcell('C1', 'COLUMN(C1:E1)', 3, array('C1:E1')), ncell('D1', 4), ncell('E1', 5)],
        2: [ncell('A2', 2)],
        3: [ncell('A3', 3)],
      })
    );
    // COLUMN(C1:E1) is 3, 4, 5 saved: the engine's 3 repeated is not that
    expect(cachedOnlyItem(loss)).toMatchObject({ count: 2, where: ['Sheet1'] });
  });
});

describe('a legacy array formula the engine evaluates to the saved values stays evaluated', () => {
  it('ROWS(A1:A3) over C1:C3 saved 3, 3, 3 is the engine scalar filled across the range, and recalculates', async () => {
    const { wb, loss } = await open(await xlsxWith({ 1: [fcell('C1', 'ROWS(A1:A3)', 3, array('C1:C3'))], 2: [ncell('C2', 3)], 3: [ncell('C3', 3)] }));
    expect(cellAt(wb, 0, 2)?.cachedOnly).toBeUndefined();
    expect(wb.getSpillRegion(undefined, 0, 2)).toMatchObject({ rows: 3, cols: 1, values: [[3], [3], [3]] });
    expect(cachedOnlyItem(loss)).toBeUndefined();
  });

  it('{=A1:A3*2} over C1:C3 follows its precedents', async () => {
    const { wb, loss } = await open(await xlsxWith({ 1: [ncell('A1', 1), fcell('C1', 'A1:A3*2', 2, array('C1:C3'))], 2: [ncell('A2', 2), ncell('C2', 4)], 3: [ncell('A3', 3), ncell('C3', 6)] }));
    expect(cellAt(wb, 0, 2)?.cachedOnly).toBeUndefined();
    expect(cachedOnlyItem(loss)).toBeUndefined();
    wb.setCellValue(undefined, 0, 0, 10);
    expect(wb.getSpillRegion(undefined, 0, 2)?.values).toEqual([[20], [4], [6]]);
  });

  it('TRANSPOSE of a column over a row, with the saved values', async () => {
    const { wb, loss } = await open(
      await xlsxWith({
        1: [ncell('A1', 5), ncell('B1', 1), ncell('C1', 2), fcell('E1', 'TRANSPOSE(A1:A3)', 5, array('E1:G1')), ncell('F1', 6), ncell('G1', 7)],
        2: [ncell('A2', 6), ncell('B2', 3), ncell('C2', 4)],
        3: [ncell('A3', 7)],
      })
    );
    expect(wb.getSpillRegion(undefined, 0, 4)?.values).toEqual([[5, 6, 7]]);
    expect(cellAt(wb, 0, 4)?.cachedOnly).toBeUndefined();
    expect(cachedOnlyItem(loss)).toBeUndefined();
  });

  it('keeps the saved 0 where the engine pads a result shorter than the range with #N/A (A1:A2 over C1:C3)', async () => {
    const { wb, loss } = await open(await xlsxWith({ 1: [ncell('A1', 4), fcell('C1', 'A1:A2', 4, array('C1:C3'))], 2: [ncell('A2', 5), ncell('C2', 5)], 3: [ncell('C3', 0)] }));
    expect(values(wb, 3, 1, 0, 2)).toEqual([[4], [5], [0]]);
    expect(cellAt(wb, 0, 2)?.cachedOnly).toBe(true);
    expect(cachedOnlyItem(loss)?.count).toBe(1);
  });

  it('MMULT, which the engine does not have, keeps the saved 2 x 2 as it always did (cached-only, counted once)', async () => {
    const { wb, loss } = await open(
      await xlsxWith({
        1: [ncell('A1', 1), ncell('B1', 2), fcell('D1', 'MMULT(A1:B2,A1:B2)', 7, array('D1:E2')), ncell('E1', 10)],
        2: [ncell('A2', 3), ncell('B2', 4), ncell('D2', 15), ncell('E2', 22)],
      })
    );
    expect(values(wb, 2, 2, 0, 3)).toEqual([[7, 10], [15, 22]]);
    expect(cellAt(wb, 0, 3)?.cachedOnly).toBe(true);
    expect(cachedOnlyItem(loss)?.count).toBe(1);
  });
});

describe('a legacy array formula whose saved results differ in one member', () => {
  const mismatched = () => xlsxWith({ 1: [ncell('A1', 1), fcell('C1', 'A1:A3*2', 2, array('C1:C3'))], 2: [ncell('A2', 2), ncell('C2', 4)], 3: [ncell('A3', 3), ncell('C3', 7)] });

  it('keeps all the saved results (the 7 as well) as cached-only and reports it', async () => {
    const { wb, loss } = await open(await mismatched());
    expect(values(wb, 3, 1, 0, 2)).toEqual([[2], [4], [7]]);
    expect(cellAt(wb, 0, 2)).toMatchObject({ formula: '=A1:A3*2', cachedOnly: true, arraySpan: { rows: 3, cols: 1 } });
    expect(cachedOnlyItem(loss)).toMatchObject({ count: 1, where: ['Sheet1'] });
    wb.setCellValue(undefined, 0, 0, 10);
    expect(values(wb, 3, 1, 0, 2)).toEqual([[2], [4], [7]]);
  });

  it('writes it back as the file had it', async () => {
    const xml = await xmlOf(await exportToXlsx((await open(await mismatched())).wb));
    expect(xml).toMatch(/<c r="C1"><f t="array" ref="C1:C3">A1:A3\*2<\/f><v>2<\/v><\/c>/);
    expect(xml).toMatch(/<c r="C3"><v>7<\/v><\/c>/);
  });

  it('joins the detector\'s cached-only item: counted next to a formula the engine does not have', async () => {
    const bytes = await xlsxWith({
      1: [ncell('A1', 1), fcell('C1', 'A1:A3*2', 2, array('C1:C3')), fcell('E1', '_xlfn.NOSUCHFUNCTION(A1)', 1)],
      2: [ncell('C2', 4)],
      3: [ncell('C3', 7)],
    });
    const { loss } = await open(bytes);
    expect(cachedOnlyItem(loss)).toMatchObject({ count: 2, where: ['Sheet1'] });
    // alone, the array makes the item itself, in the detector's words
    const alone = await open(await xlsxWith({ 1: [fcell('E1', '_xlfn.NOSUCHFUNCTION(A1)', 1)] }));
    const lone = await open(await mismatched());
    expect(cachedOnlyItem(lone.loss)).toMatchObject({ count: 1, label: cachedOnlyItem(alone.loss)?.label, effect: cachedOnlyItem(alone.loss)?.effect, state: 'kept' });
  });
});

describe('a legacy array formula whose saved results are errors', () => {
  const errorCell = (ref: string, f?: string, ref2?: string) =>
    `<c r="${ref}" t="e">${f ? `<f t="array" ref="${ref2}">${f}</f>` : ''}<v>#DIV/0!</v></c>`;
  const allErrors = (f: string) => xlsxWith({ 1: [errorCell('A1', f, 'A1:B2'), errorCell('B1')], 2: [errorCell('A2'), errorCell('B2')] });

  it('keeps four saved errors when the engine computes numbers there', async () => {
    const { wb, loss } = await open(await allErrors('1+1'));
    expect(values(wb, 2, 2)).toEqual([['#DIV/0!', '#DIV/0!'], ['#DIV/0!', '#DIV/0!']]);
    expect(cellAt(wb, 0, 0)).toMatchObject({ cachedOnly: true, formula: '=1+1' });
    expect(cachedOnlyItem(loss)?.count).toBe(1);
  });

  it('keeps four saved errors when the engine has an error and no array at all', async () => {
    const { wb, loss } = await open(await allErrors('1/0'));
    expect(wb.getSpillRegion(undefined, 0, 0)).toBeUndefined();
    expect(cellAt(wb, 0, 0)).toMatchObject({ cachedOnly: true, value: '#DIV/0!' });
    expect(values(wb, 2, 2)).toEqual([['#DIV/0!', '#DIV/0!'], ['#DIV/0!', '#DIV/0!']]);
    expect(cachedOnlyItem(loss)?.count).toBe(1);
  });

  it.each(['1+1', '1/0', 'SEQUENCE(2,2)/0'])('writes {=%s} over A1:B2 back with all four error cells, members included', async (formula) => {
    const xml = await xmlOf(await exportToXlsx((await open(await allErrors(formula))).wb));
    expect(xml).toMatch(/<c r="A1" t="e"><f t="array" ref="A1:B2">[^<]+<\/f><v>#DIV\/0!<\/v><\/c>/);
    for (const ref of ['B1', 'A2', 'B2']) expect(xml).toContain(`<c r="${ref}" t="e"><v>#DIV/0!</v></c>`);
  });

  it('finds the saved error cells of a range over millions of cells by the cells there are (budget-cut, kept as saved)', async () => {
    const bytes = await xlsxWith({ 1: [errorCell('A1', '1+1', 'A1:J400000'), errorCell('B1')], 2: [ncell('A2', 5), errorCell('B2')] });
    const { wb, loss } = await open(bytes);
    expect(cellAt(wb, 0, 0)).toMatchObject({ cachedOnly: true, arraySpan: { rows: 400000, cols: 10 } });
    expect(loss.items.find((i) => i.id === 'xlsx.package.sizeLimit')).toMatchObject({ count: 1 });
    const xml = await xmlOf(await exportToXlsx(wb));
    expect(xml).toContain('<c r="B1" t="e"><v>#DIV/0!</v></c>');
    expect(xml).toContain('<c r="B2" t="e"><v>#DIV/0!</v></c>');
    expect(xml).toContain('<c r="A2"><v>5</v></c>');
  });

  it('leaves the saved errors of a range as the text of the error once the export has used its budget for finding them', async () => {
    const xml = await xmlOf(await exportToXlsx((await open(await allErrors('1+1'))).wb, { maxTotalArrayCells: 2 }));
    expect(xml).toContain('<c r="A1" t="e">');
    expect(xml).not.toContain('<c r="B1" t="e">');
  });

  it('keeps an array of errors the engine reproduces as evaluated', async () => {
    const { wb, loss } = await open(await allErrors('SEQUENCE(2,2)/0'));
    expect(cellAt(wb, 0, 0)?.cachedOnly).toBeUndefined();
    expect(cachedOnlyItem(loss)).toBeUndefined();
    expect(wb.getSpillRegion(undefined, 0, 0)).toMatchObject({ rows: 2, cols: 2 });
  });
});

describe('what the export writes of array results', () => {
  it.each(['SEQUENCE(2)/0', 'SEQUENCE(2)*1E308*10', '(SEQUENCE(2)*1E308*10)-(SEQUENCE(2)*1E308*10)'])('writes the non-numbers of =%s as error cells', async (formula) => {
    const wb = new WorkbookImpl('wb', 'WB');
    wb.setFormula(undefined, 0, 0, `=${formula}`);
    expect(await xmlOf(await exportToXlsx(wb))).toMatch(/<c r="A2" t="e"><v>#(DIV\/0|NUM)!<\/v><\/c>/);
  });

  const sheetOf = async (wb: WorkbookImpl, options: Parameters<typeof exportToXlsx>[1] = {}) => xmlOf(await exportToXlsx(wb, options));

  it('writes the errors of a spilled result as error cells (SEQUENCE(2)/0), not as text', async () => {
    const wb = new WorkbookImpl('wb', 'WB');
    wb.setFormula(undefined, 0, 0, '=SEQUENCE(2)/0');
    const xml = await sheetOf(wb);
    expect(xml).toContain('<c r="A2" t="e"><v>#DIV/0!</v></c>');
    expect(xml).toMatch(/<c r="A1"[^>]* t="e">/);
  });

  it('never prints a negative number of cells left in the warning (budget 3, six SEQUENCE(3))', async () => {
    const wb = new WorkbookImpl('wb', 'WB');
    for (let i = 0; i < 6; i++) wb.setFormula(undefined, 0, i * 2, '=SEQUENCE(3)');
    const warnings: string[] = [];
    await exportToXlsx(wb, { maxTotalArrayCells: 3, onWarning: (m) => warnings.push(m) });
    expect(warnings).toHaveLength(5);
    for (const message of warnings) expect(message).toContain('does not fit in the 0 cells left of the workbook');
    expect(warnings.join('\n')).not.toMatch(/-\d/);
  });

  it('cuts in sheet position, not in the order the cells were entered (E1, C1, A1 with a budget of 4: A1 stays whole)', async () => {
    const wb = new WorkbookImpl('wb', 'WB');
    wb.setFormula(undefined, 0, 4, '=SEQUENCE(3)');
    wb.setFormula(undefined, 0, 2, '=SEQUENCE(3)');
    wb.setFormula(undefined, 0, 0, '=SEQUENCE(3)');
    const warnings: string[] = [];
    const xml = await sheetOf(wb, { maxTotalArrayCells: 4, onWarning: (m) => warnings.push(m) });
    expect(xml).toContain('ref="A1:A3"');
    // each cut array keeps one cell, the anchor, even with nothing left of the budget
    expect(xml).toMatch(/<f t="array" ref="C1">/);
    expect(xml).toMatch(/<f t="array" ref="E1">/);
    expect(warnings).toEqual([
      "Sheet1!C1: the array formula's result (3 x 1 cells) does not fit in the 1 cells left of the workbook's budget for array formula results, so only its first 1 x 1 are written to the file.",
      "Sheet1!E1: the array formula's result (3 x 1 cells) does not fit in the 0 cells left of the workbook's budget for array formula results, so only its first 1 x 1 are written to the file.",
    ]);
  });

  it('takes the sheets in order and then the rows: a formula in row 2 of the first sheet comes before row 1 of the second', async () => {
    const wb = new WorkbookImpl('wb', 'WB');
    const second = wb.addSheet('Two');
    wb.setFormula(second.id, 0, 0, '=SEQUENCE(3)');
    wb.setFormula(undefined, 4, 0, '=SEQUENCE(3)');
    wb.setFormula(undefined, 0, 0, '=SEQUENCE(3)');
    const warnings: string[] = [];
    await exportToXlsx(wb, { maxTotalArrayCells: 7, onWarning: (m) => warnings.push(m) });
    expect(warnings).toHaveLength(1);
    expect(warnings[0]).toMatch(/^Two!A1: .*does not fit in the 1 cells left/);
  });
});

// what the comparison decides, case by case: the engine's result for `formula` over C1:C3 against the saved members (the data-preservation decision)
describe('the comparison of the engine result with the saved members (oracle table)', () => {
  const sc = (ref: string, v: string) => `<c r="${ref}" t="str"><v>${v}</v></c>`;
  const bc = (ref: string, v: 0 | 1) => `<c r="${ref}" t="b"><v>${v}</v></c>`;
  const ec = (ref: string, code: string) => `<c r="${ref}" t="e"><v>${code}</v></c>`;
  // A1:A3 hold 1, (blank), 3; B2 is there so that row 2 exists (the engine pads a range over a row without cells with #N/A)
  const over = (anchor: string, c2: string, c3: string) => xlsxWith({ 1: [ncell('A1', 1), anchor], 2: [ncell('B2', 9), c2], 3: [ncell('A3', 3), c3] });
  const table: Array<[string, () => Promise<Uint8Array>, boolean]> = [
    ['blank source, saved 0', () => over(fcell('C1', 'A1:A3', 1, array('C1:C3')), ncell('C2', 0), ncell('C3', 3)), false],
    ['blank source (the engine\'s 0), saved empty t="str" (not the same result)', () => over(fcell('C1', 'A1:A3', 1, array('C1:C3')), sc('C2', ''), ncell('C3', 3)), true],
    ['blank source (the engine\'s 0), saved t="b" false (not the same result)', () => over(fcell('C1', 'A1:A3', 1, array('C1:C3')), bc('C2', 0), ncell('C3', 3)), true],
    ['blank source, saved 5 (not blank)', () => over(fcell('C1', 'A1:A3', 1, array('C1:C3')), ncell('C2', 5), ncell('C3', 3)), true],
    ['blank source, no saved cell at all', () => xlsxWith({ 1: [ncell('A1', 1), fcell('C1', 'A1:A3', 1, array('C1:C3'))], 2: [ncell('B2', 9)], 3: [ncell('A3', 3), ncell('C3', 3)] }), false],
    ['saved 0 where the engine has 7', () => over(fcell('C1', 'A1:A3*0+7', 7, array('C1:C3')), ncell('C2', 0), ncell('C3', 7)), true],
    ['saved t="str" members equal to the source text', () => xlsxWith({ 1: [`<c r="A1" t="str"><v>x</v></c>`, fcell('C1', 'A1:A2', 'x', array('C1:C2'))], 2: [`<c r="A2" t="str"><v>y</v></c>`, sc('C2', 'y')] }), false],
    ['saved t="str" member that differs in case', () => xlsxWith({ 1: [`<c r="A1" t="str"><v>x</v></c>`, fcell('C1', 'A1:A2', 'x', array('C1:C2'))], 2: [`<c r="A2" t="str"><v>y</v></c>`, sc('C2', 'Y')] }), true],
    ['saved t="b" members equal to the source booleans', () => xlsxWith({ 1: [bc('A1', 1), `<c r="C1" t="b"><f ${array('C1:C2')}>A1:A2</f><v>1</v></c>`], 2: [bc('A2', 0), bc('C2', 0)] }), false],
    ['float noise on a member (0.1+0.2 saved as 0.3)', () => xlsxWith({ 1: [fcell('C1', 'A1:A2*1', 1, array('C1:C2')), ncell('A1', 1), ncell('A2', 0.1)], 2: [ncell('C2', 0.1 + 0.2)] }), true],
    ['float noise: the engine 0.1+0.2 against the saved 0.3', () => xlsxWith({ 1: [fcell('C1', 'A1:A2+0.2', 0.3, array('C1:C2')), ncell('A1', 0.1), ncell('A2', 0.1)], 2: [ncell('C2', 0.3)] }), false],
    ['1e-9 relative: just inside', () => xlsxWith({ 1: [ncell('A1', 1), fcell('C1', 'A1:A2*1', 1, array('C1:C2')), ncell('A2', 5)], 2: [ncell('C2', 5 * (1 + 9e-10))] }), false],
    ['1e-9 relative: just outside', () => xlsxWith({ 1: [ncell('A1', 1), fcell('C1', 'A1:A2*1', 1, array('C1:C2')), ncell('A2', 5)], 2: [ncell('C2', 5 * (1 + 2e-9))] }), true],
    ['1e-8 difference', () => xlsxWith({ 1: [ncell('A1', 1), fcell('C1', 'A1:A2*1', 1, array('C1:C2')), ncell('A2', 5)], 2: [ncell('C2', 5.00000005)] }), true],
    ['type mismatch: the saved text "1" against the engine 1', () => xlsxWith({ 1: [ncell('A1', 1), fcell('C1', 'A1:A2*1', 1, array('C1:C2')), ncell('A2', 1)], 2: [sc('C2', '1')] }), true],
    ['type mismatch on the anchor: the saved text "1" against 1', () => xlsxWith({ 1: [ncell('A1', 1), fcell('C1', 'A1:A2*1', '1', array('C1:C2')), ncell('A2', 1)], 2: [ncell('C2', 1)] }), true],
    ['saved error member, the engine has the same error', () => xlsxWith({ 1: [ncell('A1', 1), fcell('C1', 'A1:A2/0', '#DIV/0!', array('C1:C2')).replace('t="str"', 't="e"'), ncell('A2', 1)], 2: [ec('C2', '#DIV/0!')] }), false],
    ['saved error member, the engine has another error', () => xlsxWith({ 1: [ncell('A1', 1), fcell('C1', 'A1:A2/0', '#DIV/0!', array('C1:C2')).replace('t="str"', 't="e"'), ncell('A2', 1)], 2: [ec('C2', '#N/A')] }), true],
    ['saved error member where the engine has a number', () => xlsxWith({ 1: [ncell('A1', 1), fcell('C1', 'A1:A2*1', 1, array('C1:C2')), ncell('A2', 1)], 2: [ec('C2', '#DIV/0!')] }), true],
  ];
  it.each(table)('%s', async (_name, build, cachedOnly) => {
    const { wb, loss } = await open(await build());
    expect(cellAt(wb, 0, 2)?.cachedOnly === true).toBe(cachedOnly);
    expect(cachedOnlyItem(loss) !== undefined).toBe(cachedOnly);
  });

  it('keeps an evaluated array with a blank member evaluated: it follows its precedents and its saved 0 is the engine\'s 0 for the blank', async () => {
    const { wb } = await open(await over(fcell('C1', 'A1:A3', 1, array('C1:C3')), ncell('C2', 0), ncell('C3', 3)));
    wb.setCellValue(undefined, 1, 0, 9);
    expect(wb.getSpillRegion(undefined, 0, 2)?.values).toEqual([[1], [9], [3]]);
  });

  it('shows 0 for the blank source cell and writes it back as a number cell', async () => {
    const { wb } = await open(await over(fcell('C1', 'A1:A3', 1, array('C1:C3')), ncell('C2', 0), ncell('C3', 3)));
    expect(wb.getSpillRegion(undefined, 0, 2)?.values).toEqual([[1], [0], [3]]);
    const sheet = await (await JSZip.loadAsync(await exportToXlsx(wb))).file('xl/worksheets/sheet1.xml')!.async('string');
    expect(sheet).toMatch(/<c r="C2"[^>]*><v>0<\/v><\/c>/);
    expect(sheet).not.toMatch(/<c r="C2"[^>]*t="/);
  });

  it('holds date-styled members to their serial numbers', async () => {
    const base = new WorkbookImpl('wb', 'WB');
    base.setCell(undefined, 0, 0, { value: 45000, formatId: base.getFormatPool().getOrCreate({ type: 'date', dateFormat: 'YYYY-MM-DD' }) });
    const zip = await JSZip.loadAsync(await exportToXlsx(base));
    const style = /<c r="A1" s="(\d+)"/.exec(await zip.file('xl/worksheets/sheet1.xml')!.async('string'))![1];
    const build = async (saved: number) => {
      const rows = `<row r="1"><c r="A1" s="${style}"><v>45000</v></c><c r="C1" s="${style}"><f ${array('C1:C2')}>A1:A2</f><v>45000</v></c></row><row r="2"><c r="A2" s="${style}"><v>45001</v></c><c r="C2" s="${style}"><v>${saved}</v></c></row>`;
      zip.file('xl/worksheets/sheet1.xml', (await zip.file('xl/worksheets/sheet1.xml')!.async('string')).replace(/<sheetData>[\s\S]*<\/sheetData>/, `<sheetData>${rows}</sheetData>`));
      return zip.generateAsync({ type: 'uint8array' });
    };
    expect(cellAt((await open(await build(45001))).wb, 0, 2)?.cachedOnly).toBeUndefined();
    expect(cellAt((await open(await build(45002))).wb, 0, 2)?.cachedOnly).toBe(true);
  });
});

describe('the dependents of an array that keeps its saved results', () => {
  const book = () => xlsxWith({ 1: [fcell('C1', 'ROW(A1:A3)', 1, array('C1:C3')), fcell('E1', 'SUM(C1:C3)', 6)], 2: [ncell('C2', 2)], 3: [ncell('C3', 3)] });

  it('keep what the file saved on a plain open (E1 stays 6, the sum of the saved 1, 2, 3)', async () => {
    const { wb } = await open(await book());
    expect(values(wb, 3, 1, 0, 2)).toEqual([[1], [2], [3]]);
    expect(wb.getCellValue(undefined, 0, 4)).toBe(6);
  });

  it('are recalculated from the kept results when the host recalculates on load (E1 is 6, not the 3 of the engine\'s 1, 1, 1)', async () => {
    const wb = new WorkbookImpl('wb', 'WB');
    await importFromXlsx(await book(), wb, { recalculateOnLoad: true });
    expect(values(wb, 3, 1, 0, 2)).toEqual([[1], [2], [3]]);
    expect(wb.getCellValue(undefined, 0, 4)).toBe(6);
    expect(cellAt(wb, 0, 2)?.cachedOnly).toBe(true);
  });

  it('are left alone with recalculateOnLoad false, even when the saved value is not what the kept results sum to', async () => {
    const wb = new WorkbookImpl('wb', 'WB');
    await importFromXlsx(await xlsxWith({ 1: [fcell('C1', 'ROW(A1:A3)', 1, array('C1:C3')), fcell('E1', 'SUM(C1:C3)', 99)], 2: [ncell('C2', 2)], 3: [ncell('C3', 3)] }), wb, { recalculateOnLoad: false });
    expect(wb.getCellValue(undefined, 0, 4)).toBe(99);
  });
});

describe('a dynamic array (cm="1") whose spill differs from the saved one', () => {
  // Decision (SD4, "the saved value wins on open"): a dynamic array is not compared. Files that never calculated (xlsxwriter and the like) save a
  // 0 for the anchor and nothing for the spill, so holding the engine to them would freeze every dynamic array of such a file; the engine's spill replaces the saved cells.
  const dyn = (c2: number, c3: number) => xlsxWith({ 1: [ncell('A1', 1), fcell('C1', 'A1:A3*2', 2, array('C1:C3'), ' cm="1"')], 2: [ncell('A2', 2), ncell('C2', c2)], 3: [ncell('A3', 3), ncell('C3', c3)] });

  it.each([[4, 6], [4, 7]])('is evaluated by the engine whatever the file saved (saved %d and %d)', async (c2, c3) => {
    const { wb, loss } = await open(await dyn(c2, c3));
    expect(cellAt(wb, 0, 2)).toMatchObject({ arraySpan: { rows: 3, cols: 1, spill: true } });
    expect(cellAt(wb, 0, 2)?.cachedOnly).toBeUndefined();
    expect(wb.getSpillRegion(undefined, 0, 2)?.values).toEqual([[2], [4], [6]]);
    expect(cachedOnlyItem(loss)).toBeUndefined();
  });

  it('leaves the engine\'s #SPILL! alone when a cell is in the way', async () => {
    const { wb, loss } = await open(await xlsxWith({ 1: [fcell('C1', 'A1:A3*2', 2, array('C1:C3'), ' cm="1"'), ncell('A1', 1)], 2: [fcell('C2', 'A1+1', 5)], 3: [ncell('C3', 6)] }));
    expect(wb.getCellValue(undefined, 0, 2)).toBe('#SPILL!');
    expect(cachedOnlyItem(loss)).toBeUndefined();
  });
});
