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
    expect(cachedOnlyItem(loss)?.effect).toMatch(/^kept: shown from the value saved in the file and written back unchanged with the original formula text; the value may be out of date/);
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

  it('keeps an array of errors the engine reproduces as evaluated', async () => {
    const { wb, loss } = await open(await allErrors('SEQUENCE(2,2)/0'));
    expect(cellAt(wb, 0, 0)?.cachedOnly).toBeUndefined();
    expect(cachedOnlyItem(loss)).toBeUndefined();
    expect(wb.getSpillRegion(undefined, 0, 0)).toMatchObject({ rows: 2, cols: 2 });
  });
});
