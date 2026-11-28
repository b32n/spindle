import JSZip from 'jszip';
import ExcelJS from 'exceljs';
import { exportToXlsx, importFromXlsx } from './xlsx';
import { ecell, fcell, ncell, xlsxWith } from './test-xlsx';
import { WorkbookImpl } from '../workbook';
import type { Cell } from '../types';

// S4: formula interchange. Files are made by replacing the sheet data of a workbook this editor exported, so every case is a
// small hand-written <sheetData>: what Excel writes (`_xlfn.` prefixes, `t="array"` with `cm`), with the results Excel saved.

async function open(bytes: Uint8Array, options: Parameters<typeof importFromXlsx>[2] = {}) {
  const wb = new WorkbookImpl('wb', 'WB');
  const result = await importFromXlsx(bytes, wb, options);
  return { wb, ...result };
}

const cellAt = (wb: WorkbookImpl, row: number, col: number): Cell | undefined => wb.getCell(undefined, row, col);

async function sheetXml(bytes: Uint8Array): Promise<string> {
  const zip = await JSZip.loadAsync(bytes);
  return zip.file('xl/worksheets/sheet1.xml')!.async('string');
}

describe('prefixed functions the engine knows (_xlfn., _xlws.)', () => {
  it('are read by their plain name, keep the saved result, and recalculate when a precedent changes', async () => {
    const bytes = await xlsxWith({
      1: [ncell('A1', 10), fcell('B1', '_xlfn.IFS(A1>5,"big",TRUE,"small")', 'SAVED')],
      2: [fcell('A2', '_xlfn.CONCAT(B1,"!")', 'SAVED!')],
    });
    const { wb, loss } = await open(bytes);
    expect(cellAt(wb, 0, 1)).toMatchObject({ formula: '=IFS(A1>5,"big",TRUE,"small")', value: 'SAVED' });
    expect(cellAt(wb, 0, 1)?.cachedOnly).toBeUndefined();
    expect(cellAt(wb, 1, 0)).toMatchObject({ formula: '=CONCAT(B1,"!")', value: 'SAVED!' });
    // not a formula the editor lacks: nothing to report
    expect(loss.items.map((i) => i.id)).not.toContain('xlsx.formulas.cachedOnly');
    expect(loss.items.map((i) => i.id)).not.toContain('xlsx.formulas.unsupported');

    wb.setCellValue(undefined, 0, 0, 1);
    expect(cellAt(wb, 0, 1)?.value).toBe('small');
    expect(cellAt(wb, 1, 0)?.value).toBe('small!');
  });

  it('are recalculated on load when the host asks (recalculateOnLoad)', async () => {
    const bytes = await xlsxWith({ 1: [ncell('A1', 10), fcell('B1', '_xlfn.IFS(A1>5,"big",TRUE,"small")', 'SAVED')] });
    expect((await open(bytes, { recalculateOnLoad: true })).wb.getCellValue(undefined, 0, 1)).toBe('big');
    expect((await open(bytes, { recalculateOnLoad: false })).wb.getCellValue(undefined, 0, 1)).toBe('SAVED');
  });

  it('keep a formula with no saved result evaluated, and a volatile one recalculated, on load', async () => {
    const bytes = await xlsxWith({ 1: [ncell('A1', 4), fcell('B1', 'A1*2', null), fcell('C1', 'TODAY()', 1)] });
    const { wb } = await open(bytes);
    expect(cellAt(wb, 0, 1)?.value).toBe(8);
    expect(cellAt(wb, 0, 2)?.value).not.toBe(1);
  });

  it('are written back with their prefixes, and the file reads the same again', async () => {
    const bytes = await xlsxWith({
      1: [ncell('A1', 3), fcell('B1', '_xlfn.IFS(A1>5,"big",TRUE,"small")', 'small'), fcell('C1', 'SUM(1,2)', 3), fcell('D1', '_xlfn._xlws.SORT(A1:A1)', 3)],
    });
    const { wb } = await open(bytes);
    const out = await exportToXlsx(wb);
    const xml = await sheetXml(out);
    expect(xml).toContain('<f>_xlfn.IFS(A1&gt;5,&quot;big&quot;,TRUE,&quot;small&quot;)</f>');
    expect(xml).toMatch(/<f>SUM\(1,2\)<\/f>/);
    expect(xml).toMatch(/_xlfn\._xlws\.SORT\(A1:A1\)/);
    const again = await open(out);
    expect(again.wb.getData().sheets[0].cells.map((c) => c.cell.formula).sort()).toEqual(
      wb.getData().sheets[0].cells.map((c) => c.cell.formula).sort(),
    );
  });
});

describe('implicit intersection and spill references', () => {
  it('maps _xlfn.SINGLE and _xlfn.ANCHORARRAY both ways', async () => {
    const bytes = await xlsxWith({
      1: [ncell('A1', 5), ncell('A2', 6)],
      3: [fcell('C3', '_xlfn.SINGLE(A1:A2)', 5), fcell('D3', 'SUM(_xlfn.ANCHORARRAY(C3))', 5)],
    });
    const { wb } = await open(bytes);
    expect(cellAt(wb, 2, 2)).toMatchObject({ formula: '=@A1:A2', value: 5 });
    expect(cellAt(wb, 2, 3)).toMatchObject({ formula: '=SUM(C3#)' });
    expect(cellAt(wb, 2, 2)?.cachedOnly).toBeUndefined();
    const xml = await sheetXml(await exportToXlsx(wb));
    // no array mark: `@` before a range is the intersection Excel applies to the bare range anyway (and Excel 2016 has no SINGLE)
    expect(xml).toMatch(/<f>A1:A2<\/f>/);
    expect(xml).toMatch(/<f>SUM\(_xlfn\.ANCHORARRAY\(C3\)\)<\/f>/);
    expect(xml).not.toContain('@A1');
    expect(xml).not.toMatch(/C3#/);
  });

  it('writes a formula typed with @ the same way (a bare range)', async () => {
    const wb = new WorkbookImpl('wb', 'WB');
    wb.setCellValue(undefined, 0, 0, 5);
    wb.setFormula(undefined, 0, 1, '=@A1:A2');
    const xml = await sheetXml(await exportToXlsx(wb));
    expect(xml).toMatch(/<f>A1:A2<\/f>/);
  });
});

describe('cached-only formulas (the engine cannot evaluate them)', () => {
  const LET = '_xlfn.LET(_xlpm.x,A1,_xlpm.y,A2,_xlpm.x+_xlpm.y)';

  async function opened() {
    const bytes = await xlsxWith({
      1: [ncell('A1', 2), fcell('B1', LET, 99)],
      2: [ncell('A2', 3), fcell('B2', 'A1+++', 7), fcell('C2', 'Quarter1*2', 54), fcell('D2', 'B1+1', 100)],
    });
    return open(bytes);
  }

  it('keep the file text and the saved result, whatever changes under them, and are reported', async () => {
    const { wb, loss } = await opened();
    expect(cellAt(wb, 0, 1)).toMatchObject({ formula: `=${LET}`, value: 99, cachedOnly: true });
    expect(cellAt(wb, 1, 1)).toMatchObject({ formula: '=A1+++', value: 7, cachedOnly: true });
    expect(cellAt(wb, 1, 2)).toMatchObject({ formula: '=Quarter1*2', value: 54, cachedOnly: true });
    expect(cellAt(wb, 0, 1)?.formulaAst).toBeUndefined();

    // nothing shows #NAME?; a formula that reads a cached-only cell reads its saved result
    wb.setCellValue(undefined, 0, 0, 40);
    expect(cellAt(wb, 0, 1)?.value).toBe(99);
    expect(cellAt(wb, 1, 1)?.value).toBe(7);
    expect(wb.getCellCalculatedValue(undefined, 1, 3)).toBe(100);

    const item = loss.items.find((i) => i.id === 'xlsx.formulas.cachedOnly');
    expect(item).toMatchObject({ count: 3, state: 'kept' });
    expect(item?.effect).toMatch(/out of date/);
    expect(item?.effect).toMatch(/the saved value and the formula are kept and written back/);
    expect(item?.effect).toMatch(/cell references follow row and column insert and delete \(they do not follow anything else/);
    expect(item?.effect).not.toMatch(/unchanged/);
    const unsupported = loss.items.find((i) => i.id === 'xlsx.formulas.unsupported');
    expect(unsupported?.effect).toMatch(/^kept: shown from the value saved in the file/);
  });

  // S4b: a cached-only cell keeps no parsed formula, so structural edits rewrite the references in its text (shiftFormulaRefs).
  // It keeps the saved value; Excel recalculates it on reopen.
  it('adjusts the cell references in a cached-only formula when rows or columns are inserted or deleted', async () => {
    const { wb } = await opened();
    wb.getSheet().insertRows(0, 1);
    // the cell moved down one row and its text follows the cells it read
    expect(cellAt(wb, 1, 1)).toMatchObject({ formula: '=_xlfn.LET(_xlpm.x,A2,_xlpm.y,A3,_xlpm.x+_xlpm.y)', value: 99, cachedOnly: true });
    expect(cellAt(wb, 2, 1)).toMatchObject({ formula: '=A2+++', value: 7, cachedOnly: true });
    expect(cellAt(wb, 1, 0)?.value).toBe(2);
    expect(cellAt(wb, 0, 0)).toBeUndefined();
    wb.getSheet().deleteRows(0, 1);
    expect(cellAt(wb, 0, 1)?.formula).toBe(`=${LET}`);
    wb.getSheet().insertCols(0, 1);
    expect(cellAt(wb, 0, 2)).toMatchObject({ formula: `=${LET.replace('x,A1', 'x,B1').replace('y,A2', 'y,B2')}`, value: 99 });
  });

  it('keep a saved error result as an error value (t="e") on export, and open the same again', async () => {
    const bytes = await xlsxWith({
      1: [ncell('A1', 0), ecell('B1', '_xlfn.LET(_xlpm.x,1/A1,_xlpm.x)', '#DIV/0!'), ecell('C1', 'NOSUCHFN(A1)', '#NAME?'), ecell('D1', '1/A1', '#DIV/0!')],
    });
    const { wb } = await open(bytes);
    expect(cellAt(wb, 0, 1)).toMatchObject({ value: '#DIV/0!', cachedOnly: true });
    expect(cellAt(wb, 0, 2)).toMatchObject({ value: '#NAME?', cachedOnly: true });
    const xml = await sheetXml(await exportToXlsx(wb));
    expect(xml).toMatch(/<c r="B1" t="e"><f>_xlfn\.LET\(_xlpm\.x,1\/A1,_xlpm\.x\)<\/f><v>#DIV\/0!<\/v><\/c>/);
    expect(xml).toMatch(/<c r="C1" t="e"><f>NOSUCHFN\(A1\)<\/f><v>#NAME\?<\/v><\/c>/);
    expect(xml).not.toMatch(/t="str"/);
    const again = await open(await exportToXlsx(wb));
    for (const col of [1, 2]) expect(cellAt(again.wb, 0, col)).toMatchObject({ value: cellAt(wb, 0, col)?.value, cachedOnly: true, formula: cellAt(wb, 0, col)?.formula });
  });

  it('writes the error an ordinary formula evaluates to (=1/0, an unknown function, a spill blocked by a value) as an error value', async () => {
    const wb = new WorkbookImpl('wb', 'WB');
    wb.setFormula(undefined, 0, 0, '=1/0');
    wb.setFormula(undefined, 0, 1, '=NOSUCHFN(1)');
    wb.setCellValue(undefined, 2, 0, 'in the way');
    wb.setFormula(undefined, 1, 0, '=SEQUENCE(3)');
    expect(cellAt(wb, 0, 0)?.value).toBe('#DIV/0!');
    expect(cellAt(wb, 1, 0)?.value).toBe('#SPILL!');
    const xml = await sheetXml(await exportToXlsx(wb));
    expect(xml).toMatch(/<c r="A1" t="e"><f>1\/0<\/f><v>#DIV\/0!<\/v><\/c>/);
    expect(xml).toMatch(/<c r="B1" t="e"><f>NOSUCHFN\(1\)<\/f><v>#NAME\?<\/v><\/c>/);
    expect(xml).toMatch(/<c r="A2" t="e"><f>_xlfn\.SEQUENCE\(3\)<\/f><v>#SPILL!<\/v><\/c>/);
    const again = await open(await exportToXlsx(wb));
    expect(cellAt(again.wb, 0, 0)?.value).toBe('#DIV/0!');
    expect(cellAt(again.wb, 0, 1)?.value).toBe('#NAME?');
  });

  it('are written back unchanged, formula and result, and stay cached-only on the next open', async () => {
    const { wb } = await opened();
    const out = await exportToXlsx(wb);
    const xml = await sheetXml(out);
    expect(xml).toContain(`<f>${LET}</f>`);
    expect(xml).toMatch(/<f>A1\+\+\+<\/f>/);
    expect(xml).toMatch(/<f>Quarter1\*2<\/f>/);
    expect(xml).toMatch(/<v>99<\/v>/);
    const again = await open(out);
    const flags = (w: WorkbookImpl) =>
      w
        .getData()
        .sheets[0].cells.map((c) => [c.cell.formula, c.cell.value, c.cell.cachedOnly])
        .sort((a, b) => String(a[0]).localeCompare(String(b[0])));
    expect(flags(again.wb)).toEqual(flags(wb));
    const reader = new ExcelJS.Workbook();
    // @ts-expect-error -- Buffer<ArrayBuffer> vs exceljs's plain Buffer type
    await reader.xlsx.load(Buffer.from(out));
    expect(reader.worksheets[0].getCell('B1').result).toBe(99);
  });

  it('lose the mark when the user edits the cell, as a value or as a formula', async () => {
    const { wb } = await opened();
    wb.setFormula(undefined, 0, 1, '=A1+A2');
    expect(cellAt(wb, 0, 1)).toMatchObject({ value: 5 });
    expect(cellAt(wb, 0, 1)?.cachedOnly).toBeUndefined();
    wb.setCellValue(undefined, 1, 1, 'typed');
    expect(cellAt(wb, 1, 1)?.cachedOnly).toBeUndefined();
    expect(cellAt(wb, 1, 1)?.formula).toBeUndefined();
    wb.setFormula(undefined, 1, 2, '=A1+++');
    expect(cellAt(wb, 1, 2)?.cachedOnly).toBeUndefined();
  });

  it('survive undo, redo and renaming or adding a sheet', async () => {
    const { wb } = await opened();
    wb.setCellValue(undefined, 5, 5, 1);
    wb.undo();
    expect(cellAt(wb, 0, 1)).toMatchObject({ value: 99, cachedOnly: true, formula: `=${LET}` });
    wb.redo();
    expect(cellAt(wb, 0, 1)).toMatchObject({ value: 99, cachedOnly: true });
    wb.addSheet('Quarter1');
    expect(cellAt(wb, 1, 2)).toMatchObject({ value: 54, cachedOnly: true, formula: '=Quarter1*2' });
  });
});

describe('array and dynamic-array formulas', () => {
  const sorted = (extra: string[] = []) =>
    xlsxWith({
      1: [ncell('A1', 5), ncell('B1', 1), fcell('D1', '_xlfn._xlws.SORT(A1:A3)', 1, 't="array" ref="D1:D3"', ' cm="1"')],
      2: [ncell('A2', 3), ncell('D2', 3)],
      3: [ncell('A3', 4), ncell('D3', 4), ...extra],
    });

  it('keep an array formula the engine evaluates as a spill: the saved result cells are the spill, not cells', async () => {
    const { wb } = await open(await sorted());
    expect(cellAt(wb, 0, 3)).toMatchObject({ formula: '=SORT(A1:A3)', arraySpan: { rows: 3, cols: 1 } });
    expect(cellAt(wb, 1, 3)).toBeUndefined();
    expect(wb.getSpillRegion(undefined, 0, 3)).toMatchObject({ rows: 3, cols: 1, values: [[3], [4], [5]] });
    expect(wb.getCellValue(undefined, 0, 3)).toBe(3);
    // and follows its precedents
    wb.setCellValue(undefined, 0, 0, 0);
    expect(wb.getSpillRegion(undefined, 0, 3)?.values).toEqual([[0], [3], [4]]);
  });

  it('write an array formula back with its extent, and the spilled values next to it', async () => {
    const { wb } = await open(await sorted());
    const out = await exportToXlsx(wb);
    const xml = await sheetXml(out);
    expect(xml).toMatch(/<c r="D1" cm="1"><f t="array" ref="D1:D3">_xlfn\._xlws\.SORT\(A1:A3\)<\/f><v>3<\/v><\/c>/);
    expect(xml).toMatch(/<c r="D2"><v>4<\/v><\/c>/);
    expect(xml).toMatch(/<c r="D3"><v>5<\/v><\/c>/);
    const again = await open(out);
    expect(again.wb.getData()).toEqual(expect.objectContaining({}));
    expect(cellAt(again.wb, 0, 3)).toMatchObject({ formula: '=SORT(A1:A3)', arraySpan: { rows: 3, cols: 1 } });
    expect(again.wb.getSpillRegion(undefined, 0, 3)?.values).toEqual([[3], [4], [5]]);
  });

  describe('on export, legacy and dynamic arrays stay what they were', () => {
    const exported = async (wb: WorkbookImpl) => {
      const out = await exportToXlsx(wb);
      const zip = await JSZip.loadAsync(out);
      return { out, xml: await zip.file('xl/worksheets/sheet1.xml')!.async('string'), hasMetadata: zip.file('xl/metadata.xml') !== null };
    };

    it('a legacy A1:B2 array formula stays {=...} (no cm, no metadata part) though the engine spills it', async () => {
      const { wb } = await open(await xlsxWith({ 1: [fcell('A1', '1+1', 2, 't="array" ref="A1:B2"')] }));
      expect(wb.getSpillRegion(undefined, 0, 0)).toMatchObject({ rows: 2, cols: 2 });
      const { out, xml, hasMetadata } = await exported(wb);
      expect(xml).toMatch(/<c r="A1"><f t="array" ref="A1:B2">1\+1<\/f><v>2<\/v><\/c>/);
      expect(xml).not.toContain('cm=');
      expect(hasMetadata).toBe(false);
      const again = await open(out);
      expect(cellAt(again.wb, 0, 0)?.arraySpan).toEqual({ rows: 2, cols: 2 });
    });

    it('a dynamic array stays dynamic (cm="1" and the metadata part)', async () => {
      const { wb } = await open(await sorted());
      const { out, xml, hasMetadata } = await exported(wb);
      expect(xml).toContain('<c r="D1" cm="1">');
      expect(hasMetadata).toBe(true);
      expect(cellAt((await open(out)).wb, 0, 3)?.arraySpan).toEqual({ rows: 3, cols: 1, spill: true });
    });

    it('a formula typed here that spills is written as a dynamic array', async () => {
      const wb = new WorkbookImpl('wb', 'WB');
      wb.setFormula(undefined, 0, 0, '=SEQUENCE(2,2)');
      expect(cellAt(wb, 0, 0)?.arraySpan).toBeUndefined();
      const { out, xml, hasMetadata } = await exported(wb);
      expect(xml).toMatch(/<c r="A1" cm="1"><f t="array" ref="A1:B2">/);
      expect(hasMetadata).toBe(true);
      expect(cellAt((await open(out)).wb, 0, 0)?.arraySpan).toEqual({ rows: 2, cols: 2, spill: true });
    });
  });

  it('keep a one-cell array formula as one, so the file still says array', async () => {
    const bytes = await xlsxWith({ 1: [ncell('A1', 2), ncell('B1', 3), fcell('C1', 'SUM(A1:B1*A1:B1)', 13, 't="array" ref="C1"')] });
    const { wb } = await open(bytes);
    expect(cellAt(wb, 0, 2)).toMatchObject({ formula: '=SUM(A1:B1*A1:B1)', value: 13, arraySpan: { rows: 1, cols: 1 } });
    expect(await sheetXml(await exportToXlsx(wb))).toMatch(/<f t="array" ref="C1">SUM\(A1:B1\*A1:B1\)<\/f>/);
  });

  it('keep an array formula the engine cannot evaluate as saved values: formula, result cells and extent unchanged', async () => {
    const f = '_xlfn.LET(_xlpm.x,A1:A3,_xlfn._xlws.SORT(_xlpm.x))';
    const bytes = await xlsxWith({
      1: [ncell('A1', 5), fcell('D1', f, 3, 't="array" ref="D1:D3"', ' cm="1"')],
      2: [ncell('A2', 3), ncell('D2', 4)],
      3: [ncell('A3', 4), ncell('D3', 5)],
    });
    const { wb, loss } = await open(bytes);
    expect(cellAt(wb, 0, 3)).toMatchObject({ formula: `=${f}`, value: 3, cachedOnly: true, arraySpan: { rows: 3, cols: 1 } });
    expect(cellAt(wb, 1, 3)).toMatchObject({ value: 4 });
    expect(cellAt(wb, 2, 3)).toMatchObject({ value: 5 });
    expect(loss.items.find((i) => i.id === 'xlsx.formulas.cachedOnly')?.count).toBe(1);
    const xml = await sheetXml(await exportToXlsx(wb));
    expect(xml).toContain(`<f t="array" ref="D1:D3">${f}</f>`);
    expect(xml).toMatch(/<c r="D2"><v>4<\/v><\/c>/);
  });

  it('write a formula typed as a spill as an array formula over what it spilled', async () => {
    const wb = new WorkbookImpl('wb', 'WB');
    for (let r = 0; r < 3; r++) wb.setCellValue(undefined, r, 0, 3 - r);
    wb.setFormula(undefined, 0, 2, '=SORT(A1:A3)');
    const xml = await sheetXml(await exportToXlsx(wb));
    expect(xml).toMatch(/<f t="array" ref="C1:C3">_xlfn\._xlws\.SORT\(A1:A3\)<\/f>/);
  });

  it('give an unreadable or reversed ref no extent, and keep the formula', async () => {
    const bytes = await xlsxWith({ 1: [ncell('A1', 1), fcell('B1', 'A1+1', 2, 't="array" ref="B1:A1"'), fcell('C1', 'A1+2', 3, 't="array" ref="nonsense"')] });
    const { wb } = await open(bytes);
    expect(cellAt(wb, 0, 1)).toMatchObject({ formula: '=A1+1', value: 2 });
    expect(cellAt(wb, 0, 1)?.arraySpan).toBeUndefined();
    expect(cellAt(wb, 0, 2)?.arraySpan).toBeUndefined();
  });

  it("reads an array whose range does not start at its cell as a one-cell formula, keeps the plain cells around it and warns", async () => {
    // F5 claims A1:B2: the extent is not F5:G6, so G5 and F6 are not part of it and must not be dropped
    const bytes = await xlsxWith({
      5: [fcell('F5', '1+1', 2, 't="array" ref="A1:B2"'), ncell('G5', 55)],
      6: [ncell('F6', 66)],
    });
    const { wb, warnings } = await open(bytes);
    expect(cellAt(wb, 4, 5)).toMatchObject({ formula: '=1+1', value: 2 });
    expect(cellAt(wb, 4, 5)?.arraySpan).toBeUndefined();
    expect(cellAt(wb, 4, 6)?.value).toBe(55);
    expect(cellAt(wb, 5, 5)?.value).toBe(66);
    expect(wb.getSpillRegion(undefined, 4, 5)).toBeUndefined();
    expect(warnings.filter((w) => /does not start at the cell/.test(w))).toEqual([
      "Sheet1!F5: the array formula's range A1:B2 does not start at the cell, so it is read as a one-cell formula.",
    ]);
  });

  it.each([
    ['row only', 'B2', 'B1:C2'],
    ['column only', 'B2', 'C2:D3'],
    ['a range that ends at the cell', 'B2', 'A1:B2'],
    ['a cell elsewhere', 'B2', 'A1'],
  ])('reads %s of an array ref off its anchor as a one-cell formula', async (_name, anchor, ref) => {
    const { wb, warnings } = await open(await xlsxWith({ 2: [fcell(anchor, '1+1', 2, `t="array" ref="${ref}"`, ' cm="1"'), ncell('C2', 7)], 3: [ncell('B3', 8)] }));
    expect(cellAt(wb, 1, 1)?.arraySpan).toBeUndefined();
    expect(cellAt(wb, 1, 2)?.value).toBe(7);
    expect(cellAt(wb, 2, 1)?.value).toBe(8);
    expect(warnings.some((w) => /does not start at the cell/.test(w))).toBe(true);
  });

  it('counts the warnings for misplaced array ranges after a small cap', async () => {
    const rows: Record<number, string[]> = {};
    for (let i = 2; i < 28; i++) rows[i] = [fcell(`B${i}`, '1+1', 2, 't="array" ref="A1:B2"')];
    const { warnings } = await open(await xlsxWith(rows));
    const mine = warnings.filter((w) => /does not start at the cell|whose range does not start/.test(w));
    expect(mine).toHaveLength(21);
    expect(mine[20]).toBe('6 more array formulas whose range does not start at their cell were read as one-cell formulas.');
  });

  it('keeps a legacy array formula over a huge extent as saved values and says so (xlsx.package.sizeLimit)', async () => {
    const bytes = await xlsxWith({ 1: [ncell('A1', 1), fcell('B1', 'SUM(A1:A1)', 1, 't="array" ref="B1:XFD1048576"')] });
    const { wb, loss } = await open(bytes);
    expect(cellAt(wb, 0, 1)).toMatchObject({ formula: '=SUM(A1:A1)', value: 1, cachedOnly: true });
    expect(loss.items.find((i) => i.id === 'xlsx.package.sizeLimit' && /array formula/.test(i.label))).toMatchObject({ count: 1 });
  });

  it('keeps dynamic arrays past the clearing budget as saved values and says so', async () => {
    // 2,600 dynamic-array formulas over a huge extent each, on a sheet of 2,600 cells: each costs the cell count to clear, together past 1,000,000
    const rows: Record<number, string[]> = {};
    for (let i = 0; i < 2600; i++) rows[i + 1] = [fcell(`A${i + 1}`, 'SUM(1,2)', 3, `t="array" ref="A${i + 1}:Z1000000"`, ' cm="1"')];
    const { wb, loss } = await open(await xlsxWith(rows));
    const cut = loss.items.find((i) => i.id === 'xlsx.package.sizeLimit' && /array formula/.test(i.label));
    expect(cut?.count).toBeGreaterThan(0);
    expect(cut?.count).toBeLessThan(2600);
    expect(cellAt(wb, 2599, 0)).toMatchObject({ cachedOnly: true, value: 3 });
  }, 120_000);

  it('reads a legacy array formula over 4 million cells (ref A1:J400000) as a saved value, with a loss item that names the limit and the option', async () => {
    const { wb, loss } = await open(await xlsxWith({ 1: [fcell('A1', 'SUM(1,2)', 3, 't="array" ref="A1:J400000"')] }));
    expect(cellAt(wb, 0, 0)).toMatchObject({ formula: '=SUM(1,2)', value: 3, cachedOnly: true });
    const cut = loss.items.find((i) => i.id === 'xlsx.package.sizeLimit' && /array formula/.test(i.label));
    expect(cut).toMatchObject({ count: 1 });
    expect(cut?.effect).toContain('over 1000000 cells');
    expect(cut?.effect).toContain('maxArrayCells');
    expect(cut?.effect).toContain('importWorkbookFile');
    expect(cut?.effect).toContain('WorkbookCanvas');
  });

  it('takes the array budget from the maxArrayCells option', async () => {
    const bytes = await xlsxWith({ 1: [fcell('A1', '1+1', 2, 't="array" ref="A1:B2"'), fcell('D1', '1+1', 2, 't="array" ref="D1:E2"')] });
    const small = await open(bytes, { maxArrayCells: 4 });
    expect(cellAt(small.wb, 0, 0)?.cachedOnly).toBeUndefined();
    expect(cellAt(small.wb, 0, 3)).toMatchObject({ cachedOnly: true });
    expect(small.loss.items.find((i) => /array formula/.test(i.label))).toMatchObject({ count: 1 });
    expect(small.loss.items.find((i) => /array formula/.test(i.label))?.effect).toContain('over 4 cells');
    const whole = await open(bytes);
    expect(cellAt(whole.wb, 0, 3)?.cachedOnly).toBeUndefined();
    expect(whole.loss.items.find((i) => /array formula/.test(i.label))).toBeUndefined();
  });

  it('cuts a legacy array formula to its extent: a one-cell one shows the first value, a one-column result is repeated across a wider one', async () => {
    const bytes = await xlsxWith({
      1: [ncell('A1', 1), ncell('A2', 2), fcell('B1', 'A1:A2*10', 10, 't="array" ref="B1"'), fcell('C1', 'A1:A2*10', 10, 't="array" ref="C1:E2"'), ncell('D2', 20)],
    });
    const { wb } = await open(bytes);
    expect(cellAt(wb, 0, 1)).toMatchObject({ value: 10, arraySpan: { rows: 1, cols: 1 } });
    // C1:E2 is 2 x 3 over a 2 x 1 result: Excel repeats a one-column result across the range; the cell D2 inside the extent holds the result the file saved (20, as the engine's), so the engine's spill replaces it
    expect(wb.getSpillRegion(undefined, 0, 2)).toMatchObject({ rows: 2, cols: 3, values: [[10, 10, 10], [20, 20, 20]] });
    expect(cellAt(wb, 1, 3)).toBeUndefined();
  });

  it('fills the extent of a legacy array formula with a scalar result', async () => {
    const { wb } = await open(await xlsxWith({ 1: [fcell('A1', '1+1', 2, 't="array" ref="A1:B2"')] }));
    expect(wb.getSpillRegion(undefined, 0, 0)).toMatchObject({ rows: 2, cols: 2, values: [[2, 2], [2, 2]] });
    expect(cellAt(wb, 0, 0)).toMatchObject({ arraySpan: { rows: 2, cols: 2 } });
  });

  it('keeps a dynamic array formula whose saved extent is one cell as its saved value until a precedent changes', async () => {
    const bytes = await xlsxWith({ 1: [ncell('A1', 1), ncell('A2', 2), fcell('B1', 'A1:A2*10', 777, 't="array" ref="B1"', ' cm="1"')] });
    const { wb } = await open(bytes);
    expect(cellAt(wb, 0, 1)).toMatchObject({ value: 777, arraySpan: { rows: 1, cols: 1, spill: true } });
    wb.setCellValue(undefined, 0, 0, 3);
    expect(wb.getSpillRegion(undefined, 0, 1)).toMatchObject({ rows: 2, cols: 1, values: [[30], [20]] });
    expect(cellAt(wb, 0, 1)?.arraySpan).toEqual({ rows: 2, cols: 1, spill: true });
  });
});

describe('what the export writes of one array formula (maxArrayCells)', () => {
  async function seq(rows: number, cols: number) {
    const wb = new WorkbookImpl('wb', 'WB');
    wb.setCellValue(undefined, 0, 0, 1);
    wb.setFormula(undefined, 0, 1, `=SEQUENCE(${rows},${cols})`);
    return wb;
  }

  it('writes a spilled region in full under the cap', async () => {
    const warnings: string[] = [];
    const xml = await sheetXml(await exportToXlsx(await seq(3, 2), { onWarning: (m) => warnings.push(m) }));
    expect(warnings).toEqual([]);
    expect(xml).toContain('ref="B1:C3"');
    expect(xml).toContain('r="C3"');
  });

  it.each([
    // over the cap: the first rows that fit, the range cut to match
    [10, 3, 7, 2, 3, 'B1:D2'],
    // a single row wider than the cap: cut to the cap's columns
    [1, 6, 4, 1, 4, 'B1:E1'],
  ])('cuts a %i x %i result over a cap of %i cells to %i x %i (%s) and says so', async (rows, cols, cap, wRows, wCols, ref) => {
    const warnings: string[] = [];
    const xml = await sheetXml(await exportToXlsx(await seq(rows, cols), { maxArrayCells: cap, onWarning: (m) => warnings.push(m) }));
    expect(xml).toContain(`ref="${ref}"`);
    expect(xml).not.toContain(`r="B${wRows + 1}"`);
    expect(warnings).toEqual([`Sheet1!B1: the array formula's result (${rows} x ${cols} cells) is over ${cap} cells, so only its first ${wRows} x ${wCols} are written to the file.`]);
  });

  it('shares one budget between every array formula of the workbook, over sheets', async () => {
    const wb = new WorkbookImpl('wb', 'WB');
    wb.setFormula(undefined, 0, 0, '=SEQUENCE(3,2)');
    wb.setFormula(undefined, 0, 3, '=SEQUENCE(3,2)');
    const second = wb.addSheet('Two');
    wb.setFormula(second.id, 0, 0, '=SEQUENCE(3,2)');
    const warnings: string[] = [];
    const out = await exportToXlsx(wb, { maxTotalArrayCells: 8, onWarning: (m) => warnings.push(m) });
    const zip = await JSZip.loadAsync(out);
    const one = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
    const two = await zip.file('xl/worksheets/sheet2.xml')!.async('string');
    // 6 of 8 used by A1, the 2 left cut D1 to 1 x 2 and the second sheet's to one cell
    expect(one).toContain('ref="A1:B3"');
    expect(one).toContain('ref="D1:E1"');
    expect(two).toContain('ref="A1"');
    expect(warnings).toEqual([
      "Sheet1!D1: the array formula's result (3 x 2 cells) does not fit in the 2 cells left of the workbook's budget for array formula results, so only its first 1 x 2 are written to the file.",
      "Two!A1: the array formula's result (3 x 2 cells) does not fit in the 0 cells left of the workbook's budget for array formula results, so only its first 1 x 1 are written to the file.",
    ]);
  });

  it('reports the first 20 cuts one by one and counts the others', async () => {
    const wb = new WorkbookImpl('wb', 'WB');
    for (let i = 0; i < 23; i++) wb.setFormula(undefined, 0, i, '=SEQUENCE(3)');
    const warnings: string[] = [];
    await exportToXlsx(wb, { maxArrayCells: 2, onWarning: (m) => warnings.push(m) });
    expect(warnings).toHaveLength(21);
    expect(warnings[20]).toBe('3 more array formula results were cut in the same way.');
    // no listener: the cuts are still made
    await expect(exportToXlsx(wb, { maxArrayCells: 2 })).resolves.toBeInstanceOf(Uint8Array);
  });

  it('cuts without a listener, and keeps at least one cell when the cap is zero', async () => {
    const xml = await sheetXml(await exportToXlsx(await seq(4, 4), { maxArrayCells: 0 }));
    expect(xml).toContain('ref="B1"');
  });
});
