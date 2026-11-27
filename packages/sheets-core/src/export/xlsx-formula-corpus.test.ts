import fs from 'node:fs';
import path from 'node:path';
import { exportToXlsx, importFromXlsx } from './xlsx';
import { WorkbookImpl } from '../workbook';

// S4 over the fixtures/xlsx corpus: no formula of any fixture shows `#NAME?` for a function or name Excel knows, the formulas the
// engine cannot evaluate are exactly the ones named here (each keeps its text and saved result), and a save writes them back.

const ROOT = path.resolve(__dirname, '../../../..');
const read = (file: string) => new Uint8Array(fs.readFileSync(path.join(ROOT, 'fixtures/xlsx', file)));
const corpus = fs.readdirSync(path.join(ROOT, 'fixtures/xlsx')).filter((f) => f.endsWith('.xlsx') && !f.startsWith('edge-'));

/** [file, sheet, address, formula as the formula bar shows it, saved result] of every cached-only formula of the corpus. */
const CACHED_ONLY: Array<[string, string, string, string, unknown]> = [
  ['formulas-core.xlsx', 'Data', 'I132', '=_xlfn.LET(x,2,x*3)', 6],
  ['formulas-core.xlsx', 'Data', 'I133', '=_xlfn.HSTACK(A1:A2,B1:B2)', 1],
  ['formulas-core.xlsx', 'Data', 'I134', '=_xlfn.VSTACK(A1:A2,B1:B2)', 1],
  ['formulas-core.xlsx', 'Data', 'I135', '=_xlfn.TAKE(A1:A10,2)', 1],
  ['formulas-core.xlsx', 'Data', 'I66', '=_xlfn.TEXTBEFORE("a-b","-")', 'a'],
  ['formulas-core.xlsx', 'Data', 'I67', '=_xlfn.TEXTSPLIT("a-b","-")', 'a'],
  ['formulas-modern.xlsx', 'Modern', 'G7', '=_xlfn.LET(_xlpm.x,B2,_xlpm.y,B3,_xlpm.x+_xlpm.y)', 15],
  ['formulas-modern.xlsx', 'Modern', 'G8', '=_xlfn.LAMBDA(_xlpm.n,_xlpm.n*2)(B2)', 20],
  ['xlsxwriter-rich.xlsx', 'Data', 'F2', '=_xlfn.LET(_xlpm.x,2,_xlpm.x*3)', 6],
  ['names.xlsx', 'Data', 'B8', '=SUM(Quarter1)', 27],
  ['names.xlsx', 'Data', 'B9', '=B8*(1+TaxRate)', 32.4],
  ['references.xlsx', 'Data', 'F13', '=[1]Remote!$A$1', 42],
  ['tables.xlsx', 'Tables', 'B6', '=SUBTOTAL(109,Sales[Q1])', 23],
  ['tables.xlsx', 'Tables', 'B8', '=SUM(Sales[Q1])', 23],
  ['tables.xlsx', 'Tables', 'C6', '=SUBTOTAL(109,Sales[Q2])', 45],
  ['tables.xlsx', 'Tables', 'D6', '=SUBTOTAL(109,Sales[Q3])', 67],
  ['tables.xlsx', 'Tables', 'E2', '=Sales[[#This Row],[Q1]]+Sales[[#This Row],[Q2]]+Sales[[#This Row],[Q3]]', 60],
  ['tables.xlsx', 'Tables', 'E3', '=Sales[[#This Row],[Q1]]+Sales[[#This Row],[Q2]]+Sales[[#This Row],[Q3]]', 45],
  ['tables.xlsx', 'Tables', 'E4', '=Sales[[#This Row],[Q1]]+Sales[[#This Row],[Q2]]+Sales[[#This Row],[Q3]]', 24],
  ['tables.xlsx', 'Tables', 'E5', '=Sales[[#This Row],[Q1]]+Sales[[#This Row],[Q2]]+Sales[[#This Row],[Q3]]', 6],
  ['tables.xlsx', 'Tables', 'E6', '=SUBTOTAL(109,Sales[Total])', 135],
];

const label = (c: number) => String.fromCharCode(65 + (c % 26)); // the corpus sheets are narrower than 26 columns

interface Found {
  sheet: string;
  address: string;
  formula: string;
  value: unknown;
}
function cachedOnlyOf(wb: WorkbookImpl): Found[] {
  const found: Found[] = [];
  for (const sd of wb.getData().sheets) {
    const sheet = wb.getSheet(sd.id);
    for (const [row, col, cell] of sheet.entries()) {
      if (cell.cachedOnly) found.push({ sheet: sd.name, address: `${label(col)}${row + 1}`, formula: cell.formula as string, value: cell.value });
    }
  }
  return found;
}

describe('the fixtures/xlsx corpus', () => {
  it.each(corpus)('%s: no formula shows #NAME?, and the cached-only ones are the expected ones', async (file) => {
    const wb = new WorkbookImpl('wb', 'WB');
    await importFromXlsx(read(file), wb);
    for (const sd of wb.getData().sheets) {
      for (const [, , cell] of wb.getSheet(sd.id).entries()) {
        if (cell.formula) expect(String(cell.value)).not.toMatch(/^#NAME\?/);
      }
    }
    const expected = CACHED_ONLY.filter(([f]) => f === file).map(([, sheet, address, formula, value]) => ({ sheet, address, formula, value }));
    const key = (f: Found) => `${f.sheet}!${f.address}`;
    expect(cachedOnlyOf(wb).sort((a, b) => key(a).localeCompare(key(b)))).toEqual(expected.sort((a, b) => key(a).localeCompare(key(b))));
  });

  it('writes every cached-only formula back as it came, with its saved result', async () => {
    const wb = new WorkbookImpl('wb', 'WB');
    await importFromXlsx(read('formulas-core.xlsx'), wb);
    const again = new WorkbookImpl('wb', 'WB');
    await importFromXlsx(await exportToXlsx(wb), again);
    expect(cachedOnlyOf(again)).toEqual(cachedOnlyOf(wb));
    expect(cachedOnlyOf(again).length).toBe(6);
  });

  it('recalculating on load (the harness values step) evaluates the prefixed functions the engine knows', async () => {
    const wb = new WorkbookImpl('wb', 'WB');
    await importFromXlsx(read('formulas-core.xlsx'), wb, { recalculateOnLoad: true });
    const data = wb.getData().sheets.find((s) => s.name === 'Data')!;
    const sheet = wb.getSheet(data.id);
    const at = (a: string) => sheet.getCell(Number(a.slice(1)) - 1, a.charCodeAt(0) - 65)?.value;
    // I24 MAXIFS, I32 IFS, I33 SWITCH, I36 XOR, I52 CONCAT, I53 TEXTJOIN, I123 XLOOKUP (the `_xlfn.` rows of the 27 that failed before S4)
    expect([at('I24'), at('I32'), at('I33'), at('I36'), at('I52'), at('I53'), at('I123')]).toEqual([9, 'b', 'one', true, 'ab', 'x,y,x', 3]);
  });
});
