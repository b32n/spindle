import fs from 'node:fs';
import path from 'node:path';
import { exportToXlsx, importFromXlsx } from './xlsx';
import { WorkbookImpl } from '../workbook';

const ROOT = path.resolve(__dirname, '../../../..');
const DUMP_ROWS = 12;
const DUMP_COLS = 8;

/** The model the way e2e/support/sheets-dump.ts reads it in the browser: what S-FIO-04 compares between the first open and the re-open. */
function dumpWorkbook(wb: WorkbookImpl): unknown {
  return wb.getData().sheets.map((sd) => {
    const sheet = wb.getSheet(sd.id)!;
    const cells: Record<string, unknown> = {};
    for (let r = 0; r < DUMP_ROWS; r++) {
      for (let c = 0; c < DUMP_COLS; c++) {
        const cell = wb.getCell(sd.id, r, c);
        if (!cell) continue;
        const out: Record<string, unknown> = { value: wb.getCellCalculatedValue(sd.id, r, c) ?? cell.value ?? null };
        if (cell.formula) out.formula = cell.formula;
        if (cell.cachedOnly) out.cachedOnly = true;
        if (cell.arraySpan) out.arraySpan = cell.arraySpan;
        if (cell.styleId) out.style = wb.getStylePool().get(cell.styleId);
        if (cell.formatId) out.format = wb.getFormatPool().get(cell.formatId);
        if (out.value !== null || out.formula || out.style || out.format) cells[`${String.fromCharCode(65 + c)}${r + 1}`] = out;
      }
    }
    return {
      name: sd.name,
      cells,
      merges: sheet.getMergedRegions().map((m) => [m.startRow, m.startCol, m.endRow, m.endCol]),
      frozen: { rows: sheet.getFrozenRows(), cols: sheet.getFrozenCols() },
      colWidths: Array.from({ length: DUMP_COLS }, (_, c) => sheet.getColWidth(c)),
    };
  });
}

const read = (...parts: string[]) => new Uint8Array(fs.readFileSync(path.join(ROOT, ...parts)));

const e2eFixtures = fs.readdirSync(path.join(ROOT, 'e2e/fixtures')).filter((f) => f.endsWith('.xlsx'));
// The corpus files the editor opens whole: the edge-* files are refused or empty by design (see fixtures/xlsx/README.md).
const corpusFixtures = fs.readdirSync(path.join(ROOT, 'fixtures/xlsx')).filter((f) => f.endsWith('.xlsx') && !f.startsWith('edge-'));

describe('open -> export -> re-open', () => {
  const cases: Array<[string, () => Uint8Array]> = [
    ...e2eFixtures.map((f): [string, () => Uint8Array] => [`e2e/fixtures/${f}`, () => read('e2e/fixtures', f)]),
    ...corpusFixtures.map((f): [string, () => Uint8Array] => [`fixtures/xlsx/${f}`, () => read('fixtures/xlsx', f)]),
  ];

  it('finds the fixtures', () => {
    expect(e2eFixtures).toEqual(expect.arrayContaining(['styled-multisheet.xlsx', 'exceljs-frozen.xlsx']));
    expect(corpusFixtures.length).toBeGreaterThan(10);
  });

  it.each(cases)('%s: the re-opened model equals the first open (what S-FIO-04 asserts in the browser)', async (_name, bytes) => {
    const wb = new WorkbookImpl('wb', 'WB');
    await importFromXlsx(bytes(), wb);
    const first = dumpWorkbook(wb);
    const second = new WorkbookImpl('wb', 'WB');
    await importFromXlsx(await exportToXlsx(wb), second);
    expect(dumpWorkbook(second)).toEqual(first);
  });

  it('keeps no explicit default text rotation (a file that says textRotation="0" is the same style as one that says nothing)', async () => {
    const wb = new WorkbookImpl('wb', 'WB');
    await importFromXlsx(read('e2e/fixtures/styled-multisheet.xlsx'), wb);
    const sales = wb.getData().sheets[0].id;
    const style = wb.getStylePool().get(wb.getCell(sales, 0, 0)!.styleId as string);
    expect(style).toMatchObject({ textAlign: 'center', verticalAlign: 'bottom' });
    expect(style).not.toHaveProperty('textRotation');
  });
});

describe('importing the same file twice', () => {
  it.each([['e2e/fixtures', 'styled-multisheet.xlsx'], ['e2e/fixtures', 'exceljs-frozen.xlsx'], ['fixtures/xlsx', 'styles-basic.xlsx']])(
    '%s/%s: into one workbook and into fresh ones, the model and the loss report are identical (no state is shared between imports)',
    async (dir, file) => {
      const bytes = read(dir, file);
      const shared = new WorkbookImpl('wb', 'WB');
      const runs: Array<{ model: string; loss: string }> = [];
      for (let i = 0; i < 3; i++) {
        const { loss } = await importFromXlsx(bytes, shared);
        runs.push({ model: JSON.stringify(dumpWorkbook(shared)), loss: JSON.stringify(loss) });
      }
      const fresh = new WorkbookImpl('wb', 'WB');
      const { loss } = await importFromXlsx(bytes, fresh);
      runs.push({ model: JSON.stringify(dumpWorkbook(fresh)), loss: JSON.stringify(loss) });
      for (const run of runs) expect(run).toEqual(runs[0]);
    }
  );
});
