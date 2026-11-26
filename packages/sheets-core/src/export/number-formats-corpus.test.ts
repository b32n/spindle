// The two number-format fixtures (fixtures/xlsx): the 60-code corpus file and the 33-code probe. Every code must
// survive import -> export -> import unchanged, and every value must display the way the conformance table says
// (LibreOffice's string, or Excel's where the table notes a difference).

import * as fs from 'fs';
import * as path from 'path';
import { exportToXlsx, importFromXlsx } from './xlsx';
import { WorkbookImpl } from '../workbook';
import { formatNumber } from '../utils/format-utils';
import { toExcelNumberFormatCode } from './xlsx-number-format';

const ROOT = path.resolve(__dirname, '../../../..');
const oracle = JSON.parse(fs.readFileSync(path.join(ROOT, 'packages/sheets-core/src/number-format/number-format-oracle.json'), 'utf8')) as {
  builtinSpelling: Record<string, string>;
  rows: Array<{ group: string; code: string; value: number | string; expected: string }>;
};

interface Imported {
  workbook: WorkbookImpl;
  /** [code the file means, value, shown text, code written by the format] per formatted cell */
  cells: Array<{ code: string; value: number; shown: string; stored: string }>;
}

function readCells(workbook: WorkbookImpl, codes: string[], columns: number, header: number): Imported['cells'] {
  const pool = workbook.getFormatPool();
  const cells: Imported['cells'] = [];
  codes.forEach((code, row) => {
    for (let col = 1; col <= columns; col++) {
      const cell = workbook.getCell(undefined, row + header, col);
      const format = cell?.formatId ? pool.get(cell.formatId) : undefined;
      cells.push({
        code,
        value: cell!.value as number,
        shown: formatNumber(cell!.value as number, format ?? {}),
        stored: format ? toExcelNumberFormatCode(format) : 'General',
      });
    }
  });
  return cells;
}

async function importFixture(name: string, header: number): Promise<{ codes: string[]; imported: Imported; again: Imported }> {
  const bytes = new Uint8Array(fs.readFileSync(path.join(ROOT, 'fixtures/xlsx', name)));
  const workbook = new WorkbookImpl('wb', 'WB');
  await importFromXlsx(bytes, workbook);
  // the code column holds the code as text (row 1 is the header of the corpus file)
  const codes: string[] = [];
  for (let row = header; ; row++) {
    const text = workbook.getCell(undefined, row, 0)?.value;
    if (typeof text !== 'string') break;
    codes.push(text);
  }
  const columns = name === 'number-formats.xlsx' ? 4 : 3;
  const imported = { workbook, cells: readCells(workbook, codes, columns, header) };
  const exported = new WorkbookImpl('wb2', 'WB2');
  await importFromXlsx(await exportToXlsx(workbook), exported);
  return { codes, imported, again: { workbook: exported, cells: readCells(exported, codes, columns, header) } };
}

const meant = (code: string): string => oracle.builtinSpelling[code] ?? code;
const expectedShown = (code: string, value: number): string | undefined =>
  oracle.rows.find((r) => (r.group === 'corpus' || r.group === 'probe' || r.group === 'builtin') && r.code === code && r.value === value)?.expected;

describe('number-formats.xlsx (the 60-code corpus file)', () => {
  it('keeps every code through import, export and import, and shows every value as the conformance table says', async () => {
    const { codes, imported, again } = await importFixture('number-formats.xlsx', 1);
    expect(codes).toHaveLength(60);
    expect(imported.cells).toHaveLength(240);
    for (const cell of imported.cells) {
      expect(cell.stored).toBe(meant(cell.code));
    }
    // survives a save: the same code on every cell after export and import
    expect(again.cells.map((c) => c.stored)).toEqual(imported.cells.map((c) => c.stored));
    const wrong = imported.cells.filter((c) => {
      const expected = expectedShown(meant(c.code), c.value);
      return expected === undefined || c.shown !== expected;
    });
    expect(wrong.map((c) => `${c.code} ${c.value}: ${c.shown}`)).toEqual([]);
  });
});

describe('number-format-probe.xlsx (the 33 probe codes)', () => {
  it('keeps every code (33 of 33) and shows all 99 values as the conformance table says', async () => {
    const { codes, imported, again } = await importFixture('number-format-probe.xlsx', 0);
    expect(codes).toHaveLength(33);
    expect(imported.cells).toHaveLength(99);
    expect(imported.cells.filter((c) => c.stored === meant(c.code))).toHaveLength(99);
    expect(again.cells.map((c) => c.stored)).toEqual(imported.cells.map((c) => c.stored));
    const wrong = imported.cells.filter((c) => c.shown !== expectedShown(meant(c.code), c.value));
    expect(wrong.map((c) => `${c.code} ${c.value}: ${c.shown}`)).toEqual([]);
  });
});
