import { importXlsx } from './reader';
import { MAIN_NS, workbookZip } from './read/test-support';
import { DEFAULT_IMPORT_LIMITS } from '../common/safe-zip';

// ---------------------------------------------------------------------------
// Hostile packages that are small as files: each shape was found by review to cost a hang or gigabytes of memory, and must now
// finish in seconds within a bounded heap. The checks are on time and on what the import keeps, not on exact memory (that varies
// by machine); each stays far below what the shape used to cost.
// ---------------------------------------------------------------------------

const sheetXml = (body: string, head = '') => `<worksheet ${MAIN_NS}>${head}<sheetData>${body}</sheetData></worksheet>`;
const sheetsOf = (n: number, name: (i: number) => string) =>
  `<sheets>${Array.from({ length: n }, (_, i) => `<sheet name="${name(i)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets>`;
const relsOf = (n: number, part: (i: number) => number) =>
  Array.from({ length: n }, (_, i) => [`rId${i + 1}`, 'worksheet', `worksheets/sheet${part(i)}.xml`] as [string, string, string]);

const timed = async <T>(run: () => Promise<T>): Promise<{ result: T; ms: number; heap: number }> => {
  const heap = process.memoryUsage().heapUsed;
  const t0 = Date.now();
  const result = await run();
  return { result, ms: Date.now() - t0, heap: process.memoryUsage().heapUsed - heap };
};

const MB = 1024 * 1024;

describe('hostile packages: many sheets that share one cost', () => {
  it('20 sheets that each hold an 8 KB shared master and 2,100 children: the package budget for formula text holds', async () => {
    const master = `${'1+'.repeat(4000)}A1`;
    const children = Array.from({ length: 2_100 }, (_, i) => `<row r="${i + 2}"><c r="A${i + 2}"><f t="shared" si="0"/><v>${i}</v></c></row>`).join('');
    const body = `<row r="1"><c r="A1"><f t="shared" ref="A1:A2101" si="0">${master}</f><v>0</v></c></row>${children}`;
    const parts = Object.fromEntries(Array.from({ length: 20 }, (_, i) => [`xl/worksheets/sheet${i + 1}.xml`, sheetXml(body)]));
    const bytes = await workbookZip(sheetsOf(20, (i) => `S${i}`), relsOf(20, (i) => i + 1), parts);
    const { result, ms, heap } = await timed(() => importXlsx(bytes));
    expect(result.sheets).toHaveLength(20);
    const chars = result.sheets.reduce((n, s) => n + [...s.cells.values()].reduce((m, c) => m + (c.formula?.length ?? 0), 0), 0);
    // the package budget (64 Mi cost units: a character each, 16 per formula) plus the 20 masters, not 20 sheets' worth (336 million characters)
    expect(chars).toBeLessThan(DEFAULT_IMPORT_LIMITS.maxFormulaChars + 20 * 8_200);
    expect(chars).toBeGreaterThan(DEFAULT_IMPORT_LIMITS.maxFormulaChars / 2);
    expect(result.warnings.filter((w) => w.includes('past the limit'))).not.toHaveLength(0);
    expect(ms).toBeLessThan(15_000);
    expect(heap).toBeLessThan(800 * MB);
  }, 60_000);

  it('the package budget is tunable through limits (maxFormulaChars)', async () => {
    const children = Array.from({ length: 100 }, (_, i) => `<row r="${i + 2}"><c r="A${i + 2}"><f t="shared" si="0"/><v>${i}</v></c></row>`).join('');
    const body = `<row r="1"><c r="A1"><f t="shared" ref="A1:A101" si="0">B1+1</f><v>0</v></c></row>${children}`;
    const bytes = await workbookZip(sheetsOf(2, (i) => `S${i}`), relsOf(2, (i) => i + 1), { 'xl/worksheets/sheet1.xml': sheetXml(body), 'xl/worksheets/sheet2.xml': sheetXml(body) });
    // 4 to 6 characters and 16 of overhead per formula: about 140 children fit, the first sheet takes 100 of them
    const result = await importXlsx(bytes, { limits: { maxFormulaChars: 150 * 20 } });
    const counts = result.sheets.map((s) => [...s.cells.values()].filter((c) => c.formula !== undefined).length);
    expect(counts[0]).toBe(101);
    expect(counts[1]).toBeGreaterThan(30);
    expect(counts[1]).toBeLessThan(60);
  });

  it('200 <sheet> entries naming one 100,000-cell part: the part is read once, the others are refused with a warning each', async () => {
    const rows = Array.from({ length: 10_000 }, (_, r) => `<row r="${r + 1}">${Array.from({ length: 10 }, (_, c) => `<c r="${String.fromCharCode(65 + c)}${r + 1}"><v>${r}</v></c>`).join('')}</row>`).join('');
    const bytes = await workbookZip(sheetsOf(200, (i) => `S${i}`), relsOf(200, () => 1), { 'xl/worksheets/sheet1.xml': sheetXml(rows) });
    const { result, ms, heap } = await timed(() => importXlsx(bytes));
    expect(result.sheets).toHaveLength(1);
    expect(result.sheets[0].cells.size).toBe(100_000);
    expect(result.warnings.filter((w) => /already uses; it was not read/.test(w))).toHaveLength(199);
    expect(ms).toBeLessThan(5_000);
    expect(heap).toBeLessThan(300 * MB);
  }, 30_000);

  it('names the part case-insensitively, as the package finds it', async () => {
    const bytes = await workbookZip(
      '<sheets><sheet name="A" sheetId="1" r:id="rId1"/><sheet name="B" sheetId="2" r:id="rId2"/></sheets>',
      [['rId1', 'worksheet', 'worksheets/sheet1.xml'], ['rId2', 'worksheet', 'worksheets/SHEET1.xml']],
      { 'xl/worksheets/sheet1.xml': sheetXml('<row r="1"><c r="A1"><v>1</v></c></row>') }
    );
    const result = await importXlsx(bytes);
    expect(result.sheets.map((s) => s.name)).toEqual(['A']);
    expect(result.warnings.filter((w) => w.includes('already uses'))).toHaveLength(1);
  });

  it('counts cells over all sheets (maxTotalCells): the sheet that passes the limit is cut, the sheets after it are not read', async () => {
    const body = (n: number) => Array.from({ length: n }, (_, r) => `<row r="${r + 1}"><c r="A${r + 1}"><v>${r}</v></c></row>`).join('');
    const bytes = await workbookZip(sheetsOf(3, (i) => `S${i}`), relsOf(3, (i) => i + 1), {
      'xl/worksheets/sheet1.xml': sheetXml(body(6)),
      'xl/worksheets/sheet2.xml': sheetXml(body(6)),
      'xl/worksheets/sheet3.xml': sheetXml(body(6)),
    });
    const result = await importXlsx(bytes, { limits: { maxTotalCells: 10 } });
    expect(result.sheets.map((s) => s.cells.size)).toEqual([6, 4]);
    expect(result.warnings.filter((w) => w.includes('maxTotalCells'))).toHaveLength(2);
    // exactly the limit is not over it
    const exact = await importXlsx(bytes, { limits: { maxTotalCells: 18 } });
    expect(exact.sheets.map((s) => s.cells.size)).toEqual([6, 6, 6]);
    expect(exact.warnings).toEqual([]);
  });

  it('reads the 1,000,000 cells of the performance file (100,000 x 10) under the default limit, with a wide margin', () => {
    expect(DEFAULT_IMPORT_LIMITS.maxTotalCells).toBeGreaterThanOrEqual(5 * 1_000_000);
  });

  it('caps the number of sheets (maxSheets) with one warning', async () => {
    const part = sheetXml('<row r="1"><c r="A1"><v>1</v></c></row>');
    const parts = Object.fromEntries([1, 2, 3, 4].map((i) => [`xl/worksheets/sheet${i}.xml`, part]));
    const bytes = await workbookZip(sheetsOf(4, (i) => `S${i}`), relsOf(4, (i) => i + 1), parts);
    const result = await importXlsx(bytes, { limits: { maxSheets: 3 } });
    expect(result.sheets).toHaveLength(3);
    expect(result.warnings.filter((w) => w.includes('maxSheets'))).toHaveLength(1);
    expect((await importXlsx(bytes, { limits: { maxSheets: 4 } })).sheets).toHaveLength(4);
  });
});

describe('hostile packages: many sheets named alike', () => {
  it('20,000 sheets all named A are numbered in linear time (it was quadratic: 18 s), up to the sheet cap', async () => {
    const part = sheetXml('<row r="1"><c r="A1"><v>1</v></c></row>');
    // one shared part would be refused after the first; 1,000 distinct parts are read, the rest are over the cap
    const n = 20_000;
    const parts: Record<string, string> = {};
    for (let i = 1; i <= 1_000; i++) parts[`xl/worksheets/sheet${i}.xml`] = part;
    const bytes = await workbookZip(sheetsOf(n, () => 'A'), relsOf(n, (i) => (i % 1_000) + 1), parts);
    const { result, ms } = await timed(() => importXlsx(bytes));
    expect(result.sheets).toHaveLength(DEFAULT_IMPORT_LIMITS.maxSheets);
    expect(result.sheets[0].name).toBe('A');
    expect(result.sheets[999].name).toBe('A (1000)');
    expect(ms).toBeLessThan(10_000);
  }, 60_000);

  it('numbers around names the file already holds, and sheets that differ only by case', async () => {
    const part = sheetXml('<row r="1"><c r="A1"><v>1</v></c></row>');
    const names = ['A (2)', 'A', 'a', 'A', 'A (2)', 'a (3)', 'a'];
    const bytes = await workbookZip(sheetsOf(names.length, (i) => names[i]), relsOf(names.length, (i) => i + 1), Object.fromEntries(names.map((_, i) => [`xl/worksheets/sheet${i + 1}.xml`, part])));
    const result = await importXlsx(bytes);
    expect(result.sheets.map((s) => s.name)).toEqual(['A (2)', 'A', 'a (3)', 'A (4)', 'A (2) (2)', 'a (3) (2)', 'a (5)']);
  });
});

describe('hostile packages: <cols>', () => {
  const cols = (n: number, attrs: string) => `<cols>${`<col min="1" max="16384" ${attrs}/>`.repeat(n)}</cols>`;
  const lastCell = '<row r="1"><c r="XFD1"><v>1</v></c></row>';

  it('20,000 full-width <col> elements and one cell at XFD1: near-linear (it took 9 s from a 4 KB zip), only the first 16,384 are read', async () => {
    const bytes = await workbookZip('<sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets>', [['rId1', 'worksheet', 'worksheets/sheet1.xml']], {
      'xl/worksheets/sheet1.xml': sheetXml(lastCell, cols(20_000, 'hidden="1" width="20" customWidth="1"')),
    });
    const { result, ms } = await timed(() => importXlsx(bytes));
    const sheet = result.sheets[0];
    expect(sheet.hiddenCols?.size).toBe(16_384);
    expect(sheet.columnWidths?.size).toBe(16_384);
    expect(sheet.details?.columns).toHaveLength(16_384);
    expect(result.warnings.some((w) => w.includes('20000 column ranges; only the first 16384 were read'))).toBe(true);
    expect(ms).toBeLessThan(3_000);
  });

  it('1,000,000 <col> elements: cut at the cap with one warning, never held', async () => {
    const bytes = await workbookZip('<sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets>', [['rId1', 'worksheet', 'worksheets/sheet1.xml']], {
      'xl/worksheets/sheet1.xml': sheetXml(lastCell, cols(1_000_000, 'hidden="1"')),
    });
    const { result, ms, heap } = await timed(() => importXlsx(bytes));
    expect(result.sheets[0].hiddenCols?.size).toBe(16_384);
    expect(result.warnings.filter((w) => w.includes('column ranges'))).toHaveLength(1);
    expect(ms).toBeLessThan(5_000);
    expect(heap).toBeLessThan(300 * MB);
  }, 30_000);

  it('overlapping ranges: the last one that names an attribute wins, as when each is applied in turn', async () => {
    const body = '<cols><col min="1" max="5" width="10" hidden="1"/><col min="3" max="8" width="20"/><col min="4" max="4" hidden="0"/><col min="7" max="9" hidden="1"/><col min="2" max="2" width="30"/></cols>';
    const bytes = await workbookZip('<sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets>', [['rId1', 'worksheet', 'worksheets/sheet1.xml']], {
      'xl/worksheets/sheet1.xml': sheetXml('<row r="1"><c r="H1"><v>1</v></c></row>', body),
    });
    const sheet = (await importXlsx(bytes)).sheets[0];
    // the extent is column H (7): columns 0..7
    expect([...(sheet.columnWidths as Map<number, number>)]).toEqual([[0, 75], [1, 215], [2, 145], [3, 145], [4, 145], [5, 145], [6, 145], [7, 145]]);
    expect([...(sheet.hiddenCols as Set<number>)]).toEqual([0, 1, 2, 4, 6, 7]);
  });
});
