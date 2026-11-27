import JSZip from 'jszip';
import { importXlsx } from './reader';
import { MAIN_NS, REL_NS, relsXml, workbookZip } from './read/test-support';
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

  it('200 <sheet> entries naming one 100,000-cell part: the part is read once, the others are refused, with a warning for the first few and one count', async () => {
    const rows = Array.from({ length: 10_000 }, (_, r) => `<row r="${r + 1}">${Array.from({ length: 10 }, (_, c) => `<c r="${String.fromCharCode(65 + c)}${r + 1}"><v>${r}</v></c>`).join('')}</row>`).join('');
    const bytes = await workbookZip(sheetsOf(200, (i) => `S${i}`), relsOf(200, () => 1), { 'xl/worksheets/sheet1.xml': sheetXml(rows) });
    const { result, ms, heap } = await timed(() => importXlsx(bytes));
    expect(result.sheets).toHaveLength(1);
    expect(result.sheets[0].cells.size).toBe(100_000);
    expect(result.warnings.filter((w) => /already uses; it was not read/.test(w))).toHaveLength(20);
    expect(result.warnings.filter((w) => /more sheet warnings were left out/.test(w))).toEqual(['179 more sheet warnings were left out.']);
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

describe('hostile packages: sheet entries that are not sheets read', () => {
  it('300,000 <sheet> entries naming one part: the entries are capped at 4 x maxSheets, the repeats are counted, not listed (it was 299,999 warnings, 3 s and 310 MB)', async () => {
    const part = sheetXml('<row r="1"><c r="A1"><v>1</v></c></row>');
    const n = 300_000;
    const bytes = await workbookZip(sheetsOf(n, (i) => `S${i}`), relsOf(4_000, () => 1), { 'xl/worksheets/sheet1.xml': part });
    const { result, ms, heap } = await timed(() => importXlsx(bytes));
    expect(result.sheets).toHaveLength(1);
    expect(result.warnings.length).toBeLessThan(60);
    expect(result.warnings.filter((w) => w.includes('already uses'))).toHaveLength(20);
    expect(result.warnings).toContain(`The workbook lists ${n} sheets; only the first ${4 * DEFAULT_IMPORT_LIMITS.maxSheets} entries were read (4 times the import limit maxSheets).`);
    expect(ms).toBeLessThan(5_000);
    expect(heap).toBeLessThan(300 * MB);
  }, 60_000);

  it('4,000 sheets naming a part that has 400,000 relationships: its comment parts are looked at once, not once per sheet', async () => {
    const part = sheetXml('<row r="1"><c r="A1"><v>1</v></c></row>');
    const rels = relsXml(Array.from({ length: 400_000 }, (_, i) => [`x${i}`, 'comments', `c${i}.xml`] as [string, string, string]));
    const bytes = await workbookZip(sheetsOf(4_000, (i) => `S${i}`), relsOf(1, () => 1), { 'xl/worksheets/sheet1.xml': part, 'xl/worksheets/_rels/sheet1.xml.rels': rels });
    const { result, ms } = await timed(() => importXlsx(bytes));
    expect(result.sheets).toHaveLength(1);
    expect(result.warnings.filter((w) => w.includes('has comments at'))).toHaveLength(20);
    expect(ms).toBeLessThan(10_000);
  }, 60_000);

  it('4,000 sheets looked up among 400,000 workbook relationships: by id, not by scanning', async () => {
    const part = sheetXml('<row r="1"><c r="A1"><v>1</v></c></row>');
    const rels: Array<[string, string, string]> = [...Array.from({ length: 400_000 }, (_, i) => [`x${i}`, 'styles', `s${i}.xml`] as [string, string, string]), ['rIdLast', 'worksheet', 'worksheets/sheet1.xml']];
    const bytes = await workbookZip(sheetsOf(4_000, (i) => `S${i}`).replace(/r:id="rId\d+"/g, 'r:id="rIdLast"'), rels, { 'xl/worksheets/sheet1.xml': part });
    const { result, ms } = await timed(() => importXlsx(bytes));
    expect(result.sheets).toHaveLength(1);
    expect(ms).toBeLessThan(10_000);
  }, 60_000);
});

describe('hostile packages: merges and hyperlinks (the shapes of the S2b review)', () => {
  const sheetPackage = async (part: string): Promise<Uint8Array> => {
    const zip = new JSZip();
    zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>');
    zip.file('_rels/.rels', relsXml([['rId1', 'officeDocument', 'xl/workbook.xml']]));
    zip.file('xl/workbook.xml', `<workbook ${MAIN_NS} ${REL_NS}><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`);
    zip.file('xl/_rels/workbook.xml.rels', relsXml([['rId1', 'worksheet', 'worksheets/sheet1.xml']]));
    zip.file('xl/worksheets/sheet1.xml', part);
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 1 } });
  };
  /** `n` merges of pseudo-random rows (a fixed generator): they compress about 15 to 1, as a real file of distinct ranges does, not to nothing. */
  const mergeCells = (n: number): string => {
    const items: string[] = new Array<string>(n);
    let seed = 12345;
    for (let i = 0; i < n; i++) {
      seed = (Math.imul(seed, 1103515245) + 12345) & 0x7fffffff;
      const row = 10_000 + (seed % 90_000);
      items[i] = `<mergeCell ref="A${row}:B${row + 1}"/>`;
    }
    return `<worksheet ${MAIN_NS}><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><mergeCells count="${n}">${items.join('')}</mergeCells></worksheet>`;
  };

  it('3,000,000 <mergeCell> and one cell: capped at the sheet limit with one counted warning, in seconds and with a bounded heap (it was 13.6 s and 1.4 GB)', async () => {
    const bytes = await sheetPackage(mergeCells(3_000_000));
    const { result, ms, heap } = await timed(() => importXlsx(bytes));
    const sheet = result.sheets[0];
    expect(sheet.mergedRanges).toHaveLength(1_048_576);
    expect(sheet.details?.hyperlinks).toEqual([]);
    expect(result.warnings).toEqual([`S: ${3_000_000 - 1_048_576} merged range(s) were left out (the limit is 1048576 per sheet and maxRanges for the workbook).`]);
    expect(ms).toBeLessThan(15_000);
    expect(heap).toBeLessThan(800 * MB);
  }, 90_000);

  it('2,000,000 <hyperlink r:id> with no relationships: 65,530 kept at most, 20 warnings and one count, not two million (it was 9.3 s, 1.7 GB and 2,000,000 warnings)', async () => {
    const links = Array.from({ length: 2_000_000 }, (_, i) => `<hyperlink ref="A${(i % 1_000_000) + 1}" r:id="rId${i + 9}"/>`).join('');
    const bytes = await sheetPackage(`<worksheet ${MAIN_NS} ${REL_NS}><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><hyperlinks>${links}</hyperlinks></worksheet>`);
    const { result, ms, heap } = await timed(() => importXlsx(bytes));
    expect(result.sheets[0].details?.hyperlinks.length).toBeLessThanOrEqual(65_530);
    expect(result.warnings.length).toBeLessThan(30);
    expect(result.warnings.filter((w) => w.startsWith('The hyperlink at'))).toHaveLength(20);
    expect(ms).toBeLessThan(15_000);
    expect(heap).toBeLessThan(800 * MB);
  }, 90_000);

  it('the same links with their relationships: 65,530 are kept and 1,934,470 counted as left out', async () => {
    const n = 100_000;
    const links = Array.from({ length: n }, (_, i) => `<hyperlink ref="A${i + 1}" r:id="rId1"/>`).join('');
    const zip = await JSZip.loadAsync(await sheetPackage(`<worksheet ${MAIN_NS} ${REL_NS}><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><hyperlinks>${links}</hyperlinks></worksheet>`));
    zip.file('xl/worksheets/_rels/sheet1.xml.rels', relsXml([['rId1', 'hyperlink', 'https://example.com/', 'External']]));
    const result = await importXlsx(await zip.generateAsync({ type: 'uint8array' }));
    expect(result.sheets[0].details?.hyperlinks).toHaveLength(65_530);
    expect(result.warnings).toEqual([`S: ${n - 65_530} hyperlink(s) were left out (the limit is 65530 per sheet and maxRanges for the workbook).`]);
  }, 60_000);

  it('8,000,000 <mergeCell> from a zip under the 100 MB package limit (a part of 250 MB, under every limit of X4): refused or capped within a minute, the heap bounded, never an out-of-memory', async () => {
    const bytes = await sheetPackage(mergeCells(8_000_000));
    expect(bytes.length).toBeLessThan(DEFAULT_IMPORT_LIMITS.maxPackageBytes);
    const { result, ms, heap } = await timed(() => importXlsx(bytes).then((r) => ({ r }), (e: unknown) => ({ e })));
    if ('e' in result) expect((result.e as { code: string }).code).toBe('size-limit');
    else expect(result.r.sheets[0].mergedRanges).toHaveLength(1_048_576);
    expect(ms).toBeLessThan(60_000);
    expect(heap).toBeLessThan(1_200 * MB);
  }, 180_000);

  it('a package budget (maxRanges) holds over all sheets: two sheets of 400 merges and 400 links under a budget of 1,000', async () => {
    const body = `<worksheet ${MAIN_NS} ${REL_NS}><sheetData/><mergeCells>${Array.from({ length: 400 }, (_, i) => `<mergeCell ref="A${i + 1}:B${i + 1}"/>`).join('')}</mergeCells><hyperlinks>${Array.from({ length: 400 }, (_, i) => `<hyperlink ref="C${i + 1}" location="A1"/>`).join('')}</hyperlinks></worksheet>`;
    const bytes = await workbookZip(sheetsOf(2, (i) => `S${i}`), relsOf(2, (i) => i + 1), { 'xl/worksheets/sheet1.xml': body, 'xl/worksheets/sheet2.xml': body });
    const result = await importXlsx(bytes, { limits: { maxRanges: 1_000 } });
    const kept = result.sheets.map((s) => [s.mergedRanges?.length ?? 0, s.details?.hyperlinks.length ?? 0]);
    expect(kept).toEqual([[400, 400], [200, 0]]);
    expect(result.warnings).toEqual([
      'S1: 200 merged range(s) were left out (the limit is 1048576 per sheet and maxRanges for the workbook).',
      'S1: 400 hyperlink(s) were left out (the limit is 65530 per sheet and maxRanges for the workbook).',
    ]);
  });
});
