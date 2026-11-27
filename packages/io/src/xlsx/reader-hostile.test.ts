import JSZip from 'jszip';
import { importXlsx } from './reader';
import { describeHeavy, importInChild, SURVIVED, survived } from './read/hostile/child';
import { MAIN_NS, REL_NS, relsXml, workbookZip } from './read/test-support';
import { DEFAULT_IMPORT_LIMITS } from '../common/safe-zip';

// ---------------------------------------------------------------------------
// Hostile packages that are small as files: each shape was found by review to cost a hang or gigabytes of memory, and must now
// finish within a bounded heap. What is asserted is what the import gives (counts, warnings, loss items), which does not vary by
// machine; time and memory are not read off this process (its heap depends on the Node version and on when the collector ran), but proved
// by a child process with a limited heap and a generous time limit (`importInChild`: it dies of an out-of-memory error or is killed at the
// limit, and the test fails). The shapes that are heavy run only where there is room for them (`describeHeavy`: 6 GB, not `XLSX_SKIP_HOSTILE`).
// A shape that is light is imported here, with a test timeout (a minute) generous enough for a slow machine and short enough for a hang.
// ---------------------------------------------------------------------------

const sheetXml = (body: string, head = '') => `<worksheet ${MAIN_NS}>${head}<sheetData>${body}</sheetData></worksheet>`;
const sheetsOf = (n: number, name: (i: number) => string) =>
  `<sheets>${Array.from({ length: n }, (_, i) => `<sheet name="${name(i)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets>`;
const relsOf = (n: number, part: (i: number) => number) =>
  Array.from({ length: n }, (_, i) => [`rId${i + 1}`, 'worksheet', `worksheets/sheet${part(i)}.xml`] as [string, string, string]);

/** The sheet part of every shape here that has one cell, A1. */
const ONE_CELL = '<row r="1"><c r="A1"><v>1</v></c></row>';

describeHeavy('hostile packages: many sheets that share one cost', () => {
  it('20 sheets that each hold an 8 KB shared master and 2,100 children: the package budget for formula text holds', async () => {
    const master = `${'1+'.repeat(4000)}A1`;
    const children = Array.from({ length: 2_100 }, (_, i) => `<row r="${i + 2}"><c r="A${i + 2}"><f t="shared" si="0"/><v>${i}</v></c></row>`).join('');
    const body = `<row r="1"><c r="A1"><f t="shared" ref="A1:A2101" si="0">${master}</f><v>0</v></c></row>${children}`;
    const parts = Object.fromEntries(Array.from({ length: 20 }, (_, i) => [`xl/worksheets/sheet${i + 1}.xml`, sheetXml(body)]));
    const bytes = await workbookZip(sheetsOf(20, (i) => `S${i}`), relsOf(20, (i) => i + 1), parts);
    const run = importInChild(bytes, { heapMb: 1024, timeoutMs: 120_000 });
    expect(survived(run)).toEqual(SURVIVED);
    expect(run.probe.sheets).toHaveLength(20);
    const chars = run.probe.sheets.reduce((n, s) => n + s.formulaChars, 0);
    // the package budget (64 Mi cost units: a character each, 16 per formula) plus the 20 masters, not 20 sheets' worth (336 million characters)
    expect(chars).toBeLessThan(DEFAULT_IMPORT_LIMITS.maxFormulaChars + 20 * 8_200);
    expect(chars).toBeGreaterThan(DEFAULT_IMPORT_LIMITS.maxFormulaChars / 2);
    expect(run.probe.firstWarnings.filter((w) => w.includes('past the limit'))).not.toHaveLength(0);
    expect(run.probe.limitLoss).toMatchObject([{ id: 'xlsx.package.sizeLimit', label: 'formula(s) not imported', state: 'approximated' }]);
  }, 300_000);

  it('200 <sheet> entries naming one 100,000-cell part: the part is read once, the others are refused, with a warning for the first few and one count', async () => {
    const rows = Array.from({ length: 10_000 }, (_, r) => `<row r="${r + 1}">${Array.from({ length: 10 }, (_, c) => `<c r="${String.fromCharCode(65 + c)}${r + 1}"><v>${r}</v></c>`).join('')}</row>`).join('');
    const bytes = await workbookZip(sheetsOf(200, (i) => `S${i}`), relsOf(200, () => 1), { 'xl/worksheets/sheet1.xml': sheetXml(rows) });
    const run = importInChild(bytes, { heapMb: 512, timeoutMs: 120_000 });
    expect(survived(run)).toEqual(SURVIVED);
    expect(run.probe.sheets).toHaveLength(1);
    expect(run.probe.sheets[0].cells).toBe(100_000);
    expect(run.probe.firstWarnings.filter((w) => /already uses; it was not read/.test(w))).toHaveLength(20);
    expect(run.probe.firstWarnings.filter((w) => /more sheet warnings were left out/.test(w))).toEqual(['179 more sheet warnings were left out.']);
  }, 300_000);
});

describe('hostile packages: many sheets that share one cost (light)', () => {
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
    expect(result.limitLoss).toMatchObject([{ label: 'formula(s) not imported', where: ['S1'] }]);
  });

  it('names the part case-insensitively, as the package finds it', async () => {
    const bytes = await workbookZip(
      '<sheets><sheet name="A" sheetId="1" r:id="rId1"/><sheet name="B" sheetId="2" r:id="rId2"/></sheets>',
      [['rId1', 'worksheet', 'worksheets/sheet1.xml'], ['rId2', 'worksheet', 'worksheets/SHEET1.xml']],
      { 'xl/worksheets/sheet1.xml': sheetXml(ONE_CELL) }
    );
    const result = await importXlsx(bytes);
    expect(result.sheets.map((s) => s.name)).toEqual(['A']);
    expect(result.warnings.filter((w) => w.includes('already uses'))).toHaveLength(1);
  });

  const body = (n: number) => Array.from({ length: n }, (_, r) => `<row r="${r + 1}"><c r="A${r + 1}"><v>${r}</v></c></row>`).join('');
  const threeSheets = (...parts: string[]) =>
    workbookZip(sheetsOf(parts.length, (i) => `S${i}`), relsOf(parts.length, (i) => i + 1), Object.fromEntries(parts.map((p, i) => [`xl/worksheets/sheet${i + 1}.xml`, sheetXml(p)])));

  it('counts cells over all sheets (maxTotalCells): the sheet that passes the limit is cut, the sheets after it that have cells are not kept, one warning and one loss item', async () => {
    const bytes = await threeSheets(body(6), body(6), body(6));
    const result = await importXlsx(bytes, { limits: { maxTotalCells: 10 } });
    expect(result.sheets.map((s) => [s.name, s.cells.size])).toEqual([['S0', 6], ['S1', 4]]);
    expect(result.warnings.filter((w) => w.includes('maxTotalCells'))).toEqual([
      'S1: the workbook has more cells than the import limit allows (maxTotalCells, 10; raise it with the limits option); the rest of this sheet, and of the sheets after it, was not read (4 cells kept here).',
    ]);
    expect(result.limitLoss).toEqual([
      {
        id: 'xlsx.package.sizeLimit',
        label: 'sheet(s) cut at the cell limit',
        count: 2,
        where: ['S1', 'S2'],
        effect: 'lost: the cells past maxTotalCells (10) over the workbook were not imported; raise it with the `limits` option of the import',
      },
    ]);
    // exactly the limit is not over it
    const exact = await importXlsx(bytes, { limits: { maxTotalCells: 18 } });
    expect(exact.sheets.map((s) => s.cells.size)).toEqual([6, 6, 6]);
    expect(exact.warnings).toEqual([]);
    expect(exact.limitLoss).toEqual([]);
  });

  it('a budget that lands exactly on 0 still keeps the later sheets that are empty, and only those', async () => {
    const bytes = await threeSheets(body(6), '', '<row r="1"/>');
    const result = await importXlsx(bytes, { limits: { maxTotalCells: 6 } });
    expect(result.sheets.map((s) => [s.name, s.cells.size])).toEqual([['S0', 6], ['S1', 0], ['S2', 0]]);
    expect(result.warnings).toEqual([]);
    const more = await importXlsx(await threeSheets(body(6), '', body(1)), { limits: { maxTotalCells: 6 } });
    expect(more.sheets.map((s) => s.name)).toEqual(['S0', 'S1']);
    expect(more.limitLoss).toMatchObject([{ label: 'sheet(s) cut at the cell limit', count: 1, where: ['S2'] }]);
    // the name of a sheet that was not kept is free for the next one
    const named = await workbookZip(sheetsOf(3, (i) => (i === 2 ? 'S1' : `S${i}`)), relsOf(3, (i) => i + 1), {
      'xl/worksheets/sheet1.xml': sheetXml(body(6)),
      'xl/worksheets/sheet2.xml': sheetXml(body(1)),
      'xl/worksheets/sheet3.xml': sheetXml(''),
    });
    expect((await importXlsx(named, { limits: { maxTotalCells: 6 } })).sheets.map((s) => s.name)).toEqual(['S0', 'S1']);
  });

  it('the scan budget (maxScannedElements) counts cells that are kept nowhere, rows and columns: the sheet that passes it is cut, one warning and one loss item', async () => {
    const empty = (rows: number, perRow: number) => Array.from({ length: rows }, () => `<row>${'<c/>'.repeat(perRow)}</row>`).join('');
    const bytes = await threeSheets(empty(10, 9), empty(10, 9), empty(10, 9));
    // 100 elements a sheet (10 rows and 90 cells)
    const result = await importXlsx(bytes, { limits: { maxScannedElements: 150 } });
    expect(result.sheets.map((s) => s.name)).toEqual(['S0', 'S1']);
    expect(result.warnings).toEqual([
      'S1: the workbook has more cell, row and column elements than the import limit allows (maxScannedElements, 150; raise it with the limits option); the rest of this sheet, and of the sheets after it, was not read (0 cells kept here).',
    ]);
    expect(result.limitLoss).toMatchObject([{ id: 'xlsx.package.sizeLimit', label: 'sheet(s) cut at the scan limit', count: 2, where: ['S1', 'S2'] }]);
    const cols = `<cols>${'<col min="1" max="1" width="9"/>'.repeat(200)}</cols>`;
    const columns = await importXlsx(await workbookZip(sheetsOf(1, () => 'S'), relsOf(1, () => 1), { 'xl/worksheets/sheet1.xml': sheetXml(ONE_CELL, cols) }), { limits: { maxScannedElements: 100 } });
    expect(columns.limitLoss).toMatchObject([{ label: 'sheet(s) cut at the scan limit', count: 1 }]);
    expect((await importXlsx(bytes, { limits: { maxScannedElements: 300 } })).limitLoss).toEqual([]);
  });

  it('reads the 1,000,000 cells of the performance file (100,000 x 10) under the default limit, with a wide margin', () => {
    expect(DEFAULT_IMPORT_LIMITS.maxTotalCells).toBeGreaterThanOrEqual(5 * 1_000_000);
  });

  it('caps the number of sheets (maxSheets) with one warning and one loss item that counts the sheets left unread', async () => {
    const part = sheetXml(ONE_CELL);
    const parts = Object.fromEntries([1, 2, 3, 4, 5].map((i) => [`xl/worksheets/sheet${i}.xml`, part]));
    const bytes = await workbookZip(sheetsOf(5, (i) => `S${i}`), relsOf(5, (i) => i + 1), parts);
    const result = await importXlsx(bytes, { limits: { maxSheets: 3 } });
    expect(result.sheets).toHaveLength(3);
    expect(result.warnings.filter((w) => w.includes('maxSheets'))).toEqual(['The workbook has more than 3 worksheets (the import limit maxSheets); the others were not read. Raise the limit with the limits option.']);
    expect(result.limitLoss).toMatchObject([{ id: 'xlsx.package.sizeLimit', label: 'worksheet(s) not read', count: 2 }]);
    expect((await importXlsx(bytes, { limits: { maxSheets: 5 } })).sheets).toHaveLength(5);
  });
});

describe('hostile packages: many sheets named alike', () => {
  it('20,000 sheets all named A are numbered in linear time (it was quadratic: 18 s), up to the sheet cap', async () => {
    const part = sheetXml(ONE_CELL);
    // one shared part would be refused after the first; 1,000 distinct parts are read, the rest are over the cap
    const n = 20_000;
    const parts: Record<string, string> = {};
    for (let i = 1; i <= 1_000; i++) parts[`xl/worksheets/sheet${i}.xml`] = part;
    const bytes = await workbookZip(sheetsOf(n, () => 'A'), relsOf(n, (i) => (i % 1_000) + 1), parts);
    const result = await importXlsx(bytes);
    expect(result.sheets).toHaveLength(DEFAULT_IMPORT_LIMITS.maxSheets);
    expect(result.sheets[0].name).toBe('A');
    expect(result.sheets[999].name).toBe('A (1000)');
  }, 120_000);

  it('numbers around names the file already holds, and sheets that differ only by case', async () => {
    const part = sheetXml(ONE_CELL);
    const names = ['A (2)', 'A', 'a', 'A', 'A (2)', 'a (3)', 'a'];
    const bytes = await workbookZip(sheetsOf(names.length, (i) => names[i]), relsOf(names.length, (i) => i + 1), Object.fromEntries(names.map((_, i) => [`xl/worksheets/sheet${i + 1}.xml`, part])));
    const result = await importXlsx(bytes);
    expect(result.sheets.map((s) => s.name)).toEqual(['A (2)', 'A', 'a (3)', 'A (4)', 'A (2) (2)', 'a (3) (2)', 'a (5)']);
  });
});

describe('hostile packages: <cols>', () => {
  const cols = (n: number, attrs: string) => `<cols>${`<col min="1" max="16384" ${attrs}/>`.repeat(n)}</cols>`;
  const lastCell = '<row r="1"><c r="XFD1"><v>1</v></c></row>';
  const one = (part: string) => workbookZip('<sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets>', [['rId1', 'worksheet', 'worksheets/sheet1.xml']], { 'xl/worksheets/sheet1.xml': part });

  it('20,000 full-width <col> elements and one cell at XFD1: near-linear (it took 9 s from a 4 KB zip), only the first 16,384 are read', async () => {
    const result = await importXlsx(await one(sheetXml(lastCell, cols(20_000, 'hidden="1" width="20" customWidth="1"'))));
    const sheet = result.sheets[0];
    expect(sheet.hiddenCols?.size).toBe(16_384);
    expect(sheet.columnWidths?.size).toBe(16_384);
    expect(sheet.details?.columns).toHaveLength(16_384);
    expect(result.warnings).toEqual([expect.stringMatching(/^S: 3616 column range\(s\) were left out \(the limit is 16384 per sheet and maxColumns, 1048576, for the workbook\)/)]);
  }, 60_000);

  it('overlapping ranges: the last one that names an attribute wins, as when each is applied in turn', async () => {
    const body = '<cols><col min="1" max="5" width="10" hidden="1"/><col min="3" max="8" width="20"/><col min="4" max="4" hidden="0"/><col min="7" max="9" hidden="1"/><col min="2" max="2" width="30"/></cols>';
    const sheet = (await importXlsx(await one(sheetXml('<row r="1"><c r="H1"><v>1</v></c></row>', body)))).sheets[0];
    // the extent is column H (7): columns 0..7
    expect([...(sheet.columnWidths as Map<number, number>)]).toEqual([[0, 75], [1, 215], [2, 145], [3, 145], [4, 145], [5, 145], [6, 145], [7, 145]]);
    expect([...(sheet.hiddenCols as Set<number>)]).toEqual([0, 1, 2, 4, 6, 7]);
  });

  it('the columns of all sheets are drawn from one budget (maxColumns): the ranges, and the columns they set up to the last cell', async () => {
    const part = sheetXml('<row r="1"><c r="E1"><v>1</v></c></row>', '<cols><col min="1" max="3" width="9"/><col min="5" max="9" width="9"/></cols>');
    const bytes = await workbookZip(sheetsOf(3, (i) => `S${i}`), relsOf(3, (i) => i + 1), Object.fromEntries([1, 2, 3].map((i) => [`xl/worksheets/sheet${i}.xml`, part])));
    // a sheet costs its 2 ranges and the 5 columns up to E: 7. Budget 16: two sheets, then 2 left for the third sheet's ranges and none of its columns
    const result = await importXlsx(bytes, { limits: { maxColumns: 16 } });
    expect(result.sheets.map((s) => [s.columnWidths?.size, s.details?.columns.length])).toEqual([[4, 2], [4, 2], [0, 2]]);
    expect(result.warnings).toEqual(['S2: the widths and hidden state of 5 column(s) were left out (the import limit maxColumns, 16, for the workbook); the other sheets that lose some are counted in the import report.']);
    expect(result.limitLoss).toMatchObject([{ label: 'column setting(s) left out', count: 5, where: ['S2'] }]);
    // ranges past the budget are left out as they come (the first sheet's two take 2 of 3, and its one column the last)
    const few = await importXlsx(bytes, { limits: { maxColumns: 3 } });
    expect(few.sheets.map((s) => s.details?.columns.length)).toEqual([2, 0, 0]);
    // one warning for the workbook, whatever was cut; the loss item counts the 4 columns of the first sheet and the 4 ranges of the others
    expect(few.warnings.filter((w) => /column/.test(w))).toHaveLength(1);
    expect(few.limitLoss).toMatchObject([{ label: 'column setting(s) left out', count: 8, where: ['S0', 'S1', 'S2'] }]);
  });
});

describeHeavy('hostile packages: <cols> (heavy)', () => {
  it('1,000,000 <col> elements: cut at the cap with one warning, never held', async () => {
    const cols = `<cols>${'<col min="1" max="16384" hidden="1"/>'.repeat(1_000_000)}</cols>`;
    const bytes = await workbookZip('<sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets>', [['rId1', 'worksheet', 'worksheets/sheet1.xml']], {
      'xl/worksheets/sheet1.xml': sheetXml('<row r="1"><c r="XFD1"><v>1</v></c></row>', cols),
    });
    const run = importInChild(bytes, { heapMb: 512, timeoutMs: 120_000 });
    expect(survived(run)).toEqual(SURVIVED);
    expect(run.probe.sheets[0].hiddenCols).toBe(16_384);
    expect(run.probe.firstWarnings.filter((w) => w.includes('column range(s)'))).toHaveLength(1);
  }, 300_000);
});

describe('hostile packages: sheet entries that are not sheets read', () => {
  it('300,000 <sheet> entries naming one part: the entries are capped at 4 x maxSheets, the repeats are counted, not listed (it was 299,999 warnings, 3 s and 310 MB)', async () => {
    const part = sheetXml(ONE_CELL);
    const n = 300_000;
    const bytes = await workbookZip(sheetsOf(n, (i) => `S${i}`), relsOf(4_000, () => 1), { 'xl/worksheets/sheet1.xml': part });
    const result = await importXlsx(bytes);
    expect(result.sheets).toHaveLength(1);
    expect(result.warnings.length).toBeLessThan(60);
    expect(result.warnings.filter((w) => w.includes('already uses'))).toHaveLength(20);
    expect(result.warnings).toContain(`The workbook lists ${n} sheets; only the first ${4 * DEFAULT_IMPORT_LIMITS.maxSheets} entries were read (4 times the import limit maxSheets).`);
    expect(result.limitLoss).toMatchObject([{ label: 'worksheet(s) not read', count: n - 4 * DEFAULT_IMPORT_LIMITS.maxSheets, where: ['the sheet entries of the workbook'] }]);
  }, 120_000);

  it('4,000 sheets naming a part that has 400,000 relationships: its comment parts are looked at once, not once per sheet', async () => {
    const part = sheetXml(ONE_CELL);
    const rels = relsXml(Array.from({ length: 400_000 }, (_, i) => [`x${i}`, 'comments', `c${i}.xml`] as [string, string, string]));
    const bytes = await workbookZip(sheetsOf(4_000, (i) => `S${i}`), relsOf(1, () => 1), { 'xl/worksheets/sheet1.xml': part, 'xl/worksheets/_rels/sheet1.xml.rels': rels });
    const result = await importXlsx(bytes);
    expect(result.sheets).toHaveLength(1);
    expect(result.warnings.filter((w) => w.includes('has comments at'))).toHaveLength(20);
  }, 120_000);

  it('4,000 sheets looked up among 400,000 workbook relationships: by id, not by scanning', async () => {
    const part = sheetXml(ONE_CELL);
    const rels: Array<[string, string, string]> = [...Array.from({ length: 400_000 }, (_, i) => [`x${i}`, 'styles', `s${i}.xml`] as [string, string, string]), ['rIdLast', 'worksheet', 'worksheets/sheet1.xml']];
    const bytes = await workbookZip(sheetsOf(4_000, (i) => `S${i}`).replace(/r:id="rId\d+"/g, 'r:id="rIdLast"'), rels, { 'xl/worksheets/sheet1.xml': part });
    const result = await importXlsx(bytes);
    expect(result.sheets).toHaveLength(1);
  }, 120_000);
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

  it('the same links with their relationships: 65,530 are kept and 34,470 counted as left out', async () => {
    const n = 100_000;
    const links = Array.from({ length: n }, (_, i) => `<hyperlink ref="A${i + 1}" r:id="rId1"/>`).join('');
    const zip = await JSZip.loadAsync(await sheetPackage(`<worksheet ${MAIN_NS} ${REL_NS}><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><hyperlinks>${links}</hyperlinks></worksheet>`));
    zip.file('xl/worksheets/_rels/sheet1.xml.rels', relsXml([['rId1', 'hyperlink', 'https://example.com/', 'External']]));
    const result = await importXlsx(await zip.generateAsync({ type: 'uint8array' }));
    expect(result.sheets[0].details?.hyperlinks).toHaveLength(65_530);
    expect(result.warnings).toEqual([`S: ${n - 65_530} hyperlink(s) were left out (the limit is 65530 per sheet and maxRanges for the workbook).`]);
    expect(result.limitLoss).toMatchObject([{ label: 'merged range(s) and hyperlink(s) left out', count: n - 65_530, where: ['S'] }]);
  }, 60_000);

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
    expect(result.limitLoss).toMatchObject([{ label: 'merged range(s) and hyperlink(s) left out', count: 600, where: ['S1'] }]);
  });

  describeHeavy('heavy', () => {
    it('3,000,000 <mergeCell> and one cell: capped at the sheet limit with one counted warning, in a bounded heap (it was 13.6 s and 1.4 GB)', async () => {
      const run = importInChild(await sheetPackage(mergeCells(3_000_000)), { heapMb: 1024, timeoutMs: 180_000 });
      expect(survived(run)).toEqual(SURVIVED);
      expect(run.probe.sheets[0].merges).toBe(1_048_576);
      expect(run.probe.sheets[0].hyperlinks).toBe(0);
      expect(run.probe.firstWarnings).toEqual([`S: ${3_000_000 - 1_048_576} merged range(s) were left out (the limit is 1048576 per sheet and maxRanges for the workbook).`]);
    }, 400_000);

    it('2,000,000 <hyperlink r:id> with no relationships: 65,530 kept at most, 20 warnings and one count, not two million (it was 9.3 s, 1.7 GB and 2,000,000 warnings)', async () => {
      const links = Array.from({ length: 2_000_000 }, (_, i) => `<hyperlink ref="A${(i % 1_000_000) + 1}" r:id="rId${i + 9}"/>`).join('');
      const bytes = await sheetPackage(`<worksheet ${MAIN_NS} ${REL_NS}><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><hyperlinks>${links}</hyperlinks></worksheet>`);
      const run = importInChild(bytes, { heapMb: 1024, timeoutMs: 180_000 });
      expect(survived(run)).toEqual(SURVIVED);
      expect(run.probe.sheets[0].hyperlinks).toBeLessThanOrEqual(65_530);
      expect(run.probe.warnings).toBeLessThan(30);
      expect(run.probe.firstWarnings.filter((w) => w.startsWith('The hyperlink at'))).toHaveLength(20);
    }, 400_000);

    it('8,000,000 <mergeCell> from a zip under the 100 MB package limit (a part of 250 MB, under every limit of X4): refused or capped, never an out-of-memory', async () => {
      const bytes = await sheetPackage(mergeCells(8_000_000));
      expect(bytes.length).toBeLessThan(DEFAULT_IMPORT_LIMITS.maxPackageBytes);
      const run = importInChild(bytes, { heapMb: 1536, timeoutMs: 300_000 });
      // killed (out of memory, or over the time limit) is the failure: the import either throws its size-limit error (exit 1) or gives the capped merges
      expect(run.signal).toBeNull();
      if (run.probe.error !== undefined) {
        expect(run.status).toBe(1);
        expect(run.probe.errorCode).toBe('size-limit');
      } else {
        expect(survived(run)).toEqual(SURVIVED);
        expect(run.probe.sheets[0].merges).toBe(1_048_576);
      }
    }, 600_000);
  });
});
