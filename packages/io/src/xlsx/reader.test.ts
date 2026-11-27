import { exportXlsx } from './writer';
import { MAIN_NS, relsXml, workbookZip, zipOf } from './read/test-support';
import JSZip from 'jszip';
import { importXlsx } from './reader';
import { ImportError, type ImportErrorCode } from '../common/import-error';
import type { SpreadsheetExportInput } from '@b32nio/spindle-shared/export/xlsx';

async function roundTrip(input: SpreadsheetExportInput) {
  const bytes = await exportXlsx(input);
  return importXlsx(bytes);
}

describe('importXlsx', () => {
  it('round-trips plain values of every supported type', async () => {
    const result = await roundTrip({
      sheets: [
        {
          name: 'Sheet1',
          cells: new Map([
            ['0:0', { value: 'hello' }],
            ['0:1', { value: 42 }],
            ['0:2', { value: true }],
          ]),
        },
      ],
    });
    expect(result.warnings).toEqual([]);
    const cells = result.sheets[0].cells;
    expect(cells.get('0:0')?.value).toBe('hello');
    expect(cells.get('0:1')?.value).toBe(42);
    expect(cells.get('0:2')?.value).toBe(true);
  });

  it('round-trips a formula with its cached result', async () => {
    const result = await roundTrip({
      sheets: [{ name: 'S', cells: new Map([['0:0', { value: 5, formula: 'SUM(A2:A3)' }]]) }],
    });
    const cell = result.sheets[0].cells.get('0:0');
    expect(cell?.formula).toBe('SUM(A2:A3)');
    expect(cell?.value).toBe(5);
  });

  it('round-trips a number format code', async () => {
    const result = await roundTrip({
      sheets: [{ name: 'S', cells: new Map([['0:0', { value: 1234.5, numberFormat: { code: '#,##0.00' } }]]) }],
    });
    expect(result.sheets[0].cells.get('0:0')?.numberFormat?.code).toBe('#,##0.00');
  });

  it('does not set a numberFormat for General (default) cells', async () => {
    const result = await roundTrip({ sheets: [{ name: 'S', cells: new Map([['0:0', { value: 1 }]]) }] });
    expect(result.sheets[0].cells.get('0:0')?.numberFormat).toBeUndefined();
  });

  it('round-trips font/fill/alignment style', async () => {
    const result = await roundTrip({
      sheets: [
        {
          name: 'S',
          cells: new Map([
            [
              '0:0',
              {
                value: 'styled',
                style: {
                  bold: true,
                  italic: true,
                  underline: true,
                  strikethrough: true,
                  fontColor: '#ff0000',
                  backgroundColor: '#00ff00',
                  horizontalAlign: 'center',
                  verticalAlign: 'middle',
                  wrapText: true,
                },
              },
            ],
          ]),
        },
      ],
    });
    expect(result.sheets[0].cells.get('0:0')?.style).toMatchObject({
      bold: true,
      italic: true,
      underline: true,
      strikethrough: true,
      fontColor: '#ff0000',
      backgroundColor: '#00ff00',
      horizontalAlign: 'center',
      verticalAlign: 'middle',
      wrapText: true,
    });
  });

  it('round-trips a border', async () => {
    const result = await roundTrip({
      sheets: [
        { name: 'S', cells: new Map([['0:0', { value: 1, style: { border: { top: { style: 'thick', color: '#123456' } } } }]]) },
      ],
    });
    expect(result.sheets[0].cells.get('0:0')?.style?.border?.top).toEqual({ style: 'thick', color: '#123456' });
  });

  it('round-trips merged ranges', async () => {
    const result = await roundTrip({
      sheets: [
        {
          name: 'S',
          cells: new Map([['0:0', { value: 'merged' }]]),
          mergedRanges: [{ startRow: 0, startCol: 0, endRow: 1, endCol: 1 }],
        },
      ],
    });
    expect(result.sheets[0].mergedRanges).toEqual([{ startRow: 0, startCol: 0, endRow: 1, endCol: 1 }]);
  });

  it('round-trips frozen panes', async () => {
    const result = await roundTrip({ sheets: [{ name: 'S', cells: new Map([['0:0', { value: 1 }]]), frozenRows: 2, frozenCols: 1 }] });
    expect(result.sheets[0].frozenRows).toBe(2);
    expect(result.sheets[0].frozenCols).toBe(1);
  });

  it('round-trips hidden rows/columns', async () => {
    const result = await roundTrip({
      sheets: [
        {
          name: 'S',
          cells: new Map([['0:0', { value: 1 }]]),
          hiddenRows: new Set([0]),
          hiddenCols: new Set([0]),
        },
      ],
    });
    expect(result.sheets[0].hiddenRows?.has(0)).toBe(true);
    expect(result.sheets[0].hiddenCols?.has(0)).toBe(true);
  });

  it('round-trips multiple sheets, in order', async () => {
    const result = await roundTrip({
      sheets: [
        { name: 'First', cells: new Map() },
        { name: 'Second', cells: new Map() },
      ],
    });
    expect(result.sheets.map((s) => s.name)).toEqual(['First', 'Second']);
  });

  it('round-trips a hyperlink', async () => {
    const result = await roundTrip({
      sheets: [{ name: 'S', cells: new Map([['0:0', { value: 'Spindle', hyperlink: 'https://example.com' }]]) }],
    });
    const cell = result.sheets[0].cells.get('0:0');
    expect(cell?.value).toBe('Spindle');
    expect(cell?.hyperlink).toBe('https://example.com');
  });

  it('surfaces a warning list even when empty', async () => {
    const result = await roundTrip({ sheets: [{ name: 'S', cells: new Map() }] });
    expect(result.warnings).toEqual([]);
  });
});

describe('importXlsx: the only error is an ImportError', () => {
  const ole = (...streams: string[]) =>
    Uint8Array.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, ...Array(512).fill(0), ...streams.flatMap((s) => [...Array.from(s).flatMap((c) => [c.charCodeAt(0), 0]), 0, 0])]);
  const zipOf = async (files: Record<string, string>) => {
    const zip = new JSZip();
    for (const [name, content] of Object.entries(files)) zip.file(name, content);
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  };
  const codeOf = async (bytes: Uint8Array, options?: Parameters<typeof importXlsx>[1]): Promise<ImportErrorCode | 'no error'> =>
    importXlsx(bytes, options).then(
      () => 'no error',
      (e: unknown) => {
        expect(e).toBeInstanceOf(ImportError);
        return (e as ImportError).code;
      }
    );

  it.each<[string, () => Promise<Uint8Array> | Uint8Array, ImportErrorCode]>([
    ['an empty file', () => new Uint8Array(0), 'not-a-package'],
    ['text', () => new TextEncoder().encode('a,b,c'), 'not-a-package'],
    ['an encrypted package', () => ole('EncryptionInfo', 'EncryptedPackage'), 'encrypted'],
    ['a legacy .xls', () => ole('Workbook'), 'legacy-binary'],
    ['a zip that is not a workbook', () => zipOf({ 'hello.txt': 'hi' }), 'not-a-package'],
    ['a package with an unreadable workbook part', () => zipOf({ '[Content_Types].xml': '<Types/>', 'xl/workbook.xml': '<<<' }), 'corrupt'],
  ])('%s', async (_name, make, code) => {
    expect(await codeOf(await make())).toBe(code);
  });

  it('refuses a forged or corrupt mimetype entry as corrupt (an ImportError, not a plain Error)', async () => {
    const zip = new JSZip();
    zip.file('[Content_Types].xml', '<Types/>');
    zip.file('mimetype', new Uint8Array(1024 * 1024));
    const bytes = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
    const forged = bytes.slice();
    const view = new DataView(forged.buffer);
    for (let i = 0; i + 30 < forged.length; i++) {
      if (view.getUint32(i, true) === 0x04034b50 && new TextDecoder().decode(forged.subarray(i + 30, i + 38)) === 'mimetype') view.setUint32(i + 22, 10, true);
      if (view.getUint32(i, true) === 0x02014b50 && new TextDecoder().decode(forged.subarray(i + 46, i + 54)) === 'mimetype') view.setUint32(i + 24, 10, true);
    }
    expect(await codeOf(forged)).toBe('corrupt');
  });

  it('refuses a truncated workbook as corrupt', async () => {
    const bytes = await exportXlsx({ sheets: [{ name: 'S', cells: new Map([['0:0', { value: 'x'.repeat(5000) }]]) }] });
    expect(await codeOf(bytes.subarray(0, Math.floor(bytes.length / 2)))).toBe('corrupt');
  });

  it('applies the limits it is given, and the defaults otherwise', async () => {
    const bytes = await exportXlsx({ sheets: [{ name: 'S', cells: new Map([['0:0', { value: 1 }]]) }] });
    expect(await codeOf(bytes)).toBe('no error');
    expect(await codeOf(bytes, { limits: { maxEntries: 2 } })).toBe('size-limit');
    expect(await codeOf(bytes, { limits: { maxPackageBytes: 10 } })).toBe('size-limit');
  });

  it('refuses a zip bomb with a small header declaration without inflating it', async () => {
    const zip = new JSZip();
    zip.file('[Content_Types].xml', '<Types/>');
    zip.file('xl/worksheets/sheet1.xml', new Uint8Array(40 * 1024 * 1024));
    const bomb = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
    expect(bomb.length).toBeLessThan(200 * 1024);
    expect(await codeOf(bomb)).toBe('size-limit');
  });
});

describe('importXlsx: the own reader', () => {
  const sheetXml = (body: string) => `<worksheet ${MAIN_NS}><sheetData>${body}</sheetData></worksheet>`;
  const book = (sheets: string, rels: Array<[string, string, string]>, parts: Record<string, string>) =>
    workbookZip(`<sheets>${sheets}</sheets>`, rels, parts);

  it('reads a workbook through the own reader: sheets, and the workbook, styles, strings and theme beside them', async () => {
    const bytes = await exportXlsx({ sheets: [{ name: 'One', cells: new Map([['0:0', { value: 'hello', style: { bold: true } }]]) }] });
    const result = await importXlsx(bytes);
    expect(result.sheets.map((s) => s.name)).toEqual(['One']);
    expect(result.sheets[0].cells.get('0:0')).toMatchObject({ value: 'hello', style: { bold: true }, sharedString: 0 });
    expect(result.sharedStrings).toEqual([{ text: 'hello' }]);
    expect(result.workbook?.sheets[0]).toMatchObject({ name: 'One', kind: 'worksheet' });
    expect(result.styles?.cellXfs.length).toBeGreaterThan(1);
    expect(result.theme?.colors).toHaveLength(12);
    expect(result.sheets[0].details?.part).toMatch(/worksheets\/sheet1\.xml$/);
  });

  it('refuses a binary workbook (.xlsb) with a message that says what to do', async () => {
    const bytes = await zipOf({ '_rels/.rels': relsXml([['rId1', 'officeDocument', 'xl/book.bin']]), 'xl/book.bin': 'BIFF12' });
    await expect(importXlsx(bytes)).rejects.toMatchObject({ name: 'ImportError', code: 'legacy-binary', message: expect.stringMatching(/\.xlsb.*Save As.*\.xlsx/s) });
  });

  it('refuses a workbook part that is not XML as corrupt, but only skips a sheet that is not', async () => {
    await expect(importXlsx(await zipOf({ '_rels/.rels': relsXml([['rId1', 'officeDocument', 'xl/workbook.xml']]), 'xl/workbook.xml': 'not xml <<<' }))).rejects.toMatchObject({ code: 'corrupt' });
    const bytes = await book('<sheet name="A" sheetId="1" r:id="rId1"/><sheet name="B" sheetId="2" r:id="rId2"/>', [['rId1', 'worksheet', 'worksheets/sheet1.xml'], ['rId2', 'worksheet', 'worksheets/sheet2.xml']], {
      'xl/worksheets/sheet1.xml': '<worksheet><sheetData><row><c><v>1</v></c></row></worksheet>',
      'xl/worksheets/sheet2.xml': sheetXml('<row r="1"><c r="A1"><v>2</v></c></row>'),
    });
    const result = await importXlsx(bytes);
    expect(result.sheets.map((s) => [s.name, s.cells.size, s.details?.wellFormed])).toEqual([['A', 1, false], ['B', 1, true]]);
    expect(result.warnings).toContainEqual(expect.stringContaining('not well-formed'));
  });

  it('reads only worksheets, in order, and skips a sheet whose part is missing, with a warning', async () => {
    const bytes = await book(
      '<sheet name="Ws" sheetId="1" r:id="rId1"/><sheet name="Chart" sheetId="2" r:id="rId2"/><sheet name="Gone" sheetId="3" r:id="rId3"/><sheet name="Last" sheetId="4" r:id="rId4"/>',
      [['rId1', 'worksheet', 'worksheets/sheet1.xml'], ['rId2', 'chartsheet', 'chartsheets/sheet1.xml'], ['rId3', 'worksheet', 'worksheets/gone.xml'], ['rId4', 'worksheet', 'worksheets/sheet2.xml']],
      { 'xl/worksheets/sheet1.xml': sheetXml(''), 'xl/chartsheets/sheet1.xml': `<chartsheet ${MAIN_NS}/>`, 'xl/worksheets/sheet2.xml': sheetXml('') }
    );
    const result = await importXlsx(bytes);
    expect(result.sheets.map((s) => s.name)).toEqual(['Ws', 'Last']);
    expect(result.warnings).toContainEqual(expect.stringContaining('"Gone"'));
  });

  it('gives sheets that clash by name (any case) a numbered name and a warning each', async () => {
    const bytes = await book(
      ['Data', 'data', 'Data', 'Data (2)'].map((n, i) => `<sheet name="${n}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join(''),
      [1, 2, 3, 4].map((i) => [`rId${i}`, 'worksheet', `worksheets/sheet${i}.xml`] as [string, string, string]),
      Object.fromEntries([1, 2, 3, 4].map((i) => [`xl/worksheets/sheet${i}.xml`, sheetXml(`<row r="1"><c r="A1"><v>${i}</v></c></row>`)]))
    );
    const result = await importXlsx(bytes);
    expect(result.sheets.map((s) => s.name)).toEqual(['Data', 'data (2)', 'Data (3)', 'Data (2) (2)']);
    expect(result.sheets.map((s) => s.cells.get('0:0')?.value)).toEqual([1, 2, 3, 4]);
    expect(result.warnings.filter((w) => w.includes('more than one sheet'))).toHaveLength(3);
  });

  it('quotes sheet names, and the part of a repeated sheet, shortened in its warnings', async () => {
    const long = 'L'.repeat(200);
    const bytes = await book(
      [long, long.toLowerCase(), 'Third'].map((n, i) => `<sheet name="${n}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join(''),
      [['rId1', 'worksheet', 'worksheets/sheet1.xml'], ['rId2', 'worksheet', 'worksheets/sheet2.xml'], ['rId3', 'worksheet', `worksheets/${'x'.repeat(300)}.xml`]],
      {
        'xl/worksheets/sheet1.xml': sheetXml('<row r="1"><c r="A1"><v>1</v></c></row>'),
        'xl/worksheets/sheet2.xml': sheetXml('<row r="1"><c r="A1"><v>2</v></c></row>'),
        [`xl/worksheets/${'x'.repeat(300)}.xml`]: sheetXml('<row r="1"><c r="A1"><v>3</v></c></row>'),
      }
    );
    const repeated = await book(
      [`<sheet name="${long}" sheetId="1" r:id="rId1"/>`, `<sheet name="${long}2" sheetId="2" r:id="rId2"/>`].join(''),
      [['rId1', 'worksheet', 'worksheets/sheet1.xml'], ['rId2', 'worksheet', 'worksheets/SHEET1.xml']],
      { 'xl/worksheets/sheet1.xml': sheetXml('<row r="1"><c r="A1"><v>1</v></c></row>') }
    );
    const result = await importXlsx(bytes);
    expect(result.warnings).toContain(`The workbook has more than one sheet named "${'l'.repeat(63)}…"; one was renamed "${'l'.repeat(63)}…".`);
    expect(result.sheets).toHaveLength(3);
    expect((await importXlsx(repeated)).warnings).toEqual([`The sheet "${'L'.repeat(63)}…" names the part xl/worksheets/sheet1.xml, which another sheet already uses; it was not read.`].map((w) => w));
  });

  it('keeps the whole warning text small for a 10 MB sheet name', async () => {
    const bytes = await book(`<sheet name="${'N'.repeat(10_000_000)}" sheetId="1" r:id="rId1"/><sheet name="${'N'.repeat(10_000_000)}" sheetId="2" r:id="rId2"/>`, [['rId1', 'worksheet', 'worksheets/sheet1.xml'], ['rId2', 'worksheet', 'worksheets/sheet2.xml']], {
      'xl/worksheets/sheet1.xml': sheetXml('<row r="1"><c r="A1" t="d"><v>x</v></c></row>'),
      'xl/worksheets/sheet2.xml': sheetXml('<row r="1"><c r="A1" t="d"><v>y</v></c></row>'),
    });
    const result = await importXlsx(bytes);
    expect(result.sheets.map((s) => s.name.length)).toEqual([255, 255 + ' (2)'.length]);
    expect(result.warnings.reduce((n, w) => n + w.length, 0)).toBeLessThan(2_000);
  });

  it('takes the budget for rows with settings from the limits (maxRows), over all the sheets', async () => {
    const rows = Array.from({ length: 10 }, (_, i) => `<row r="${i + 1}" ht="20"/>`).join('');
    const bytes = await book('<sheet name="A" sheetId="1" r:id="rId1"/><sheet name="B" sheetId="2" r:id="rId2"/>', [['rId1', 'worksheet', 'worksheets/sheet1.xml'], ['rId2', 'worksheet', 'worksheets/sheet2.xml']], {
      'xl/worksheets/sheet1.xml': sheetXml(rows),
      'xl/worksheets/sheet2.xml': sheetXml(rows),
    });
    const result = await importXlsx(bytes, { limits: { maxRows: 12 } });
    expect(result.sheets.map((s) => s.details?.rows.length)).toEqual([10, 2]);
    expect(result.warnings).toEqual(['B: the settings of 8 row(s) were left out (the import limit maxRows).']);
    expect((await importXlsx(bytes)).warnings).toEqual([]);
  });

  it('reads the defined names of a workbook as a list under a cap (one warning, one loss item), refuses a workbook whose other elements are over the budget with an error that says "too many elements", and reads a styles part over it as a skipped part', async () => {
    const many = `<sheet name="S" sheetId="1" r:id="rId1"/></sheets><definedNames>${'<definedName name="n">A1</definedName>'.repeat(20)}</definedNames><sheets>`;
    const bytes = await book(many, [['rId1', 'worksheet', 'worksheets/sheet1.xml']], { 'xl/worksheets/sheet1.xml': sheetXml('<row r="1"><c r="A1"><v>1</v></c></row>') });
    const capped = await importXlsx(bytes, { limits: { maxPartElements: 10 } });
    expect(capped.workbook?.definedNames).toHaveLength(10);
    expect(capped.warnings).toEqual(['The workbook has more than 10 defined names (the import limit maxPartElements); 10 were not read.']);
    expect(capped.limitLoss).toMatchObject([{ id: 'xlsx.package.sizeLimit', label: 'defined name(s) not kept', count: 10 }]);
    const crowded = await book(`${'<sheet name="S" sheetId="1" r:id="rId1"/>'.repeat(20)}`, [['rId1', 'worksheet', 'worksheets/sheet1.xml']], { 'xl/worksheets/sheet1.xml': sheetXml('<row r="1"><c r="A1"><v>1</v></c></row>') });
    await expect(importXlsx(crowded, { limits: { maxPartElements: 10 } })).rejects.toMatchObject({
      code: 'corrupt',
      message: 'The workbook part of the file has too many elements (more than the import limit maxPartElements, 10; raise it with the limits option).',
    });
    expect((await importXlsx(bytes)).sheets).toHaveLength(1);
    const styled = await book('<sheet name="S" sheetId="1" r:id="rId1"/>', [['rId1', 'worksheet', 'worksheets/sheet1.xml'], ['rId2', 'styles', 'styles.xml']], {
      'xl/worksheets/sheet1.xml': sheetXml('<row r="1"><c r="A1" s="1"><v>1</v></c></row>'),
      'xl/styles.xml': `<styleSheet ${MAIN_NS}><cellXfs>${'<xf/>'.repeat(20)}</cellXfs></styleSheet>`,
    });
    const result = await importXlsx(styled, { limits: { maxPartElements: 10 } });
    expect(result.styles?.cellXfs).toEqual([]);
    expect(result.warnings).toEqual(['The part xl/styles.xml has more than 10 elements (the import limit maxPartElements) and was skipped.', 'S: 1 cell(s) use a format the file does not have; imported unformatted.']);
    expect(result.limitLoss).toMatchObject([{ label: 'part(s) over the element limit', where: ['xl/styles.xml'] }]);
  });

  it('reads the workbook part once: a workbook part that is damaged is refused as corrupt, with a message that does not blame the element budget', async () => {
    const zip = await zipOf({ '_rels/.rels': relsXml([['rId1', 'officeDocument', 'xl/workbook.xml']]), 'xl/workbook.xml': '<workbook><sheets>' });
    await expect(importXlsx(zip)).rejects.toMatchObject({ code: 'corrupt', message: 'The workbook part of the file is damaged (it is not well-formed XML).' });
  });

  it('carries what the limits cut as loss items (limitLoss), empty for a file that is not cut, and the same ids and wording on every kind', async () => {
    const part = sheetXml('<row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row>');
    const bytes = await book('<sheet name="S" sheetId="1" r:id="rId1"/><sheet name="T" sheetId="2" r:id="rId2"/>', [['rId1', 'worksheet', 'worksheets/sheet1.xml'], ['rId2', 'worksheet', 'worksheets/sheet2.xml']], {
      'xl/worksheets/sheet1.xml': part,
      'xl/worksheets/sheet2.xml': part,
    });
    expect((await importXlsx(bytes)).limitLoss).toEqual([]);
    const cut = await importXlsx(bytes, { limits: { maxTotalCells: 3 } });
    expect(cut.sheets.map((s) => s.cells.size)).toEqual([2, 1]);
    expect(cut.limitLoss).toEqual([
      {
        id: 'xlsx.package.sizeLimit',
        label: 'sheet(s) cut at the cell limit',
        count: 1,
        where: ['T'],
        effect: 'lost: the cells past maxTotalCells (3) over the workbook were not imported; raise it with the `limits` option of the import',
      },
    ]);
  });

  it('warns once of a number format code over 255 characters, for every sheet and format that uses it, and keeps it as written', async () => {
    const code = `yyyy${'0'.repeat(300)}`;
    const files = {
      'xl/worksheets/sheet1.xml': sheetXml('<row r="1"><c r="A1" s="1"><v>1</v></c><c r="B1" s="2"><v>2</v></c></row>'),
      'xl/worksheets/sheet2.xml': sheetXml('<row r="1"><c r="A1" s="1"><v>1</v></c></row>'),
      'xl/styles.xml': `<styleSheet ${MAIN_NS}><numFmts><numFmt numFmtId="164" formatCode="${code}"/></numFmts><cellXfs><xf/><xf numFmtId="164"/><xf numFmtId="164" bold="1"/></cellXfs></styleSheet>`,
    };
    const rels: Array<[string, string, string]> = [['rId1', 'worksheet', 'worksheets/sheet1.xml'], ['rId2', 'worksheet', 'worksheets/sheet2.xml'], ['rId3', 'styles', 'styles.xml']];
    const result = await importXlsx(await book('<sheet name="S" sheetId="1" r:id="rId1"/><sheet name="T" sheetId="2" r:id="rId2"/>', rels, files));
    expect(result.warnings).toEqual(['A number format code is over 255 characters (Excel allows no more); it is kept as written and not read as a date or time.']);
    expect(result.sheets[0].cells.get('0:0')?.numberFormat).toEqual({ code });
    expect(result.sheets[1].cells.get('0:0')?.numberFormat).toEqual({ code });
    expect(result.limitLoss).toMatchObject([{ label: 'number format code(s) over 255 characters', count: 1, state: 'approximated' }]);
  });

  it('records the shared strings past the cap as a loss, and a damaged shared-strings part only as a warning', async () => {
    const files = (strings: string) => ({ 'xl/worksheets/sheet1.xml': sheetXml('<row r="1"><c r="A1" t="s"><v>0</v></c></row>'), 'xl/sharedStrings.xml': `<sst ${MAIN_NS}>${strings}</sst>` });
    const rels: Array<[string, string, string]> = [['rId1', 'worksheet', 'worksheets/sheet1.xml'], ['rId2', 'sharedStrings', 'sharedStrings.xml']];
    const many = await importXlsx(await book('<sheet name="S" sheetId="1" r:id="rId1"/>', rels, files('<si><t>a</t></si><si><t>b</t></si><si><t>c</t></si>')), { limits: { maxTotalCells: 2 } });
    expect(many.sharedStrings).toHaveLength(2);
    expect(many.limitLoss).toMatchObject([{ label: 'part(s) of the shared strings not read', where: ['xl/sharedStrings.xml'] }]);
    const damaged = await importXlsx(await book('<sheet name="S" sheetId="1" r:id="rId1"/>', rels, files('<si><t>a</t></si><si><t>b')));
    expect(damaged.warnings.some((w) => w.includes('is not well-formed XML; it was read as far as the damage'))).toBe(true);
    expect(damaged.limitLoss).toEqual([]);
  });

  it.each(['__proto__', 'constructor', 'toString', 'hasOwnProperty'])('a sheet named %s is an ordinary sheet', async (name) => {
    const bytes = await book(`<sheet name="${name}" sheetId="1" r:id="rId1"/><sheet name="${name}" sheetId="2" r:id="rId2"/>`, [['rId1', 'worksheet', 'worksheets/sheet1.xml'], ['rId2', 'worksheet', 'worksheets/sheet2.xml']], {
      'xl/worksheets/sheet1.xml': sheetXml('<row r="1"><c r="A1"><v>1</v></c></row>'),
      'xl/worksheets/sheet2.xml': sheetXml('<row r="1"><c r="A1"><v>2</v></c></row>'),
    });
    const result = await importXlsx(bytes);
    expect(result.sheets.map((s) => s.name)).toEqual([name, `${name} (2)`]);
    expect(Object.keys(Object.prototype)).toEqual([]);
  });

  it('reports the active tab of the workbook, and repeats no warning', async () => {
    const bytes = await workbookZip(`<bookViews><workbookView activeTab="1"/></bookViews><sheets><sheet name="A" sheetId="1" r:id="rId1"/><sheet name="B" sheetId="2" r:id="rId2"/></sheets>`, [['rId1', 'worksheet', 'worksheets/sheet1.xml'], ['rId2', 'worksheet', 'worksheets/sheet2.xml']], {
      'xl/worksheets/sheet1.xml': sheetXml('<row r="1"><c r="A1" t="s"><v>5</v></c></row>'),
      'xl/worksheets/sheet2.xml': sheetXml('<row r="1"><c r="A1" t="s"><v>5</v></c></row>'),
    });
    const result = await importXlsx(bytes);
    expect(result.activeSheetIndex).toBe(1);
    expect(new Set(result.warnings).size).toBe(result.warnings.length);
  });

  describe('the active sheet when sheets are skipped', () => {
    const sheets = '<sheet name="Chart" sheetId="1" r:id="rId1"/><sheet name="A" sheetId="2" r:id="rId2"/><sheet name="Gone" sheetId="3" r:id="rId3"/><sheet name="B" sheetId="4" r:id="rId4"/><sheet name="Chart2" sheetId="5" r:id="rId5"/>';
    const rels: Array<[string, string, string]> = [
      ['rId1', 'chartsheet', 'chartsheets/sheet1.xml'], ['rId2', 'worksheet', 'worksheets/sheet1.xml'], ['rId3', 'worksheet', 'worksheets/gone.xml'], ['rId4', 'worksheet', 'worksheets/sheet2.xml'], ['rId5', 'chartsheet', 'chartsheets/sheet2.xml'],
    ];
    const parts = { 'xl/worksheets/sheet1.xml': sheetXml(''), 'xl/worksheets/sheet2.xml': sheetXml(''), 'xl/chartsheets/sheet1.xml': `<chartsheet ${MAIN_NS}/>`, 'xl/chartsheets/sheet2.xml': `<chartsheet ${MAIN_NS}/>` };
    const active = async (activeTab: string | undefined, over: { sheets?: string; rels?: Array<[string, string, string]>; parts?: Record<string, string> } = {}) => {
      const view = activeTab === undefined ? '' : `<bookViews><workbookView activeTab="${activeTab}"/></bookViews>`;
      const bytes = await workbookZip(`${view}<sheets>${over.sheets ?? sheets}</sheets>`, over.rels ?? rels, over.parts ?? parts);
      const result = await importXlsx(bytes);
      return { index: result.activeSheetIndex, names: result.sheets.map((s) => s.name) };
    };

    it('maps the tab to the sheet it is, among the sheets that were read', async () => {
      expect(await active('1')).toEqual({ index: 0, names: ['A', 'B'] });
      expect(await active('3')).toEqual({ index: 1, names: ['A', 'B'] });
    });

    it('moves a tab that is a chart sheet or a sheet without a part to the next sheet that was read', async () => {
      expect((await active('0')).index).toBe(0);
      expect((await active('2')).index).toBe(1);
    });

    it('moves a tab past the last sheet read, or past the end, to the last sheet read; a negative one to the first', async () => {
      expect((await active('4')).index).toBe(1);
      expect((await active('99')).index).toBe(1);
      expect((await active('-3')).index).toBe(0);
    });

    it('leaves it unset when the workbook says nothing, and is 0 when no sheet was read', async () => {
      expect((await active(undefined)).index).toBeUndefined();
      const onlyChart = await active('0', { sheets: '<sheet name="Chart" sheetId="1" r:id="rId1"/>', rels: [rels[0]], parts: { 'xl/chartsheets/sheet1.xml': parts['xl/chartsheets/sheet1.xml'] } });
      expect(onlyChart).toEqual({ index: 0, names: [] });
    });
  });

  it('takes the budget for expanded shared formulas from the limits', async () => {
    const children = Array.from({ length: 10 }, (_, i) => `<row r="${i + 2}"><c r="A${i + 2}"><f t="shared" si="0"/><v>1</v></c></row>`).join('');
    const bytes = await book('<sheet name="A" sheetId="1" r:id="rId1"/>', [['rId1', 'worksheet', 'worksheets/sheet1.xml']], {
      'xl/worksheets/sheet1.xml': sheetXml(`<row r="1"><c r="A1"><f t="shared" ref="A1:A11" si="0">B1+1</f><v>1</v></c></row>${children}`),
    });
    const formulas = async (limits?: { maxSheetFormulaChars?: number; maxFormulaChars?: number }) => {
      const result = await importXlsx(bytes, limits === undefined ? {} : { limits });
      return { count: [...result.sheets[0].cells.values()].filter((c) => c.formula !== undefined).length, warnings: result.warnings };
    };
    expect(await formulas()).toEqual({ count: 11, warnings: [] });
    // a formula costs its length (4) and FORMULA_OVERHEAD (16)
    const limited = await formulas({ maxSheetFormulaChars: 3 * 20 });
    expect(limited.count).toBe(1 + 3);
    expect(limited.warnings).toEqual(['A: formulas left out, their values kept: 7 past the limit of 60 characters of expanded shared formulas.']);
    // the package's budget is the same cost over all its sheets
    const packaged = await formulas({ maxFormulaChars: 3 * 20 });
    expect(packaged.count).toBe(1 + 3);
    expect(packaged.warnings).toHaveLength(1);
  });

  it('applies the limits to every sheet part it reads', async () => {
    const bytes = await book('<sheet name="A" sheetId="1" r:id="rId1"/>', [['rId1', 'worksheet', 'worksheets/sheet1.xml']], { 'xl/worksheets/sheet1.xml': sheetXml(`<row r="1">${'<c r="A1"><v>1</v></c>'.repeat(1000)}</row>`) });
    await expect(importXlsx(bytes, { limits: { maxXmlPartBytes: 1000 } })).rejects.toMatchObject({ code: 'size-limit' });
  });
});
