import { ImportError } from '../../common/import-error';
import { XlsxPackage } from './package';
import { findWorkbookPart, isBinaryWorkbook, MAX_SHEET_NAME_CHARS, readWorkbook, SHEET_ENTRY_FACTOR } from './workbook';
import { MAIN_NS, REL_NS, relsXml, workbookZip, zipOf } from './test-support';

const WB_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml';
const types = (...overrides: Array<[string, string]>) =>
  `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/>${overrides
    .map(([part, type]) => `<Override PartName="${part}" ContentType="${type}"/>`)
    .join('')}</Types>`;

const read = async (bytes: Uint8Array) => {
  const pkg = await XlsxPackage.open(bytes);
  return { info: await readWorkbook(pkg), pkg };
};

describe('isBinaryWorkbook', () => {
  it('is true for a .bin workbook part under any name (the conventional xl/workbook.bin is already refused by openSafeZip) and for the binary content type, and false for an XML one', async () => {
    const bin = await XlsxPackage.open(await zipOf({ '_rels/.rels': relsXml([['rId1', 'officeDocument', 'xl/book.bin']]), 'xl/book.bin': 'BIFF12' }));
    const part = await findWorkbookPart(bin);
    expect(part).toBe('xl/book.bin');
    expect(isBinaryWorkbook(bin, part)).toBe(true);
    const typed = await XlsxPackage.open(
      await zipOf({
        '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/xl/book" ContentType="application/vnd.ms-excel.sheet.binary.macroEnabled.main"/></Types>',
        'xl/book': 'x',
      }),
    );
    expect(isBinaryWorkbook(typed, 'xl/book')).toBe(true);
    const xml = await XlsxPackage.open(await zipOf({ 'xl/workbook.xml': '<workbook/>' }));
    expect(isBinaryWorkbook(xml, 'xl/workbook.xml')).toBe(false);
    expect(isBinaryWorkbook(xml, 'no-extension-or-type')).toBe(false);
  });
});

describe('findWorkbookPart', () => {
  const book = `<workbook ${MAIN_NS}/>`;
  it('follows the officeDocument relationship, absolute or relative, case-insensitively', async () => {
    for (const target of ['xl/book.xml', '/xl/book.xml', 'XL/Book.xml']) {
      const pkg = await XlsxPackage.open(await zipOf({ '_rels/.rels': relsXml([['rId1', 'officeDocument', target]]), 'xl/book.xml': book }));
      expect(await findWorkbookPart(pkg)).toBe('xl/book.xml');
    }
  });
  it('accepts the strict relationship type', async () => {
    const pkg = await XlsxPackage.open(
      await zipOf({ '_rels/.rels': relsXml([['rId1', 'http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument', 'xl/book.xml']]), 'xl/book.xml': book }),
    );
    expect(await findWorkbookPart(pkg)).toBe('xl/book.xml');
  });
  it('skips relationships that are external or point nowhere, then uses the content types', async () => {
    const pkg = await XlsxPackage.open(
      await zipOf({
        '[Content_Types].xml': types(['/ghost.xml', WB_TYPE], ['/XL/Book.xml', WB_TYPE]),
        '_rels/.rels': relsXml([['rId1', 'officeDocument', 'http://x/y.xml', 'External'], ['rId2', 'officeDocument', 'xl/gone.xml']]),
        'xl/book.xml': book,
      }),
    );
    expect(await findWorkbookPart(pkg)).toBe('xl/book.xml');
  });
  it.each([
    'application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml',
    'application/vnd.ms-excel.sheet.macroEnabled.main+xml',
    'application/vnd.ms-excel.template.macroEnabled.main+xml',
  ])('recognises the workbook content type %s', async (type) => {
    const pkg = await XlsxPackage.open(await zipOf({ '[Content_Types].xml': types(['/xl/book.xml', type]), 'xl/book.xml': book }));
    expect(await findWorkbookPart(pkg)).toBe('xl/book.xml');
  });
  it('falls back to xl/workbook.xml', async () => {
    const pkg = await XlsxPackage.open(await zipOf({ 'xl/workbook.xml': book }));
    expect(await findWorkbookPart(pkg)).toBe('xl/workbook.xml');
  });
  it('is a corrupt-file error when there is no workbook', async () => {
    const pkg = await XlsxPackage.open(await zipOf({ 'docProps/core.xml': '<a/>' }));
    await expect(findWorkbookPart(pkg)).rejects.toMatchObject({ code: 'corrupt', message: expect.stringContaining('no workbook part') });
    await expect(findWorkbookPart(pkg)).rejects.toBeInstanceOf(ImportError);
  });
});

describe('readWorkbook: sheets', () => {
  it('reads names, ids, states, kinds and parts (relative and absolute targets)', async () => {
    const { info, pkg } = await read(
      await workbookZip(
        `<sheets>
          <sheet name="Data" sheetId="1" r:id="rId1"/>
          <sheet name="Hid" sheetId="2" state="hidden" r:id="rId2"/>
          <sheet name="VeryHid" sheetId="7" state="veryHidden" r:id="rId3"/>
          <sheet name="Odd" sheetId="9" state="strange" r:id="rId4"/>
          <sheet name="Chart" sheetId="10" r:id="rId5"/>
          <sheet name="Macro" sheetId="11" r:id="rId6"/>
          <sheet name="Intl" sheetId="12" r:id="rId7"/>
          <sheet name="Dialog" sheetId="13" r:id="rId8"/>
          <sheet name="Other" sheetId="14" r:id="rId9"/>
        </sheets>`,
        [
          ['rId1', 'worksheet', 'worksheets/sheet1.xml'],
          ['rId2', 'worksheet', '/xl/worksheets/sheet2.xml'],
          ['rId3', 'worksheet', 'worksheets/sheet3.xml'],
          ['rId4', 'worksheet', 'worksheets/sheet4.xml'],
          ['rId5', 'chartsheet', 'chartsheets/sheet1.xml'],
          ['rId6', 'xlMacrosheet', 'macrosheets/sheet1.xml'],
          ['rId7', 'xlIntlMacrosheet', 'macrosheets/sheet2.xml'],
          ['rId8', 'dialogsheet', 'dialogsheets/sheet1.xml'],
          ['rId9', 'customThing', 'other/sheet1.xml'],
        ],
        Object.fromEntries(
          ['worksheets/sheet1', 'worksheets/sheet2', 'worksheets/sheet3', 'worksheets/sheet4', 'chartsheets/sheet1', 'macrosheets/sheet1', 'macrosheets/sheet2', 'dialogsheets/sheet1', 'other/sheet1'].map((p) => [`xl/${p}.xml`, '<x/>']),
        ),
      ),
    );
    expect(info.sheets.map((s) => [s.name, s.sheetId, s.state, s.kind, s.part])).toEqual([
      ['Data', 1, 'visible', 'worksheet', 'xl/worksheets/sheet1.xml'],
      ['Hid', 2, 'hidden', 'worksheet', 'xl/worksheets/sheet2.xml'],
      ['VeryHid', 7, 'veryHidden', 'worksheet', 'xl/worksheets/sheet3.xml'],
      ['Odd', 9, 'visible', 'worksheet', 'xl/worksheets/sheet4.xml'],
      ['Chart', 10, 'visible', 'chartsheet', 'xl/chartsheets/sheet1.xml'],
      ['Macro', 11, 'visible', 'macrosheet', 'xl/macrosheets/sheet1.xml'],
      ['Intl', 12, 'visible', 'macrosheet', 'xl/macrosheets/sheet2.xml'],
      ['Dialog', 13, 'visible', 'dialogsheet', 'xl/dialogsheets/sheet1.xml'],
      ['Other', 14, 'visible', 'customThing', 'xl/other/sheet1.xml'],
    ]);
    expect(info.sheets[0].relId).toBe('rId1');
    expect(pkg.warnings).toEqual([]);
  });
  it('keeps repeated names and gives a nameless or id-less sheet a default', async () => {
    const { info } = await read(
      await workbookZip(
        '<sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Data" sheetId="2" r:id="rId1"/><sheet r:id="rId1"/><sheet name="X" sheetId="oops" r:id="rId1"/></sheets>',
        [['rId1', 'worksheet', 'worksheets/sheet1.xml']],
        { 'xl/worksheets/sheet1.xml': '<x/>' },
      ),
    );
    expect(info.sheets.map((s) => [s.name, s.sheetId])).toEqual([['Data', 1], ['Data', 2], ['Sheet3', 3], ['X', 4]]);
  });
  it('warns for a sheet without a usable relationship or whose part is missing, and keeps the sheet', async () => {
    const { info, pkg } = await read(
      await workbookZip(
        '<sheets><sheet name="A" sheetId="1" r:id="rId9"/><sheet name="B" sheetId="2"/><sheet name="C" sheetId="3" r:id="rId2"/><sheet name="D" sheetId="4" r:id="rId3"/></sheets>',
        [['rId2', 'worksheet', 'worksheets/gone.xml'], ['rId3', 'worksheet', 'http://x/y', 'External']],
      ),
    );
    expect(info.sheets.map((s) => [s.name, s.part, s.rels, s.kind])).toEqual([
      ['A', undefined, [], 'worksheet'],
      ['B', undefined, [], 'worksheet'],
      ['C', undefined, [], 'worksheet'],
      ['D', undefined, [], 'worksheet'],
    ]);
    expect(pkg.warnings).toEqual([
      'The sheet "A" has no usable relationship (rId9).',
      'The sheet "B" has no usable relationship (no r:id).',
      'The sheet "C" points at xl/worksheets/gone.xml, which the package does not contain.',
      'The sheet "D" has no usable relationship (rId3).',
    ]);
    expect(info.sheets[1].relId).toBe('');
  });
  it('words the warning for a sheet without a name too', async () => {
    const { pkg } = await read(await workbookZip('<sheets><sheet sheetId="1"/><sheet sheetId="2" r:id="rId2"/></sheets>', [['rId2', 'worksheet', 'gone.xml']]));
    // the sheet is named as it will be (Sheet1, Sheet2), as the warning for a sheet with a name does
    expect(pkg.warnings).toEqual(['The sheet "Sheet1" has no usable relationship (no r:id).', 'The sheet "Sheet2" points at xl/gone.xml, which the package does not contain.']);
  });
  it.each(['constructor', '__proto__', 'toString', 'hasOwnProperty'])('a relationship type ending in the inherited key %s keeps that text as the sheet kind', async (key) => {
    const { info } = await read(
      await workbookZip('<sheets><sheet name="A" sheetId="1" r:id="rId1"/></sheets>', [['rId1', `http://example.com/rel/${key}`, 'worksheets/sheet1.xml']], { 'xl/worksheets/sheet1.xml': '<x/>' }),
    );
    expect(info.sheets[0].kind).toBe(key);
  });
  it('maps the macro sheet kinds and keeps chart and dialog sheets', async () => {
    const { info } = await read(
      await workbookZip(
        '<sheets><sheet name="A" sheetId="1" r:id="rId1"/><sheet name="B" sheetId="2" r:id="rId2"/><sheet name="C" sheetId="3" r:id="rId3"/><sheet name="D" sheetId="4" r:id="rId4"/></sheets>',
        [['rId1', 'chartsheet', 's1.xml'], ['rId2', 'dialogsheet', 's1.xml'], ['rId3', 'XlMacroSheet', 's1.xml'], ['rId4', 'xlIntlMacroSheet', 's1.xml']],
        { 'xl/s1.xml': '<x/>' },
      ),
    );
    expect(info.sheets.map((s) => s.kind)).toEqual(['chartsheet', 'dialogsheet', 'macrosheet', 'macrosheet']);
  });
  it('warns when a comments, VML or threaded-comment target is missing from the package', async () => {
    const { info, pkg } = await read(
      await workbookZip('<sheets><sheet name="A" sheetId="1" r:id="rId1"/><sheet r:id="rId1"/></sheets>', [['rId1', 'worksheet', 'worksheets/sheet1.xml']], {
        'xl/worksheets/sheet1.xml': '<x/>',
        'xl/worksheets/_rels/sheet1.xml.rels': relsXml([['rId1', 'comments', '../gone-c.xml'], ['rId2', 'vmlDrawing', '../gone.vml'], ['rId3', 'hyperlink', 'http://x', 'External']]),
      }),
    );
    expect(info.sheets[0].commentParts).toEqual({ threadedComments: [] });
    // a part is looked at once, however many sheet entries name it: the first entry's name is the one in the warning
    expect(pkg.warnings).toEqual([
      'The sheet "A" has comments at xl/gone-c.xml, which the package does not contain.',
      'The sheet "A" has vmlDrawing at xl/gone.vml, which the package does not contain.',
    ]);
    expect(info.sheets[1].commentParts).toBe(info.sheets[0].commentParts);
  });

  describe('sheet entries are bounded', () => {
    const entries = (n: number, rel: (i: number) => string) => Array.from({ length: n }, (_, i) => `<sheet name="S${i}" sheetId="${i + 1}" r:id="${rel(i)}"/>`).join('');

    it('reads at most 4 times maxSheets entries and says so once', async () => {
      const bytes = await workbookZip(`<sheets>${entries(30, () => 'rId1')}</sheets>`, [['rId1', 'worksheet', 'worksheets/sheet1.xml']], { 'xl/worksheets/sheet1.xml': '<x/>' });
      const pkg = await XlsxPackage.open(bytes, { limits: { maxSheets: 5 } });
      const info = await readWorkbook(pkg);
      expect(info.sheets).toHaveLength(20);
      expect(info.sheets.map((s) => s.name).slice(-1)).toEqual(['S19']);
      expect(pkg.warnings).toEqual(['The workbook lists 30 sheets; only the first 20 entries were read (4 times the import limit maxSheets).']);
      expect(SHEET_ENTRY_FACTOR).toBe(4);
    });

    it('does not warn at exactly the limit, and the package carries the limits it was opened with', async () => {
      const bytes = await workbookZip(`<sheets>${entries(20, () => 'rId1')}</sheets>`, [['rId1', 'worksheet', 'worksheets/sheet1.xml']], { 'xl/worksheets/sheet1.xml': '<x/>' });
      const pkg = await XlsxPackage.open(bytes, { limits: { maxSheets: 5 } });
      expect((await readWorkbook(pkg)).sheets).toHaveLength(20);
      expect(pkg.warnings).toEqual([]);
      expect(pkg.limits.maxSheets).toBe(5);
      const defaults = (await XlsxPackage.open(bytes)).limits;
      expect([defaults.maxSheets, defaults.maxRanges]).toEqual([1_000, 1_048_576]);
    });

    it('gives the first few warnings in full and counts the rest: 100 entries with no relationship', async () => {
      const bytes = await workbookZip(`<sheets>${entries(100, (i) => `rId${i + 5}`)}</sheets>`, [['rId1', 'worksheet', 'worksheets/sheet1.xml']], { 'xl/worksheets/sheet1.xml': '<x/>' });
      const pkg = await XlsxPackage.open(bytes);
      await readWorkbook(pkg);
      expect(pkg.warnings).toHaveLength(21);
      expect(pkg.warnings[0]).toBe('The sheet "S0" has no usable relationship (rId5).');
      expect(pkg.warnings[20]).toBe('80 more sheet warnings were left out.');
    });

    it('finds the relationship of an entry by id (the first of a repeated id), not by scanning them', async () => {
      const bytes = await workbookZip(
        `<sheets><sheet name="A" sheetId="1" r:id="rId1"/><sheet name="B" sheetId="2" r:id="constructor"/><sheet name="C" sheetId="3" r:id="__proto__"/></sheets>`,
        [['rId1', 'worksheet', 'worksheets/sheet1.xml'], ['rId1', 'worksheet', 'worksheets/other.xml'], ['constructor', 'worksheet', 'worksheets/sheet2.xml']],
        { 'xl/worksheets/sheet1.xml': '<x/>', 'xl/worksheets/sheet2.xml': '<x/>' }
      );
      const { info } = await read(bytes);
      expect(info.sheets.map((s) => s.part)).toEqual(['xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml', undefined]);
    });
  });
  it('locates the comment parts of a sheet: notes, their VML drawing and threaded comments; external ones and missing parts are not located', async () => {
    const { info } = await read(
      await workbookZip('<sheets><sheet name="A" sheetId="1" r:id="rId1"/><sheet name="B" sheetId="2" r:id="rId2"/></sheets>', [['rId1', 'worksheet', 'worksheets/sheet1.xml'], ['rId2', 'worksheet', 'worksheets/sheet2.xml']], {
        'xl/worksheets/sheet1.xml': '<x/>',
        'xl/worksheets/sheet2.xml': '<x/>',
        'xl/worksheets/_rels/sheet1.xml.rels': relsXml([
          ['rId1', 'comments', '../comments1.xml'],
          ['rId2', 'vmlDrawing', '/xl/drawings/vmlDrawing1.vml'],
          ['rId3', 'http://schemas.microsoft.com/office/2017/10/relationships/threadedComment', '../threadedComments/threadedComment1.xml'],
          ['rId4', 'http://schemas.microsoft.com/office/2017/10/relationships/threadedComment', '../threadedComments/threadedComment2.xml'],
          ['rId5', 'hyperlink', 'https://example.com', 'External'],
        ]),
        'xl/worksheets/_rels/sheet2.xml.rels': relsXml([['rId1', 'comments', 'http://x/c.xml', 'External'], ['rId2', 'comments', '../missing.xml']]),
        'xl/comments1.xml': '<comments/>',
        'xl/drawings/vmlDrawing1.vml': '<xml/>',
        'xl/threadedComments/threadedComment1.xml': '<t/>',
        'xl/threadedComments/threadedComment2.xml': '<t/>',
      }),
    );
    expect(info.sheets[0].commentParts).toEqual({
      comments: 'xl/comments1.xml',
      vmlDrawing: 'xl/drawings/vmlDrawing1.vml',
      threadedComments: ['xl/threadedComments/threadedComment1.xml', 'xl/threadedComments/threadedComment2.xml'],
    });
    expect(info.sheets[0].rels.map((r) => r.id)).toEqual(['rId1', 'rId2', 'rId3', 'rId4', 'rId5']);
    expect(info.sheets[1].commentParts).toEqual({ threadedComments: [] });
  });
});

describe('readWorkbook: workbook-level properties', () => {
  it('reads definedNames: global, local, hidden, built-in, commented; a leading = is dropped; nameless entries skipped', async () => {
    const { info } = await read(
      await workbookZip(
        `<definedNames>
          <definedName name="Total">Data!$B$2:$B$5</definedName>
          <definedName name="Rate" localSheetId="1" hidden="1" comment="Interest">=0.05</definedName>
          <definedName name="_xlnm.Print_Area" localSheetId="0">Data!$A$1:$C$9</definedName>
          <definedName name="Shown" hidden="0"> Data!$A$1 </definedName>
          <definedName name="">x</definedName>
          <definedName>y</definedName>
          <definedName name="Empty"/>
        </definedNames>`,
        [],
      ),
    );
    expect(info.definedNames).toEqual([
      { name: 'Total', ref: 'Data!$B$2:$B$5', builtin: false },
      { name: 'Rate', ref: '0.05', builtin: false, localSheetId: 1, hidden: true, comment: 'Interest' },
      { name: '_xlnm.Print_Area', ref: 'Data!$A$1:$C$9', builtin: true, localSheetId: 0 },
      { name: 'Shown', ref: 'Data!$A$1', builtin: false, hidden: false },
      { name: 'Empty', ref: '', builtin: false },
    ]);
  });
  it('reads calcPr', async () => {
    const { info } = await read(
      await workbookZip(
        '<calcPr calcId="191029" calcMode="manual" fullCalcOnLoad="1" refMode="R1C1" iterate="1" iterateCount="50" iterateDelta="0.01" fullPrecision="0" calcOnSave="0" concurrentCalc="1" forceFullCalc="true"/>',
        [],
      ),
    );
    expect(info.calc).toEqual({
      calcId: 191029,
      calcMode: 'manual',
      fullCalcOnLoad: true,
      refMode: 'R1C1',
      iterate: true,
      iterateCount: 50,
      iterateDelta: 0.01,
      fullPrecision: false,
      calcOnSave: false,
      concurrentCalc: true,
      forceFullCalc: true,
    });
  });
  it('ignores unknown calcMode and refMode values, and has no calc properties when calcPr is absent', async () => {
    expect((await read(await workbookZip('<calcPr calcMode="weird" refMode="Z9"/>', []))).info.calc).toEqual({});
    expect((await read(await workbookZip('', []))).info.calc).toEqual({});
  });
  it.each([
    ['<workbookPr date1904="1"/>', true],
    ['<workbookPr date1904="true"/>', true],
    ['<workbookPr date1904="0"/>', false],
    ['<workbookPr/>', false],
    ['', false],
  ])('date1904 from %s is %s', async (body, expected) => expect((await read(await workbookZip(body, []))).info.date1904).toBe(expected));
  it('reads codeName, appName and the active tab and first sheet', async () => {
    const { info } = await read(await workbookZip('<fileVersion appName="Calc"/><workbookPr codeName="ThisWorkbook"/><bookViews><workbookView activeTab="2" firstSheet="1"/></bookViews>', []));
    expect(info).toMatchObject({ appName: 'Calc', codeName: 'ThisWorkbook', activeTab: 2, firstSheet: 1, part: 'xl/workbook.xml' });
    const none = (await read(await workbookZip('', []))).info;
    expect(none.activeTab).toBeUndefined();
    expect(none.appName).toBeUndefined();
  });
  it('an unparseable workbook part gives no sheets and a warning', async () => {
    const { info, pkg } = await read(await zipOf({ 'xl/workbook.xml': '<workbook><sheets' }));
    expect(info.sheets).toEqual([]);
    expect(info.part).toBe('xl/workbook.xml');
    expect(pkg.warnings).toContain('The part xl/workbook.xml is not well-formed XML and was skipped.');
  });
});

describe('readWorkbook: locating the shared parts', () => {
  it('follows relationships first', async () => {
    const { info } = await read(
      await workbookZip('', [['rId1', 'sharedStrings', 'strings/s.xml'], ['rId2', 'styles', '/xl/st.xml'], ['rId3', 'theme', 'theme/t.xml'], ['rId4', 'http://schemas.microsoft.com/office/2017/10/relationships/person', 'persons/person.xml']], {
        'xl/strings/s.xml': '<x/>', 'xl/st.xml': '<x/>', 'xl/theme/t.xml': '<x/>', 'xl/persons/person.xml': '<x/>',
      }),
    );
    expect(info.parts).toEqual({ sharedStrings: 'xl/strings/s.xml', styles: 'xl/st.xml', theme: 'xl/theme/t.xml', persons: 'xl/persons/person.xml' });
  });
  it('falls back to the content types, then to the conventional names, and leaves absent parts out', async () => {
    const byType = await read(
      await workbookZip('', [], {
        '[Content_Types].xml': types(
          ['/xl/ss.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml'],
          ['/xl/sty.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'],
          ['/xl/th.xml', 'application/vnd.openxmlformats-officedocument.theme+xml'],
          ['/xl/pp.xml', 'application/vnd.ms-excel.person+xml'],
          ['/xl/ghost.xml', 'application/vnd.ms-excel.person+xml'],
        ),
        'xl/ss.xml': '<x/>', 'xl/sty.xml': '<x/>', 'xl/th.xml': '<x/>', 'xl/pp.xml': '<x/>',
      }),
    );
    expect(byType.info.parts).toEqual({ sharedStrings: 'xl/ss.xml', styles: 'xl/sty.xml', theme: 'xl/th.xml', persons: 'xl/pp.xml' });
    const byName = await read(await workbookZip('', [['rId1', 'styles', 'http://x/s.xml', 'External'], ['rId2', 'sharedStrings', 'gone.xml']], { 'xl/styles.xml': '<x/>', 'xl/sharedStrings.xml': '<x/>', 'xl/theme/theme1.xml': '<x/>' }));
    expect(byName.info.parts).toEqual({ sharedStrings: 'xl/sharedStrings.xml', styles: 'xl/styles.xml', theme: 'xl/theme/theme1.xml' });
    expect((await read(await workbookZip('', []))).info.parts).toEqual({});
  });
});

describe('readWorkbook: namespaces', () => {
  it('reads a workbook with prefixed main-namespace elements and one without a namespace', async () => {
    const prefixed = await zipOf({
      '_rels/.rels': relsXml([['rId1', 'officeDocument', 'xl/workbook.xml']]),
      'xl/workbook.xml': `<x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ${REL_NS}><x:sheets><x:sheet name="P" sheetId="1" r:id="rId1"/></x:sheets></x:workbook>`,
      'xl/_rels/workbook.xml.rels': relsXml([['rId1', 'worksheet', 'worksheets/s.xml']]),
      'xl/worksheets/s.xml': '<x/>',
    });
    expect((await read(prefixed)).info.sheets[0]).toMatchObject({ name: 'P', part: 'xl/worksheets/s.xml' });
    const bare = await zipOf({ 'xl/workbook.xml': '<workbook><sheets><sheet name="N" sheetId="3"/></sheets></workbook>' });
    expect((await read(bare)).info.sheets[0]).toMatchObject({ name: 'N', sheetId: 3 });
  });
});

describe('readWorkbook: sheet names and what the warnings quote', () => {
  const one = (name: string, extra = '', rels: Array<[string, string, string]> = [['rId1', 'worksheet', 'worksheets/sheet1.xml']]) =>
    workbookZip(`<sheets><sheet name="${name}" sheetId="1" r:id="rId1"${extra}/></sheets>`, rels, { 'xl/worksheets/sheet1.xml': '<worksheet/>' });

  it('keeps a name of up to MAX_SHEET_NAME_CHARS as it is (Excel\'s 31 is not enforced), with no warning', async () => {
    const name = 'n'.repeat(MAX_SHEET_NAME_CHARS);
    const { info, pkg } = await read(await one(name));
    expect(info.sheets[0].name).toBe(name);
    expect(pkg.warnings).toEqual([]);
  });

  it('cuts a longer name to MAX_SHEET_NAME_CHARS with one warning that quotes it shortened', async () => {
    const name = 'n'.repeat(MAX_SHEET_NAME_CHARS + 1);
    const { info, pkg } = await read(await one(name));
    expect(info.sheets[0].name).toBe('n'.repeat(MAX_SHEET_NAME_CHARS));
    expect(pkg.warnings).toEqual([`The sheet name "${'n'.repeat(63)}…" is ${MAX_SHEET_NAME_CHARS + 1} characters long (Excel allows 31); it was cut to ${MAX_SHEET_NAME_CHARS}.`]);
  });

  it('cuts a 10 MB name to MAX_SHEET_NAME_CHARS, and the warnings stay small whatever the sheet has wrong', async () => {
    const name = 'N'.repeat(10_000_000);
    const { info, pkg } = await read(await one(name, '', [['rId2', 'worksheet', 'worksheets/gone.xml']]));
    expect(info.sheets[0].name).toHaveLength(MAX_SHEET_NAME_CHARS);
    expect(pkg.warnings).toHaveLength(2);
    expect(pkg.warnings.reduce((n, w) => n + w.length, 0)).toBeLessThan(1_000);
    expect(pkg.warnings[1]).toBe(`The sheet "${'N'.repeat(63)}…" has no usable relationship (rId1).`);
  });

  it('does not leave half a surrogate pair at the cut', async () => {
    const name = `${'a'.repeat(MAX_SHEET_NAME_CHARS - 1)}\u{1F600}`;
    const { info } = await read(await one(name));
    expect(info.sheets[0].name).toBe('a'.repeat(MAX_SHEET_NAME_CHARS - 1));
  });

  it('quotes a relationship target and a relationship id of any length shortened', async () => {
    const { pkg } = await read(await one('S', '', [['rId1', 'worksheet', `worksheets/${'x'.repeat(5_000)}.xml`]]));
    expect(pkg.warnings).toEqual([`The sheet "S" points at xl/worksheets/${'x'.repeat(49)}…, which the package does not contain.`]);
    const noRel = await read(await workbookZip(`<sheets><sheet name="S" sheetId="1" r:id="${'r'.repeat(5_000)}"/></sheets>`, []));
    expect(noRel.pkg.warnings).toEqual([`The sheet "S" has no usable relationship (${'r'.repeat(63)}…).`]);
  });

  it('quotes the sheet name and the target of a comment part the package does not contain shortened', async () => {
    const name = 'C'.repeat(200);
    const sheetRels = relsXml([['rId1', 'comments', `../${'c'.repeat(1_000)}.xml`]]);
    const bytes = await workbookZip(`<sheets><sheet name="${name}" sheetId="1" r:id="rId1"/></sheets>`, [['rId1', 'worksheet', 'worksheets/sheet1.xml']], {
      'xl/worksheets/sheet1.xml': '<worksheet/>',
      'xl/worksheets/_rels/sheet1.xml.rels': sheetRels,
    });
    const { pkg } = await read(bytes);
    expect(pkg.warnings).toHaveLength(1);
    expect(pkg.warnings[0]).toBe(`The sheet "${'C'.repeat(63)}…" has comments at xl/${'c'.repeat(60)}…, which the package does not contain.`);
  });
});
