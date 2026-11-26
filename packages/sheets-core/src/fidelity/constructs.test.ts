import { CONSTRUCTS, findConstruct } from '../../../../scripts/xlsx-fidelity/constructs';
import type { OoxmlParts } from '../../../../scripts/xlsx-fidelity/ooxml-package';

type Xml = Record<string, string>;
const mk = (xml: Xml = {}, extra: Partial<OoxmlParts> = {}): OoxmlParts => ({
  kind: 'zip',
  raw: new Uint8Array(),
  names: Object.keys(xml),
  xml,
  sizes: {},
  oversized: [],
  ...extra,
});
const sheet = (body: string): Xml => ({
  'xl/worksheets/sheet1.xml': `<worksheet>${body}</worksheet>`,
});
const cells = (c: string) => sheet(`<sheetData><row r="1">${c}</row></sheetData>`);
const wb = (body: string): Xml => ({ 'xl/workbook.xml': `<workbook>${body}</workbook>` });
const styles = (body: string): Xml => ({ 'xl/styles.xml': `<styleSheet>${body}</styleSheet>` });
const f = (formula: string, cell = 'A1', cellAttrs = '', attrs = ''): Xml =>
  cells(`<c r="${cell}"${cellAttrs}><f${attrs}>${formula}</f><v>1</v></c>`);
const names = (...n: string[]): OoxmlParts => mk({}, { names: n });
const ole = (name: string) =>
  mk({}, { kind: 'ole', raw: new Uint8Array(Buffer.from(`HEADER${name}`, 'utf16le')) });
const rels = (name: string, ...items: string[]): Xml => ({
  [name]: `<Relationships>${items.join('')}</Relationships>`,
});
const chartPart = (inner: string): Xml => ({
  'xl/charts/chart1.xml': `<c:chartSpace>${inner}</c:chartSpace>`,
});
const cf = (type: string) =>
  sheet(
    `<conditionalFormatting sqref="A1"><cfRule type="${type}" priority="1"/></conditionalFormatting>`
  );
const dv = (type: string) =>
  sheet(`<dataValidations count="1"><dataValidation type="${type}" sqref="A1"/></dataValidations>`);

/** One minimal package per construct, with the count the probe must return. */
const POSITIVE: Array<[string, OoxmlParts, number]> = [
  [
    'cells.numbers',
    mk(
      cells(
        '<c r="A1"><v>1</v></c><c r="B1" t="s"><v>0</v></c><c r="C1"><f>1</f><v>1</v></c><c r="D1" t="n"><v>2</v></c>'
      )
    ),
    1,
  ],
  ['cells.sharedStrings', mk(cells('<c r="A1" t="s"><v>0</v></c>')), 1],
  ['cells.inlineStrings', mk(cells('<c r="A1" t="inlineStr"><is><t>x</t></is></c>')), 1],
  ['cells.booleans', mk(cells('<c r="A1" t="b"><v>1</v></c>')), 1],
  ['cells.errors', mk(cells('<c r="A1" t="e"><v>#N/A</v></c>')), 1],
  [
    'cells.dates',
    mk({
      ...cells(
        '<c r="A1" s="1"><v>1</v></c><c r="B1" s="2"><v>1</v></c><c r="C1" s="3"><v>1</v></c><c r="D1" s="4"><v>1</v></c><c r="E1" s="5"><v>1</v></c><c r="F1" s="6"><v>1</v></c><c r="G1"><v>1</v></c>'
      ),
      ...styles(
        '<numFmts count="4"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/><numFmt numFmtId="165" formatCode="&quot;Total: &quot;0"/><numFmt numFmtId="166" formatCode="[Red]0.0;[Blue]\\d0"/><numFmt numFmtId="167" formatCode="&quot;Date: &quot;dd"/></numFmts><cellXfs count="7"><xf numFmtId="0"/><xf numFmtId="14"/><xf numFmtId="164"/><xf numFmtId="165"/><xf numFmtId="166"/><xf numFmtId="167"/><xf numFmtId="0"/></cellXfs>'
      ),
    }),
    3,
  ],
  ['cells.unicode', mk({ 'xl/sharedStrings.xml': '<sst><si><t>café</t></si></sst>' }), 1],
  ['cells.rtl', mk({ 'xl/sharedStrings.xml': '<sst><si><t>שלום</t></si></sst>' }), 4],
  [
    'cells.richText',
    mk({
      'xl/sharedStrings.xml': '<sst><si><r><t>a</t></r></si><si><t>b</t></si></sst>',
      ...cells('<c r="A1" t="inlineStr"><is><r><t>x</t></r></is></c>'),
    }),
    2,
  ],
  ['workbook.date1904', mk(wb('<workbookPr date1904="1"/>')), 1],
  ['workbook.calcSettings', mk(wb('<calcPr calcId="1"/>')), 1],
  ['styles.basic', mk(styles('<cellXfs count="3"><xf/><xf/><xf/></cellXfs>')), 2],
  ['styles.alignment', mk(styles('<cellXfs><xf><alignment wrapText="1"/></xf></cellXfs>')), 1],
  [
    'styles.customNumFmts',
    mk(styles('<numFmts><numFmt numFmtId="164" formatCode="0.0"/></numFmts>')),
    1,
  ],
  ['styles.indexedColors', mk(styles('<fonts><font><color indexed="10"/></font></fonts>')), 1],
  [
    'styles.borderStyles',
    mk(
      styles(
        '<borders><border><left style="thin"/><right style="thin"/><top style="thick"/><bottom/><diagonal style="dashed"/></border></borders>'
      )
    ),
    3,
  ],
  ['formulas.basic', mk(f('A2+1')), 1],
  [
    'formulas.cachedValues',
    mk(
      cells(
        '<c r="A1"><f>1+1</f><v>2</v></c><c r="B1"><f t="shared" si="0"/><v>2</v></c><c r="C1"><f>3</f></c>'
      )
    ),
    2,
  ],
  ['formulas.array', mk(f('SUM(B1:B2)', 'A1', '', ' t="array" ref="A1"')), 1],
  [
    'formulas.dynamicArray',
    mk(f('_xlfn.SEQUENCE(3)', 'A1', ' cm="1"', ' t="array" ref="A1:A3"')),
    1,
  ],
  ['formulas.xlfn', mk(f('_xlfn.IFS(1,2)')), 1],
  ['formulas.lambda', mk(f('_xlfn.LET(_xlpm.x,1,_xlpm.x)')), 1],
  ['formulas.wholeColumn', mk(f('SUM(B:B)+SUM($C:$D)+SUM(Data!E:E)')), 1],
  ['formulas.wholeRow', mk(f('SUM(2:2)+SUM($3:$4)')), 1],
  ['formulas.crossSheet', mk(f('Other!A1')), 1],
  ['formulas.threeD', mk(f("SUM(Sheet1:Sheet3!A1)+SUM('A b:C d'!A1)")), 1],
  [
    'names.defined',
    mk(wb('<definedNames><definedName name="A">B!$A$1</definedName></definedNames>')),
    1,
  ],
  [
    'names.sheetScoped',
    mk(
      wb(
        '<definedName name="A" localSheetId="0">B!$A$1</definedName><definedName name="B">B!$A$1</definedName>'
      )
    ),
    1,
  ],
  ['names.printArea', mk(wb('<definedName name="_xlnm.Print_Area">B!$A$1</definedName>')), 1],
  ['names.printTitles', mk(wb('<definedName name="_xlnm.Print_Titles">B!$1:$1</definedName>')), 1],
  [
    'names.constants',
    mk(
      wb(
        '<definedName name="N">0.2</definedName><definedName name="S">&quot;Hi&quot;</definedName><definedName name="Q">"x"</definedName><definedName name="L">{1,2}</definedName><definedName name="R">B!$A$1</definedName>'
      )
    ),
    4,
  ],
  [
    'layout.colWidths',
    mk(
      sheet(
        '<cols><col min="1" max="1" width="9" customWidth="1"/><col min="2" max="2" width="9"/></cols>'
      )
    ),
    1,
  ],
  [
    'layout.rowHeights',
    mk(sheet('<sheetData><row r="1" ht="30" customHeight="1"/></sheetData>')),
    1,
  ],
  ['layout.defaultSizes', mk(sheet('<sheetFormatPr defaultRowHeight="18"/>')), 1],
  [
    'layout.hidden',
    mk(
      sheet(
        '<cols><col min="1" max="1" hidden="1"/></cols><sheetData><row r="2" hidden="1"/></sheetData>'
      )
    ),
    2,
  ],
  [
    'layout.outline',
    mk(
      sheet(
        '<cols><col min="1" max="1" outlineLevel="1"/></cols><sheetData><row r="2" outlineLevel="2"/><row r="3" outlineLevel="0"/></sheetData>'
      )
    ),
    2,
  ],
  [
    'layout.columnStyles',
    mk(sheet('<cols><col min="1" max="1" style="3"/><col min="2" max="2" style="0"/></cols>')),
    1,
  ],
  [
    'layout.frozenPanes',
    mk(sheet('<sheetViews><sheetView><pane ySplit="1" state="frozen"/></sheetView></sheetViews>')),
    1,
  ],
  ['layout.merges', mk(sheet('<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>')), 1],
  [
    'layout.zoom',
    mk(sheet('<sheetViews><sheetView zoomScale="130"/><sheetView zoomScale="100"/></sheetViews>')),
    1,
  ],
  ['layout.noGridlines', mk(sheet('<sheetViews><sheetView showGridLines="0"/></sheetViews>')), 1],
  ['layout.rtl', mk(sheet('<sheetViews><sheetView rightToLeft="1"/></sheetViews>')), 1],
  ['sheets.multiple', mk(wb('<sheets><sheet name="A"/><sheet name="B"/></sheets>')), 2],
  ['sheets.hidden', mk(wb('<sheets><sheet name="A" state="hidden"/></sheets>')), 1],
  ['sheets.veryHidden', mk(wb('<sheets><sheet name="A" state="veryHidden"/></sheets>')), 1],
  ['sheets.tabColor', mk(sheet('<sheetPr><tabColor rgb="FF0000"/></sheetPr>')), 1],
  ['sheets.activeTab', mk(wb('<bookViews><workbookView activeTab="2"/></bookViews>')), 1],
  ['cf.cellIs', mk(cf('cellIs')), 1],
  ['cf.expression', mk(cf('expression')), 1],
  ['cf.top10', mk(cf('top10')), 1],
  ['cf.aboveAverage', mk(cf('aboveAverage')), 1],
  ['cf.duplicates', mk(cf('uniqueValues')), 1],
  ['cf.text', mk(cf('beginsWith')), 1],
  ['cf.dates', mk(cf('timePeriod')), 1],
  ['cf.colorScale', mk(cf('colorScale')), 1],
  ['cf.dataBar', mk(cf('dataBar')), 1],
  ['cf.iconSet', mk(cf('iconSet')), 1],
  ['cf.stopIfTrue', mk(sheet('<cfRule type="expression" stopIfTrue="1"/>')), 1],
  ['dv.list', mk(dv('list')), 1],
  ['dv.whole', mk(dv('whole')), 1],
  ['dv.decimal', mk(dv('decimal')), 1],
  ['dv.date', mk(dv('date')), 1],
  ['dv.time', mk(dv('time')), 1],
  ['dv.textLength', mk(dv('textLength')), 1],
  ['dv.custom', mk(dv('custom')), 1],
  [
    'dv.messages',
    mk(sheet('<dataValidation type="list" promptTitle="t"/><dataValidation type="list"/>')),
    1,
  ],
  [
    'comments.notes',
    mk({
      'xl/comments1.xml': '<comments><comment ref="A1"/></comments>',
      'xl/comments/comment2.xml': '<comments><comment ref="A1"/></comments>',
    }),
    2,
  ],
  ['images.png', names('xl/media/image1.png', 'xl/media/image2.PNG'), 2],
  ['charts.bar', mk(chartPart('<c:barChart/><barChart/>')), 2],
  ['charts.line', mk(chartPart('<c:lineChart/>')), 1],
  ['tables.listObjects', names('xl/tables/table1.xml', 'xl/tables/table2.xml'), 2],
  ['print.pageSetup', mk(sheet('<pageSetup orientation="landscape"/><pageSetup r:id="rId1"/>')), 1],
  [
    'edge.truncatedZip',
    mk({}, { kind: 'unreadable', raw: new Uint8Array([0x50, 0x4b, 0x03, 0x04, 1, 2]) }),
    1,
  ],
  ['edge.encrypted', ole('EncryptionInfo'), 1],
  ['edge.legacyBinary', ole('Workbook'), 1],
  ['edge.noWorkbook', mk({ '[Content_Types].xml': '<Types/>' }), 1],
  [
    'edge.brokenRels',
    mk({
      ...rels(
        'xl/_rels/workbook.xml.rels',
        '<Relationship Id="rId1" Target="worksheets/missing.xml"/>',
        '<Relationship Id="rId2" Target="https://e.com" TargetMode="External"/>'
      ),
      ...wb('<sheets><sheet name="A" r:id="rId9"/><sheet name="B" r:id="rId2"/></sheets>'),
      'readme.txt': 'x',
    }),
    2,
  ],
  ['edge.invalidXml', mk({ 'xl/a.xml': '<a><b></a>', 'xl/b.xml': '<a/>' }), 1],
  ['edge.xmlEntities', mk({ 'xl/a.xml': '<!DOCTYPE a [<!ENTITY e "v"><!ENTITY f "w">]><a/>' }), 2],
  [
    'edge.duplicateSheetNames',
    mk(wb('<sheets><sheet name="Data"/><sheet name="data"/><sheet name="Other"/></sheets>')),
    1,
  ],
  [
    'edge.indexOutOfRange',
    mk({
      ...cells(
        '<c r="A1" t="s"><v>5</v></c><c r="B1" t="s"><v>0</v></c><c r="C1" s="3"><f>1</f><v>1</v></c><c r="D1" t="s"/><c r="E1" t="s"><f>1</f></c>'
      ),
      'xl/sharedStrings.xml': '<sst><si/></sst>',
      ...styles('<cellXfs><xf/></cellXfs>'),
    }),
    2,
  ],
  ['edge.emptyWorkbook', mk(sheet('<sheetData/>')), 1],
  ['edge.hugeDimension', mk(sheet('<dimension ref="A1:XFD1048576"/>')), 1],
  [
    'edge.zipBomb',
    mk(
      {},
      {
        sizes: {
          'a.xml': { compressed: 1000, uncompressed: 300_000_000 },
          'b.xml': { compressed: 0, uncompressed: 2_000_000 },
          'c.xml': { compressed: 10, uncompressed: 500 },
          'd.xml': { compressed: 100_000, uncompressed: 1_500_000 },
        },
      }
    ),
    2,
  ],
  ['perf.largeSheet', mk(sheet('<dimension ref="A1:J100000"/>')), 1],
];

describe('XLSX construct probes', () => {
  it('has one positive sample per registry entry and unique ids and labels', () => {
    expect(POSITIVE.map(([id]) => id).sort()).toEqual(CONSTRUCTS.map((c) => c.id).sort());
    expect(new Set(CONSTRUCTS.map((c) => c.id)).size).toBe(CONSTRUCTS.length);
    expect(new Set(CONSTRUCTS.map((c) => c.label)).size).toBe(CONSTRUCTS.length);
    for (const c of CONSTRUCTS) expect(c.id).toMatch(/^[a-z]+(?:\.[A-Za-z0-9]+)*$/);
  });

  it.each(POSITIVE)('%s counts its construct', (id, parts, expected) => {
    expect(findConstruct(id)?.count(parts)).toBe(expected);
  });

  it.each(CONSTRUCTS.map((c) => [c.id, c] as const))(
    '%s counts nothing in an empty package',
    (_id, c) => {
      expect(c.count(mk())).toBe(0);
      expect(c.count(mk({}, { kind: 'unreadable' }))).toBe(0);
      expect(c.count(mk({}, { kind: 'ole' }))).toBe(0);
    }
  );

  it('does not mistake look-alikes for constructs', () => {
    const none = (id: string, p: OoxmlParts) =>
      expect({ id, n: findConstruct(id)?.count(p) }).toEqual({ id, n: 0 });
    none('formulas.wholeColumn', mk(f('SUM(A1:B2)+SUM(Data!A1:B2)+LOG10(2)+A1:A9')));
    none('formulas.wholeRow', mk(f('ROW(A1)+LOG10(2)')));
    none('formulas.threeD', mk(f('SUM(A1:B2)+Other!A1')));
    none('formulas.dynamicArray', mk(f('1', 'A1', ' cm="1"')));
    none('formulas.dynamicArray', mk(f('1', 'A1', '', ' t="array"')));
    none('names.constants', mk(wb('<definedName name="R">Data!$A$1</definedName>')));
    none('cells.numbers', mk(cells('<c r="A1" t="str"><v>x</v></c>')));
    none('sheets.multiple', mk(wb('<sheets><sheet name="A"/></sheets>')));
    none(
      'edge.noWorkbook',
      mk({ 'xl/workbook.xml': '<workbook/>', '[Content_Types].xml': '<Types/>' })
    );
    none('edge.noWorkbook', mk({ 'xl/styles.xml': '<styleSheet/>' }));
    none('edge.noWorkbook', mk({ '[Content_Types].xml': '<Types/>' }, { kind: 'ole' }));
    none(
      'edge.truncatedZip',
      mk({}, { kind: 'unreadable', raw: new Uint8Array([0x50, 0x4b, 0x01, 0x02]) })
    );
    none(
      'edge.truncatedZip',
      mk({}, { kind: 'zip', raw: new Uint8Array([0x50, 0x4b, 0x03, 0x04]) })
    );
    none(
      'edge.encrypted',
      mk({}, { kind: 'zip', raw: new Uint8Array(Buffer.from('EncryptionInfo', 'utf16le')) })
    );
    none('edge.legacyBinary', ole('Other'));
    none(
      'edge.emptyWorkbook',
      mk({
        ...sheet('<sheetData/>'),
        'xl/worksheets/sheet2.xml': '<worksheet><sheetData><row/></sheetData></worksheet>',
      })
    );
    none('edge.emptyWorkbook', mk(sheet('<dimension ref="A1"/>')));
    none('edge.hugeDimension', mk(sheet('<dimension ref="A1:Z100"/>')));
    none('edge.hugeDimension', mk(sheet('<sheetData/>')));
    none('perf.largeSheet', mk(sheet('<dimension ref="A1:B99999"/>')));
    none(
      'edge.indexOutOfRange',
      mk({
        ...cells('<c r="A1" t="s"><v>0</v></c><c r="B1" s="0"/>'),
        'xl/sharedStrings.xml': '<sst><si/></sst>',
        ...styles('<cellXfs><xf/></cellXfs>'),
      })
    );
    none(
      'edge.zipBomb',
      mk({}, { sizes: { 'a.xml': { compressed: 5_001, uncompressed: 1_000_000 } } })
    );
  });

  it('reads single-cell dimensions and cell dimensions without a range', () => {
    const p = mk(sheet('<dimension ref="XFD1048576"/>'));
    expect(findConstruct('edge.hugeDimension')?.count(p)).toBe(0);
    expect(
      findConstruct('perf.largeSheet')?.count(mk(sheet('<dimension ref="A1:A100000"/>')))
    ).toBe(1);
    expect(
      findConstruct('edge.hugeDimension')?.count(
        mk({
          ...sheet('<dimension ref="A1:B2"/>'),
          'xl/worksheets/sheet2.xml': '<worksheet><dimension ref="A1:XFD1048576"/></worksheet>',
        })
      )
    ).toBe(1);
  });

  it('counts self-closing formulas (shared-formula followers)', () => {
    expect(
      findConstruct('formulas.basic')?.count(mk(cells('<c r="A2"><f t="shared" si="0"/></c>')))
    ).toBe(1);
  });

  it('decodes escaped formula text before matching', () => {
    expect(
      findConstruct('formulas.wholeColumn')?.count(mk(f('SUM(A:A)&amp;&quot;x&quot;&lt;&gt;1')))
    ).toBe(1);
  });

  it('finds constructs by id', () => {
    expect(findConstruct('cells.numbers')?.label).toBe('Numeric cells');
    expect(findConstruct('nope')).toBeUndefined();
  });
});
