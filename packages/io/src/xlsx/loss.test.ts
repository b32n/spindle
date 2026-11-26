import JSZip from 'jszip';
import { detectXlsxLoss, LIMITS, type DetectXlsxLossOptions } from './loss';
import type { ImportLossItem, ImportLossReport } from '@b32nio/spindle-shared';

// ---------------------------------------------------------------------------
// Minimal synthetic workbooks, one construct at a time. Each package is a real zip built here, so the
// detector runs through the same code path as for a file from disk.
// ---------------------------------------------------------------------------

const MAIN = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const X14 = 'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"';
const XDR = 'xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"';

interface SheetSpec {
  name?: string;
  state?: string;
  /** Rows (`<row>...</row>`) placed inside sheetData. */
  rows?: string;
  /** Elements after sheetData (worksheet children), and before it (`head`: sheetPr, sheetViews, cols). */
  tail?: string;
  head?: string;
  /** Relationship elements of this sheet. */
  rels?: string[];
  /** The whole sheet part, replacing the generated one. */
  raw?: string;
  /** `worksheet` (default) or `chartsheet`: the workbook relationship type. */
  kind?: string;
}

interface Spec {
  sheets?: SheetSpec[];
  /** Children of `<workbook>` after `<sheets>`. */
  workbook?: string;
  /** Attributes-bearing elements before `<sheets>` (workbookPr, bookViews). */
  workbookHead?: string;
  styles?: string;
  sharedStrings?: string;
  files?: Record<string, string>;
}

const STYLES = (inner = '') => `<styleSheet ${MAIN}>${inner}</styleSheet>`;

async function build(spec: Spec = {}): Promise<Uint8Array> {
  const sheets = spec.sheets ?? [{ rows: '<row r="1"><c r="A1"><v>1</v></c></row>' }];
  const zip = new JSZip();
  const wbRels: string[] = [];
  const sheetEls: string[] = [];
  sheets.forEach((s, i) => {
    const n = i + 1;
    const kind = s.kind ?? 'worksheet';
    const part = kind === 'chartsheet' ? `xl/chartsheets/sheet${n}.xml` : `xl/worksheets/sheet${n}.xml`;
    wbRels.push(`<Relationship Id="rId${n}" Type="${REL}/${kind}" Target="${part.slice(3)}"/>`);
    sheetEls.push(`<sheet name="${s.name ?? `Sheet${n}`}" sheetId="${n}"${s.state ? ` state="${s.state}"` : ''} r:id="rId${n}"/>`);
    zip.file(
      part,
      s.raw ?? `<worksheet ${MAIN}>${s.head ?? ''}<sheetData>${s.rows ?? ''}</sheetData>${s.tail ?? ''}</worksheet>`
    );
    if (s.rels) {
      const dir = kind === 'chartsheet' ? 'chartsheets' : 'worksheets';
      zip.file(`xl/${dir}/_rels/sheet${n}.xml.rels`, `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${s.rels.join('')}</Relationships>`);
    }
  });
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>');
  zip.file('xl/_rels/workbook.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${wbRels.join('')}</Relationships>`);
  zip.file('xl/workbook.xml', `<workbook ${MAIN}>${spec.workbookHead ?? ''}<sheets>${sheetEls.join('')}</sheets>${spec.workbook ?? ''}</workbook>`);
  if (spec.styles !== undefined) zip.file('xl/styles.xml', spec.styles);
  if (spec.sharedStrings !== undefined) zip.file('xl/sharedStrings.xml', spec.sharedStrings);
  for (const [name, content] of Object.entries(spec.files ?? {})) zip.file(name, content);
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

const full = (bytes: Uint8Array, options: DetectXlsxLossOptions = {}) => detectXlsxLoss(bytes, { detail: 'full', ...options });
const ids = (r: ImportLossReport) => r.items.map((i) => i.id).sort();
const item = (r: ImportLossReport, id: string): ImportLossItem | undefined => r.items.find((i) => i.id === id);
const cell = (ref: string, inner: string, attrs = '') => `<c r="${ref}"${attrs}>${inner}</c>`;
const f = (text: string, attrs = '', cellAttrs = '') => `<row r="1">${cell('A1', `<f${attrs}>${text.replace(/&/g, '&amp;').replace(/</g, '&lt;')}</f><v>1</v>`, cellAttrs)}</row>`;
const rel = (id: string, type: string, target: string, external = false) =>
  `<Relationship Id="${id}" Type="${REL}/${type}" Target="${target}"${external ? ' TargetMode="External"' : ''}/>`;

const KNOWN = new Set(['SUM', 'IF']);

describe('plain content', () => {
  it('reports nothing for a plain workbook, in either detail', async () => {
    const bytes = await build({ styles: STYLES('<cellXfs><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>') });
    expect(await detectXlsxLoss(bytes)).toEqual({ format: 'xlsx', items: [] });
    expect(await full(bytes)).toEqual({ format: 'xlsx', items: [] });
  });

  it('defaults to the summary detail', async () => {
    const bytes = await build({ sheets: [{ head: '<sheetPr><tabColor rgb="FFFF0000"/></sheetPr>', tail: '<autoFilter ref="A1:A2"/>' }] });
    expect(ids(await detectXlsxLoss(bytes))).toEqual(['xlsx.autoFilter']);
    expect(ids(await full(bytes))).toEqual(['xlsx.autoFilter', 'xlsx.sheet.tabColor']);
  });
});

// ---------------------------------------------------------------------------
// The summary ids (what the dialog shows): semantics of the old detector, counts fixed
// ---------------------------------------------------------------------------

describe('summary ids', () => {
  it('counts conditional formatting, validation, filters, protection, outline and sparklines per sheet', async () => {
    const bytes = await build({
      sheets: [
        {
          name: 'Data',
          head: '<cols><col min="1" max="2" width="9" outlineLevel="1"/><col min="3" max="3" outlineLevel="0"/></cols>',
          rows: '<row r="1" outlineLevel="2"><c r="A1"/></row><row r="2" outlineLevel="0"/>',
          tail:
            '<sheetProtection sheet="1"/><autoFilter ref="A1:B2"/>' +
            '<conditionalFormatting sqref="A1"><cfRule type="cellIs" priority="1"/></conditionalFormatting>' +
            '<conditionalFormatting sqref="B1"><cfRule type="cellIs" priority="2"/></conditionalFormatting>' +
            '<dataValidations count="1"><dataValidation type="list" sqref="C1"><formula1>"x,y"</formula1></dataValidation></dataValidations>' +
            `<extLst><ext uri="u" ${X14}><x14:sparklineGroups><x14:sparklineGroup><x14:sparklines/></x14:sparklineGroup><x14:sparklineGroup/></x14:sparklineGroups></ext></extLst>`,
        },
        { name: 'Other', tail: '<sheetProtection sheet="0"/>' },
      ],
    });
    const report = await detectXlsxLoss(bytes);
    expect(item(report, 'xlsx.conditionalFormatting')).toMatchObject({ count: 2, where: ['Data'] });
    expect(item(report, 'xlsx.dataValidation')).toMatchObject({ count: 1, where: ['Data'] });
    expect(item(report, 'xlsx.autoFilter')).toMatchObject({ count: 1 });
    expect(item(report, 'xlsx.sheetProtection')).toMatchObject({ count: 1, where: ['Data'] });
    // one column group (the col element) plus one row group; the level-0 ones are not groups
    expect(item(report, 'xlsx.outline')).toMatchObject({ count: 2, where: ['Data'] });
    expect(item(report, 'xlsx.sparklines')).toMatchObject({ count: 2 });
  });

  it('counts hidden and very hidden sheets, the names (not the built-in ones) and workbook protection', async () => {
    const bytes = await build({
      sheets: [{ name: 'A' }, { name: 'B', state: 'hidden' }, { name: 'C', state: 'veryHidden' }, { name: 'D', state: 'visible' }],
      workbook:
        '<workbookProtection lockStructure="true"/><definedNames><definedName name="Rate">A!$A$1</definedName><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">A!$A$1</definedName><definedName name="_xlnm.Print_Area" localSheetId="0">A!$A$1</definedName></definedNames>',
    });
    const report = await detectXlsxLoss(bytes);
    expect(item(report, 'xlsx.hiddenSheets')).toMatchObject({ count: 2, where: ['B', 'C'] });
    expect(item(report, 'xlsx.definedNames')).toMatchObject({ count: 1, where: ['Rate'] });
    expect(item(report, 'xlsx.workbookProtection')).toMatchObject({ count: 1 });
  });

  it('treats lockWindows as workbook protection and an empty element as none', async () => {
    expect(ids(await detectXlsxLoss(await build({ workbook: '<workbookProtection lockWindows="1"/>' })))).toEqual(['xlsx.workbookProtection']);
    expect(ids(await detectXlsxLoss(await build({ workbook: '<workbookProtection/>' })))).toEqual([]);
  });

  it('tolerates sheets and names with missing attributes, and a foreign-namespace conditional format', async () => {
    const zip = await JSZip.loadAsync(await build());
    zip.file(
      'xl/workbook.xml',
      `<workbook ${MAIN}><sheets><sheet/><sheet name="B" sheetId="2"/></sheets><definedNames><definedName>x</definedName><definedName name="K">y</definedName></definedNames></workbook>`
    );
    const report = await detectXlsxLoss(await zip.generateAsync({ type: 'uint8array' }));
    expect(item(report, 'xlsx.definedNames')).toMatchObject({ count: 1, where: ['K'] });
    const foreign = await build({ sheets: [{ tail: `<x14:conditionalFormatting ${X14}><x14:cfRule type="expression"/></x14:conditionalFormatting>` }] });
    expect(ids(await detectXlsxLoss(foreign))).toEqual([]);
  });

  it('counts a hidden sheet whose relationship does not resolve', async () => {
    const bytes = await build({ sheets: [{ name: 'A' }, { name: 'Lost', state: 'hidden' }] });
    const zip = await JSZip.loadAsync(bytes);
    zip.file('xl/_rels/workbook.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>');
    const report = await detectXlsxLoss(await zip.generateAsync({ type: 'uint8array' }));
    expect(item(report, 'xlsx.hiddenSheets')).toMatchObject({ count: 1, where: ['Lost'] });
  });

  it('counts a chart once, however many parts it has, and a chart sheet through its drawing', async () => {
    const chartFrame = (id: number) =>
      `<xdr:twoCellAnchor><xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="Chart"/></xdr:nvGraphicFramePr><a:graphic><a:graphicData uri="x"><c:chart r:id="rId1"/></a:graphicData></a:graphic></xdr:graphicFrame></xdr:twoCellAnchor>`;
    const bytes = await build({
      sheets: [
        { name: 'S', rels: [rel('rId1', 'drawing', '../drawings/drawing1.xml')] },
        { name: 'Chart', kind: 'chartsheet', raw: `<chartsheet ${MAIN}/>`, rels: [rel('rId1', 'drawing', '../drawings/drawing2.xml')] },
      ],
      files: {
        'xl/drawings/drawing1.xml': `<xdr:wsDr ${XDR}>${chartFrame(2)}${chartFrame(3)}</xdr:wsDr>`,
        'xl/drawings/drawing2.xml': `<xdr:wsDr ${XDR}>${chartFrame(2)}</xdr:wsDr>`,
        'xl/charts/chart1.xml': '<c:chartSpace/>',
        'xl/charts/chart2.xml': '<c:chartSpace/>',
        'xl/charts/chart3.xml': '<c:chartSpace/>',
        'xl/charts/style1.xml': '<cs:chartStyle/>',
        'xl/charts/colors1.xml': '<cs:colorStyle/>',
      },
    });
    expect(item(await detectXlsxLoss(bytes), 'xlsx.charts')).toMatchObject({ count: 3, where: ['S', 'Chart'] });
  });

  it('falls back to the chart parts when no drawing leads to them', async () => {
    const bytes = await build({ files: { 'xl/charts/chart1.xml': '<c:chartSpace/>', 'xl/charts/chart2.xml': '<c:chartSpace/>', 'xl/charts/style1.xml': '<x/>' } });
    expect(item(await detectXlsxLoss(bytes), 'xlsx.charts')).toMatchObject({ count: 2 });
  });

  it('counts pictures (not media files), sheet backgrounds and cell pictures; falls back to media files', async () => {
    const pic = (id: number) => `<xdr:oneCellAnchor><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${id}" name="P"/></xdr:nvPicPr></xdr:pic></xdr:oneCellAnchor>`;
    const bytes = await build({
      sheets: [{ name: 'S', tail: '<picture r:id="rId9"/>', rels: [rel('rId1', 'drawing', '../drawings/drawing1.xml')] }],
      files: {
        'xl/drawings/drawing1.xml': `<xdr:wsDr ${XDR}>${pic(2)}${pic(3)}</xdr:wsDr>`,
        'xl/media/image1.png': 'x',
        'xl/media/image2.png': 'x',
        'xl/media/image3.png': 'x',
        'xl/cellimages.xml': '<etc:cellImages xmlns:etc="e"><etc:cellImage/></etc:cellImages>',
      },
    });
    // two pictures + one background + one cell picture; three media files are not the count
    expect(item(await detectXlsxLoss(bytes), 'xlsx.images')).toMatchObject({ count: 4, where: ['S'] });

    const orphan = await build({ files: { 'xl/media/image1.png': 'x', 'xl/media/image2.png': 'x' } });
    expect(item(await detectXlsxLoss(orphan), 'xlsx.images')).toMatchObject({ count: 2 });
  });

  it('counts the sheets with comments once (notes plus threaded comments), and falls back to the parts', async () => {
    const bytes = await build({
      sheets: [
        { name: 'A', rels: [rel('rId1', 'comments', '../comments1.xml'), rel('rId2', 'threadedComment', '../threadedComments/threadedComment1.xml')] },
        { name: 'B', rels: [rel('rId1', 'threadedComment', '../threadedComments/threadedComment2.xml')] },
        { name: 'C' },
      ],
      files: {
        'xl/comments1.xml': '<comments/>',
        'xl/threadedComments/threadedComment1.xml': '<ThreadedComments/>',
        'xl/threadedComments/threadedComment2.xml': '<ThreadedComments/>',
      },
    });
    expect(item(await detectXlsxLoss(bytes), 'xlsx.comments')).toMatchObject({ count: 2, where: ['A', 'B'] });

    const orphan = await build({ files: { 'xl/comments1.xml': '<comments/>', 'xl/comments2.xml': '<comments/>' } });
    expect(item(await detectXlsxLoss(orphan), 'xlsx.comments')).toMatchObject({ count: 2 });
    const orphanThreaded = await build({ files: { 'xl/threadedComments/t1.xml': '<x/>' } });
    expect(item(await detectXlsxLoss(orphanThreaded), 'xlsx.comments')).toMatchObject({ count: 1 });
  });

  it('counts tables, pivot tables, macros, external links and slicers from their parts', async () => {
    const bytes = await build({
      files: {
        'xl/tables/table1.xml': '<table/>',
        'xl/tables/table2.xml': '<table/>',
        'xl/tables/_rels/table1.xml.rels': '<Relationships/>',
        'xl/pivotTables/pivotTable1.xml': '<pivotTableDefinition/>',
        'xl/pivotTables/_rels/pivotTable1.xml.rels': '<Relationships/>',
        'xl/vbaProject.bin': 'x',
        'xl/externalLinks/externalLink1.xml': '<externalLink/>',
        'xl/externalLinks/_rels/externalLink1.xml.rels': '<Relationships/>',
        // one slicer part holding two slicers, one that cannot be read, and a rels part that is not a slicer
        'xl/slicers/slicer1.xml': '<slicers xmlns="x"><slicer name="a"/><slicer name="b"/></slicers>',
        'xl/slicers/slicer2.xml': 'not xml',
        'xl/slicers/_rels/slicer1.xml.rels': '<Relationships/>',
        'xl/slicerCaches/slicerCache1.xml': '<slicerCacheDefinition/>',
      },
    });
    const report = await detectXlsxLoss(bytes);
    expect(item(report, 'xlsx.tables')?.count).toBe(2);
    expect(item(report, 'xlsx.pivotTables')?.count).toBe(1);
    expect(item(report, 'xlsx.macros')?.count).toBe(1);
    expect(item(report, 'xlsx.externalLinks')?.count).toBe(1);
    expect(item(report, 'xlsx.slicers')?.count).toBe(3);
  });

  it('counts rich text in shared and inline strings, including a first run without formatting', async () => {
    const bytes = await build({
      sharedStrings:
        `<sst ${MAIN}><si><t>plain</t></si><si><r><t>a</t></r><r><rPr><b/></rPr><t>b</t></r></si><si><r><t>a</t></r><r><t>b</t></r></si><si><r><rPr><i/></rPr><t>c</t></r></si></sst>`,
      sheets: [{ rows: '<row r="1"><c r="A1" t="inlineStr"><is><r><rPr><b/></rPr><t>x</t></r></is></c><c r="B1" t="inlineStr"><is><t>plain</t></is></c></row>' }],
    });
    expect(item(await detectXlsxLoss(bytes), 'xlsx.richText')).toMatchObject({ count: 3 });
  });
});

describe('effect text convention', () => {
  it('starts every effect with its state word', async () => {
    const bytes = await kitchenSink();
    const report = await full(bytes, { knownFunctions: KNOWN, isNumberFormatExact: () => false });
    expect(report.items.length).toBeGreaterThan(50);
    for (const i of report.items) {
      expect(i.effect).toMatch(/^(lost|approximated|kept): \S/);
      expect(i.id).toMatch(/^xlsx\.[a-zA-Z]+(\.[a-zA-Z0-9]+)?$/);
      expect(i.count).toBeGreaterThan(0);
      expect(i.where === undefined || (i.where.length > 0 && i.where.length <= 5)).toBe(true);
    }
  });
});

// ---------------------------------------------------------------------------
// New summary ids and option handling
// ---------------------------------------------------------------------------

const SUMMARY_IDS = [
  'xlsx.autoFilter',
  'xlsx.charts',
  'xlsx.comments',
  'xlsx.conditionalFormatting',
  'xlsx.dataValidation',
  'xlsx.definedNames',
  'xlsx.externalLinks',
  'xlsx.formulas.structuredRefs',
  'xlsx.formulas.unsupported',
  'xlsx.hiddenSheets',
  'xlsx.images',
  'xlsx.macros',
  'xlsx.numberFormats.approximated',
  'xlsx.outline',
  'xlsx.pivotTables',
  'xlsx.richText',
  'xlsx.sheetProtection',
  'xlsx.slicers',
  'xlsx.sparklines',
  'xlsx.tables',
  'xlsx.workbookProtection',
];

async function kitchenSink(): Promise<Uint8Array> {
  const pic = '<xdr:oneCellAnchor><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="P"/></xdr:nvPicPr></xdr:pic></xdr:oneCellAnchor>';
  const frame = '<xdr:twoCellAnchor><xdr:graphicFrame><a:graphic><a:graphicData uri="x"><c:chart/></a:graphicData></a:graphic></xdr:graphicFrame></xdr:twoCellAnchor>';
  const sp = '<xdr:twoCellAnchor><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="5" name="S"/><xdr:cNvSpPr/></xdr:nvSpPr></xdr:sp></xdr:twoCellAnchor>';
  const box = '<xdr:twoCellAnchor><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="6" name="T"/><xdr:cNvSpPr txBox="1"/></xdr:nvSpPr></xdr:sp></xdr:twoCellAnchor>';
  return build({
    workbookHead: '<workbookPr date1904="1"/>',
    sheets: [
      {
        name: 'Data',
        head:
          '<sheetPr><tabColor rgb="FFFF0000"/><pageSetUpPr fitToPage="1"/></sheetPr><sheetViews><sheetView zoomScale="150" rightToLeft="1" showGridLines="0"/></sheetViews>' +
          '<cols><col min="1" max="1" style="1" outlineLevel="1"/></cols>',
        rows:
          '<row r="1" outlineLevel="1" s="1" customFormat="1">' +
          '<c r="A1" s="1"><f>SUM(Table1[Col])+_xlfn.FANCY(1)+NOW()</f><v>1</v></c>' +
          '<c r="B1"><f t="array" ref="B1:B2">SUM(A:A)</f><v>1</v></c>' +
          '<c r="C1" cm="1"><f t="array" ref="C1:C2">SORT(A1)</f><v>1</v></c>' +
          "<c r=\"D1\"><f>'[1]S'!A1</f><v>1</v></c><c r=\"E1\" s=\"2\"/><c r=\"F1\" s=\"3\"/><c r=\"G1\" s=\"4\"/></row>",
        tail:
          '<sheetProtection sheet="1"/><autoFilter ref="A1:B2"/>' +
          '<conditionalFormatting sqref="A1"><cfRule type="expression" priority="1" stopIfTrue="1"/></conditionalFormatting>' +
          '<dataValidations><dataValidation type="custom" sqref="A1" error="no"><formula1>Other!A1</formula1></dataValidation></dataValidations>' +
          '<hyperlinks><hyperlink ref="A1" location="Other!A1"/><hyperlink ref="A2" r:id="rId3"/></hyperlinks>' +
          '<pageSetup orientation="landscape"/><headerFooter><oddHeader>&amp;CHi</oddHeader></headerFooter><rowBreaks count="1"><brk id="3"/></rowBreaks>' +
          '<customSheetViews><customSheetView guid="g"/></customSheetViews>' +
          `<extLst><ext uri="u" ${X14}><x14:conditionalFormattings><x14:conditionalFormatting><x14:cfRule type="dataBar" id="i"/><x14:cfRule type="iconSet" id="j"/></x14:conditionalFormatting></x14:conditionalFormattings><x14:sparklineGroups><x14:sparklineGroup/></x14:sparklineGroups></ext></extLst>`,
        rels: [
          rel('rId1', 'drawing', '../drawings/drawing1.xml'),
          rel('rId2', 'comments', '../comments1.xml'),
          rel('rId3', 'hyperlink', 'javascript:alert(1)', true),
          rel('rId4', 'threadedComment', '../threadedComments/threadedComment1.xml'),
        ],
      },
      { name: 'Hidden', state: 'veryHidden' },
    ],
    workbook: '<workbookProtection lockStructure="1"/><definedNames><definedName name="N">Data!$A$1</definedName><definedName name="_xlnm.Print_Area" localSheetId="0">Data!$A$1</definedName><definedName name="_xlnm.Print_Titles" localSheetId="0">Data!$1:$1</definedName></definedNames><calcPr calcMode="manual"/>',
    styles: STYLES(
      '<numFmts><numFmt numFmtId="164" formatCode="[$€-407] #,##0.00"/></numFmts>' +
        '<fonts><font/><font><color theme="4"/></font></fonts>' +
        '<fills><fill/><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="lightGray"/></fill><fill><gradientFill><stop position="0"><color rgb="FF000000"/></stop></gradientFill></fill></fills>' +
        '<borders><border/><border><diagonal style="thin"/></border></borders>' +
        '<cellXfs><xf/><xf numFmtId="164" fontId="1" fillId="0" borderId="0" xfId="1"/><xf fillId="2"/><xf fillId="3"/><xf borderId="1"/></cellXfs>' +
        '<cellStyles><cellStyle name="Normal" xfId="0"/><cellStyle name="Good" xfId="1"/></cellStyles>'
    ),
    sharedStrings: `<sst ${MAIN}><si><r><rPr><b/></rPr><t>a</t></r></si></sst>`,
    files: {
      'xl/drawings/drawing1.xml': `<xdr:wsDr ${XDR}>${pic}${frame}${sp}${box}</xdr:wsDr>`,
      'xl/comments1.xml': '<comments><commentList><comment ref="A1"><text><r><rPr><b/></rPr><t>x</t></r></text></comment></commentList></comments>',
      'xl/threadedComments/threadedComment1.xml': '<ThreadedComments><threadedComment ref="A1"/></ThreadedComments>',
      'xl/tables/table1.xml': '<table/>',
      'xl/pivotTables/pivotTable1.xml': '<pivotTableDefinition/>',
      'xl/vbaProject.bin': 'x',
      'xl/externalLinks/externalLink1.xml': '<externalLink/>',
      'xl/slicers/slicer1.xml': '<slicers><slicer/></slicers>',
      'xl/embeddings/oleObject1.bin': 'x',
      'xl/ctrlProps/ctrlProp1.xml': '<formControlPr/>',
      'xl/connections.xml': '<connections><connection id="1"/></connections>',
      'customXml/item1.xml': '<x/>',
    },
  });
}

describe('summary versus full', () => {
  it('pins the summary id set: the old ids plus the three new ones, and nothing from the full level', async () => {
    const bytes = await kitchenSink();
    const summary = await detectXlsxLoss(bytes, { knownFunctions: KNOWN, isNumberFormatExact: () => false });
    expect(ids(summary)).toEqual(SUMMARY_IDS);
    const everything = await full(bytes, { knownFunctions: KNOWN, isNumberFormatExact: () => false });
    for (const id of SUMMARY_IDS) expect(ids(everything)).toContain(id);
    expect(ids(everything).length).toBeGreaterThan(SUMMARY_IDS.length + 30);
  });

  it('keeps the full-only ids out of the summary and reports identical summary items in both', async () => {
    const bytes = await kitchenSink();
    const options = { knownFunctions: KNOWN, isNumberFormatExact: () => false };
    const summary = await detectXlsxLoss(bytes, options);
    const everything = await full(bytes, options);
    for (const i of summary.items) expect(everything.items).toContainEqual(i);
  });

  it('does not report unsupported functions or approximated number formats without the options', async () => {
    const bytes = await kitchenSink();
    const without = await detectXlsxLoss(bytes);
    expect(ids(without)).not.toContain('xlsx.formulas.unsupported');
    expect(ids(without)).not.toContain('xlsx.numberFormats.approximated');
    expect(ids(without)).toContain('xlsx.formulas.structuredRefs');
  });

  it('survives a number-format or function option that throws, keeping what it found before', async () => {
    const bytes = await kitchenSink();
    const report = await detectXlsxLoss(bytes, {
      isNumberFormatExact: () => {
        throw new Error('boom');
      },
    });
    expect(ids(report)).toContain('xlsx.charts');
    expect(ids(report)).not.toContain('xlsx.numberFormats.approximated');
  });
});

describe('formulas', () => {
  it('reports unsupported functions with the names, prefixes stripped, once per formula', async () => {
    const rows =
      '<row r="1">' +
      cell('A1', '<f>SUM(1)+_xlfn.IFS(1,2)+_xlws.SORT(A1)+STDEV.S(1)+_xlfn.IFS(3)</f><v>1</v>') +
      cell('A2', '<f>SUM(1)+IF(1,2,3)</f><v>1</v>') +
      cell('A3', '<f>_xlfn.LET(_xlpm.x,1,_xlpm.x(2))</f><v>1</v>') +
      cell('A4', '<f t="shared" si="0"/><v>1</v>') +
      '</row>';
    const report = await detectXlsxLoss(await build({ sheets: [{ name: 'F', rows }] }), { knownFunctions: KNOWN });
    expect(item(report, 'xlsx.formulas.unsupported')).toMatchObject({ count: 2, where: ['IFS', 'SORT', 'STDEV.S', 'LET'] });
  });

  it('counts every cell of a shared formula with the facts of its master', async () => {
    const rows =
      '<row r="1">' +
      cell('A1', '<f t="shared" ref="A1:A3" si="0">SUM(Table1[a])+NOW()+FANCY(1)+SUM(A:A)+[1]S!A1</f><v>1</v>') +
      cell('A2', '<f t="shared" si="0"/><v>1</v>') +
      cell('A3', '<f t="shared" si="0"/><v>1</v>') +
      cell('A4', '<f t="shared" si="9"/><v>1</v>') +
      cell('A5', '<f t="shared"/><v>1</v>') +
      cell('B1', '<f t="shared" ref="B1:B2" si="1">SUM(1)</f><v>1</v>') +
      cell('B2', '<f t="shared" si="1"/><v>1</v>') +
      '</row>';
    const report = await full(await build({ sheets: [{ name: 'S', rows }] }), { knownFunctions: KNOWN });
    for (const id of ['structuredRefs', 'externalRef', 'volatile', 'unsupported']) expect(item(report, `xlsx.formulas.${id}`)?.count).toBe(3);
  });

  it('scans cells written with single-quoted attributes', async () => {
    const rows = "<row r='1'><c r='A1' s='1'><f>SUM(Table1[a])</f></c></row>";
    const saved = LIMITS.chunkBytes;
    LIMITS.chunkBytes = 10;
    try {
      const report = await full(await build({ sheets: [{ rows: rows + "<row r='2' outlineLevel='1'/>" + "<row r='3'><c r='A3' s='3'/></row>" }], styles: STYLES('<cellXfs><xf/><xf/><xf/><xf fontId="1"/></cellXfs><fonts><font/><font><color theme="4"/></font></fonts>') }));
      expect(item(report, 'xlsx.formulas.structuredRefs')?.count).toBe(1);
      expect(item(report, 'xlsx.outline')?.count).toBe(1);
      expect(item(report, 'xlsx.styles.themeColors')?.count).toBe(1);
    } finally {
      LIMITS.chunkBytes = saved;
    }
  });

  it('reports nothing when every function is known, and the formula text is not required', async () => {
    const bytes = await build({ sheets: [{ rows: f('SUM(1)+_xlfn.IF(1,2,3)') }] });
    expect(ids(await detectXlsxLoss(bytes, { knownFunctions: KNOWN }))).toEqual([]);
  });

  it('reports structured references (summary) and not the same text in a string or an external reference', async () => {
    const rows =
      '<row r="1">' +
      cell('A1', '<f>SUM(Table1[Amount])</f>') +
      cell('A2', '<f>Table1[[#This Row],[Amount]]*2</f>') +
      cell('A3', '<f>"Table1[x]"</f>') +
      cell('A4', '<f>[1]S!A1</f>') +
      '</row>';
    const report = await detectXlsxLoss(await build({ sheets: [{ name: 'T', rows }] }));
    expect(item(report, 'xlsx.formulas.structuredRefs')).toMatchObject({ count: 2, where: ['T'] });
  });

  it('reports array, dynamic-array, external and volatile formulas at the full level, and no longer reports whole-column references', async () => {
    const rows =
      '<row r="1">' +
      cell('A1', '<f t="array" ref="A1:A2">SUM(B1:B2)</f><v>1</v>') +
      cell('A2', '<f t="array" ref="A2:A3">SEQUENCE(2)</f><v>1</v>', ' cm="1"') +
      cell('A3', "<f>'[1]Sheet 1'!A1+[2]S!B2</f><v>1</v>") +
      cell('A4', '<f>SUM(A:A)+SUM(Sheet2!1:1)</f><v>1</v>') +
      cell('A5', '<f>"A:A"&amp;SUM(A1:B2)</f><v>1</v>') +
      cell('A6', '<f>NOW()+_xlfn.RANDARRAY(2)</f><v>1</v>') +
      cell('A7', '<f>OFFSET(A1,1,1)</f><v>1</v>') +
      cell('A8', '<f>SUM(1)</f><v>1</v>') +
      '</row>';
    const bytes = await build({ sheets: [{ name: 'F', rows }] });
    const report = await full(bytes);
    expect(item(report, 'xlsx.formulas.array')).toMatchObject({ count: 1, where: ['F'] });
    expect(item(report, 'xlsx.formulas.dynamicArray')).toMatchObject({ count: 1 });
    expect(item(report, 'xlsx.formulas.externalRef')).toMatchObject({ count: 1 });
    // A:A and 1:1 (cells A4) are calculated by the engine now, so they are not a loss and have no id.
    expect(item(report, 'xlsx.formulas.wholeColumn')).toBeUndefined();
    expect(report.items.some((i) => /whole/i.test(i.id))).toBe(false);
    expect(item(report, 'xlsx.formulas.volatile')).toMatchObject({ count: 2 });
    expect(ids(await detectXlsxLoss(bytes))).toEqual([]);
  });

  it('copes with a prefixed worksheet, a missing worksheet root, an empty and a truncated sheetData', async () => {
    const prefixed =
      `<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:sheetData><x:row r="1"><x:c r="A1"><x:f>SUM(Table1[a])</x:f></x:c></x:row></x:sheetData><x:autoFilter ref="A1"/></x:worksheet>`;
    const report = await detectXlsxLoss(await build({ sheets: [{ name: 'P', raw: prefixed }] }));
    expect(ids(report)).toEqual(['xlsx.autoFilter', 'xlsx.formulas.structuredRefs']);

    const noRoot = await detectXlsxLoss(await build({ sheets: [{ raw: '<sheetData><row r="1"><c r="A1"><f>Table1[a]</f></c></row></sheetData>' }] }));
    expect(ids(noRoot)).toEqual(['xlsx.formulas.structuredRefs']);

    const selfClosing = await build({ sheets: [{ raw: `<worksheet ${MAIN}><sheetData/><autoFilter ref="A1"/></worksheet>` }] });
    expect(ids(await detectXlsxLoss(selfClosing))).toEqual(['xlsx.autoFilter']);

    const noSheetData = await build({ sheets: [{ raw: `<worksheet ${MAIN}><autoFilter ref="A1"/></worksheet>` }] });
    expect(ids(await detectXlsxLoss(noSheetData))).toEqual(['xlsx.autoFilter']);

    const truncated = await build({
      sheets: [{ raw: `<worksheet ${MAIN}><sheetData><row r="1"><c r="A1"><f>SUM(Table1[a])</f></c></row><row r="2"><c r="A2"><f>Tab` }],
    });
    expect(ids(await detectXlsxLoss(truncated))).toEqual(['xlsx.formulas.structuredRefs']);
  });

  it('ends the last chunk at the end of the cell data even when markup after it starts with <row', async () => {
    const saved = LIMITS.chunkBytes;
    LIMITS.chunkBytes = 100;
    try {
      const rows = '<row r="1"><c r="A1" s="1"><f>SUM(Table1[a])</f></c></row>'.repeat(3);
      const withBreaks = await build({ sheets: [{ rows, tail: '<rowBreaks count="1"><brk id="3"/></rowBreaks>' }] });
      expect(item(await full(withBreaks), 'xlsx.formulas.structuredRefs')?.count).toBe(3);
      const alone = await build({ sheets: [{ rows }] });
      expect(item(await full(alone), 'xlsx.formulas.structuredRefs')?.count).toBe(3);
    } finally {
      LIMITS.chunkBytes = saved;
    }
  });

  it('skips a chunk of cells that cannot be parsed and keeps the rest', async () => {
    const deep = `<row r="1" s="1">${'<a>'.repeat(300)}</row>`;
    const ok = '<row r="2"><c r="A2" s="1"><f>SUM(Table1[a])</f></c></row>';
    const saved = LIMITS.chunkBytes;
    LIMITS.chunkBytes = 50;
    try {
      const report = await detectXlsxLoss(await build({ sheets: [{ rows: deep + ok }] }));
      expect(item(report, 'xlsx.formulas.structuredRefs')?.count).toBe(1);
    } finally {
      LIMITS.chunkBytes = saved;
    }
  });

  it('scans a large sheet in chunks and reports each formula once', async () => {
    const rows = Array.from({ length: 400 }, (_, i) => `<row r="${i + 1}"><c r="A${i + 1}" s="1"><f>SUM(Table1[a])</f><v>1</v></c></row>`).join('');
    const saved = LIMITS.chunkBytes;
    LIMITS.chunkBytes = 1000;
    try {
      const report = await detectXlsxLoss(await build({ sheets: [{ rows }], styles: STYLES('<cellXfs><xf/><xf/></cellXfs>') }));
      expect(item(report, 'xlsx.formulas.structuredRefs')?.count).toBe(400);
    } finally {
      LIMITS.chunkBytes = saved;
    }
  });

  it('skips chunks that hold nothing to report and stops at the cell-data limit, naming the sheet', async () => {
    const plain = Array.from({ length: 300 }, (_, i) => `<row r="${i + 1}"><c r="A${i + 1}"><v>1</v></c></row>`).join('');
    const withFormula = '<row r="301"><c r="A301"><f>SUM(Table1[a])</f></c></row>';
    const bytes = await build({ sheets: [{ name: 'Big', rows: plain + withFormula }] });
    const saved = { ...LIMITS };
    LIMITS.chunkBytes = 500;
    try {
      expect(item(await detectXlsxLoss(bytes), 'xlsx.formulas.structuredRefs')?.count).toBe(1);
      LIMITS.cellBytes = 2000;
      const limited = await full(bytes);
      expect(ids(limited)).toEqual(['xlsx.package.sizeLimit']);
      expect(item(limited, 'xlsx.package.sizeLimit')).toMatchObject({ count: 1, where: ['xl/worksheets/sheet1.xml'] });
      expect(ids(await detectXlsxLoss(bytes))).toEqual([]);
    } finally {
      Object.assign(LIMITS, saved);
    }
  });
});

describe('styles and number formats', () => {
  const styles = STYLES(
    '<numFmts><numFmt numFmtId="164" formatCode="0.00"/><numFmt numFmtId="165" formatCode="[$€-407] #,##0.00"/></numFmts>' +
      '<fonts><font><color theme="1"/></font><font><color theme="4"/></font></fonts>' +
      '<fills><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
      '<fill><patternFill patternType="lightGray"/></fill><fill><gradientFill><stop position="0"><color rgb="FF000000"/></stop></gradientFill></fill></fills>' +
      '<borders><border/><border><diagonal style="thin"/></border></borders>' +
      '<cellStyleXfs><xf/><xf/></cellStyleXfs>' +
      '<cellXfs><xf/><xf fontId="1" xfId="0"/><xf fillId="2" xfId="1"/><xf fillId="3"/><xf borderId="1"/><xf numFmtId="164"/><xf numFmtId="165"/><xf numFmtId="14"/><xf numFmtId="27"/><xf fontId="1"/></cellXfs>' +
      '<cellStyles><cellStyle name="Normal" xfId="0"/><cellStyle name="Good" xfId="1"/></cellStyles>'
  );
  const used = (...ids: number[]) =>
    `<row r="1">${ids.map((s, i) => `<c r="${String.fromCharCode(65 + i)}1" s="${s}"/>`).join('')}</row>`;

  it('reports colours, fills, diagonal borders and named styles only for formats in use', async () => {
    const bytes = await build({ styles, sheets: [{ rows: used(1, 2, 3, 4) }] });
    const report = await full(bytes);
    expect(item(report, 'xlsx.styles.themeColors')).toMatchObject({ count: 1 });
    expect(item(report, 'xlsx.styles.patternFill')).toMatchObject({ count: 1 });
    expect(item(report, 'xlsx.styles.gradientFill')).toMatchObject({ count: 1 });
    expect(item(report, 'xlsx.styles.diagonalBorder')).toMatchObject({ count: 1 });
    expect(item(report, 'xlsx.styles.namedStyles')).toMatchObject({ count: 1, where: ['Good'] });
    const none = await full(await build({ styles, sheets: [{ rows: used(0) }] }));
    expect(ids(none)).toEqual([]);
    expect(ids(await detectXlsxLoss(bytes))).toEqual([]);
  });

  it('counts the formats of rows and columns as in use, and the column and row styles themselves', async () => {
    const bytes = await build({
      styles,
      sheets: [
        {
          name: 'S',
          head: '<cols><col min="1" max="1" style="0"/><col min="2" max="3" style="1"/><col min="4" max="4"/></cols>',
          rows: '<row r="1" s="9" customFormat="1"><c r="A1"/></row><row r="2" s="9"/><row r="3" s="0" customFormat="1"/>',
        },
      ],
    });
    const report = await full(bytes);
    expect(item(report, 'xlsx.styles.columnRowStyles')).toMatchObject({ count: 2, where: ['S'] });
    // the column format (1) and the row format (9) are both themed
    expect(item(report, 'xlsx.styles.themeColors')).toMatchObject({ count: 2 });
  });

  it('reports number formats the editor cannot reproduce, by code, through the predicate', async () => {
    const bytes = await build({ styles, sheets: [{ rows: used(5, 6, 7, 8, 5) }] });
    const asked: string[] = [];
    const report = await detectXlsxLoss(bytes, {
      isNumberFormatExact: (code) => {
        asked.push(code);
        return code === '0.00';
      },
    });
    expect(item(report, 'xlsx.numberFormats.approximated')).toMatchObject({
      count: 3,
      where: ['[$€-407] #,##0.00', 'm/d/yyyy', 'built-in format 27'],
    });
    expect(asked.sort()).toEqual(['0.00', '[$€-407] #,##0.00', 'm/d/yyyy']);
    const exact = await detectXlsxLoss(bytes, { isNumberFormatExact: () => true });
    expect(item(exact, 'xlsx.numberFormats.approximated')).toMatchObject({ count: 1, where: ['built-in format 27'] });
  });

  it('is quiet about number formats on a workbook with no styles part', async () => {
    const report = await detectXlsxLoss(await build({ sheets: [{ rows: used(0, 3) }] }), { isNumberFormatExact: () => false });
    expect(ids(report)).toEqual([]);
  });
});

describe('sheet, print and workbook settings', () => {
  it('reports tab colour, zoom, right to left, gridlines off and custom views', async () => {
    const bytes = await build({
      sheets: [
        {
          name: 'A',
          head: '<sheetPr><tabColor theme="4"/></sheetPr><sheetViews><sheetView zoomScale="85" rightToLeft="true" showGridLines="false"/></sheetViews>',
          tail: '<customSheetViews><customSheetView guid="a"/><customSheetView guid="b"/></customSheetViews>',
        },
        { name: 'B', head: '<sheetViews><sheetView zoomScale="100" showGridLines="1" rightToLeft="0"/></sheetViews>' },
        { name: 'C', head: '<sheetViews><sheetView/></sheetViews>' },
        { name: 'D', state: 'veryHidden' },
      ],
      workbook: '<customWorkbookViews><customWorkbookView name="v" guid="g"/></customWorkbookViews>',
    });
    const report = await full(bytes);
    expect(item(report, 'xlsx.sheet.tabColor')).toMatchObject({ count: 1, where: ['A'] });
    expect(item(report, 'xlsx.sheet.zoom')).toMatchObject({ count: 1, where: ['A'] });
    expect(item(report, 'xlsx.sheet.rtl')).toMatchObject({ count: 1, where: ['A'] });
    expect(item(report, 'xlsx.sheet.gridlines')).toMatchObject({ count: 1, where: ['A'] });
    expect(item(report, 'xlsx.sheet.veryHidden')).toMatchObject({ count: 1, where: ['D'] });
    expect(item(report, 'xlsx.sheet.customViews')).toMatchObject({ count: 3, where: ['A'] });
    expect(ids(await detectXlsxLoss(bytes))).toEqual(['xlsx.hiddenSheets']);
  });

  it.each([
    ['landscape', '<pageSetup orientation="landscape"/>', ''],
    ['scale', '<pageSetup scale="50"/>', ''],
    ['paper size', '<pageSetup paperSize="5"/>', ''],
    ['first page number', '<pageSetup useFirstPageNumber="1"/>', ''],
    ['fit to page', '<pageSetup/>', '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>'],
    ['print options', '<printOptions gridLines="1" headings="0"/>', ''],
    ['margins', '<pageMargins left="1" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>', ''],
  ])('reports page setup for %s', async (_name, tail, head) => {
    const report = await full(await build({ sheets: [{ name: 'P', head, tail }] }));
    expect(item(report, 'xlsx.print.setup')).toMatchObject({ count: 1, where: ['P'] });
  });

  it('does not report Excel default page setup', async () => {
    const tail =
      '<printOptions horizontalCentered="0"/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>' +
      '<pageSetup paperSize="9" orientation="portrait" scale="100" fitToWidth="1" fitToHeight="1"/>';
    expect(ids(await full(await build({ sheets: [{ tail }] })))).toEqual([]);
    expect(ids(await full(await build({ sheets: [{ tail: '<pageSetup paperSize="1"/><pageMargins/>' }] })))).toEqual([]);
  });

  it('reports print areas and titles by sheet, page breaks and header/footer text', async () => {
    const bytes = await build({
      sheets: [
        {
          name: 'P',
          tail:
            '<headerFooter><oddHeader>&amp;CTitle</oddHeader><oddFooter></oddFooter></headerFooter>' +
            '<rowBreaks count="2"><brk id="3"/><brk id="9"/></rowBreaks><colBreaks count="1"><brk id="2"/></colBreaks>',
        },
        { name: 'Q', tail: '<headerFooter><oddHeader>   </oddHeader><extLst><ext/></extLst></headerFooter>' },
      ],
      workbook:
        '<definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">P!$A$1</definedName><definedName name="_xlnm.Print_Area">$A$1</definedName>' +
        '<definedName name="_xlnm.Print_Titles" localSheetId="1">Q!$1:$1</definedName><definedName name="_xlnm.Print_Titles" localSheetId="9">x</definedName></definedNames>',
    });
    const report = await full(bytes);
    expect(item(report, 'xlsx.print.areas')).toMatchObject({ count: 2, where: ['P', '_xlnm.Print_Area'] });
    expect(item(report, 'xlsx.print.titles')).toMatchObject({ count: 2, where: ['Q', '_xlnm.Print_Titles'] });
    expect(item(report, 'xlsx.print.breaks')).toMatchObject({ count: 3, where: ['P'] });
    expect(item(report, 'xlsx.print.headersFooters')).toMatchObject({ count: 1, where: ['P'] });
    expect(ids(await detectXlsxLoss(bytes))).toEqual([]);
  });

  it('reports 1904 dates and each non-default calculation setting', async () => {
    const report = await full(await build({ workbookHead: '<workbookPr date1904="true"/>', workbook: '<calcPr calcMode="autoNoTable" iterate="1" fullPrecision="0" refMode="R1C1"/>' }));
    expect(item(report, 'xlsx.workbook.date1904')).toMatchObject({ count: 1 });
    expect(item(report, 'xlsx.workbook.calcSettings')).toMatchObject({ count: 4 });
    const defaults = await full(await build({ workbookHead: '<workbookPr date1904="0"/>', workbook: '<calcPr calcId="191029" calcMode="auto" iterate="0" fullPrecision="1" refMode="A1"/>' }));
    expect(ids(defaults)).toEqual([]);
    expect(item(await full(await build({ workbook: '<calcPr calcMode="manual"/>' })), 'xlsx.workbook.calcSettings')?.count).toBe(1);
    expect(ids(await full(await build({ workbook: '<calcPr/>' })))).toEqual([]);
  });

  it('reports data connections, query tables and the data model', async () => {
    const connections = await full(await build({ files: { 'xl/connections.xml': '<connections><connection id="1"/><connection id="2"/></connections>' } }));
    expect(item(connections, 'xlsx.workbook.externalData')?.count).toBe(2);
    const tables = await full(await build({ files: { 'xl/queryTables/queryTable1.xml': '<queryTable/>', 'xl/queryTables/queryTable2.xml': '<queryTable/>', 'xl/model/item.data': 'x' } }));
    expect(item(tables, 'xlsx.workbook.externalData')?.count).toBe(3);
  });
});

describe('conditional formatting, validation, comments and links', () => {
  it('reports formula rules, stop-if-true and the Excel 2010 data bar and icon set extensions', async () => {
    const bytes = await build({
      sheets: [
        {
          name: 'C',
          tail:
            '<conditionalFormatting sqref="A1"><cfRule type="expression" priority="1" stopIfTrue="1"/><cfRule type="cellIs" priority="2" stopIfTrue="0"/><cfRule type="dataBar" priority="3"/></conditionalFormatting>' +
            `<extLst><ext uri="u" ${X14}><x14:conditionalFormattings>` +
            '<x14:conditionalFormatting><x14:cfRule type="dataBar" id="a"/><x14:cfRule type="iconSet" id="b" stopIfTrue="true"/><x14:cfRule type="expression" id="c"/><x14:cfRule type="cellIs" id="d"/></x14:conditionalFormatting>' +
            '</x14:conditionalFormattings></ext></extLst>',
        },
      ],
    });
    const report = await full(bytes);
    expect(item(report, 'xlsx.cf.formulaRule')).toMatchObject({ count: 2, where: ['C'] });
    expect(item(report, 'xlsx.cf.stopIfTrue')).toMatchObject({ count: 2 });
    expect(item(report, 'xlsx.cf.dataBarExt')).toMatchObject({ count: 1 });
    expect(item(report, 'xlsx.cf.iconSetExt')).toMatchObject({ count: 1 });
    // the summary counts the main-namespace rules only, as before
    expect(item(await detectXlsxLoss(bytes), 'xlsx.conditionalFormatting')?.count).toBe(1);
  });

  it('reports custom-formula validation, messages and references to another sheet (main and x14 markup)', async () => {
    const bytes = await build({
      sheets: [
        {
          name: 'V',
          tail:
            '<dataValidations count="4">' +
            '<dataValidation type="custom" sqref="A1"><formula1>A2&gt;0</formula1></dataValidation>' +
            '<dataValidation type="list" sqref="A2" error="bad" errorTitle="T"><formula1>Other!$A$1:$A$3</formula1></dataValidation>' +
            '<dataValidation type="list" sqref="A3" prompt=""><formula1>"a,b"</formula1></dataValidation>' +
            '<dataValidation type="whole" sqref="A4" promptTitle="P"><formula1>1</formula1><formula2>9</formula2></dataValidation>' +
            '</dataValidations>' +
            `<extLst><ext uri="u" ${X14}><x14:dataValidations><x14:dataValidation type="list" error="x"><x14:formula1><xm:f>Other!A1:A2</xm:f></x14:formula1><xm:sqref>B1</xm:sqref></x14:dataValidation><x14:dataValidation type="custom"/></x14:dataValidations></ext></extLst>`,
        },
      ],
    });
    const report = await full(bytes);
    expect(item(report, 'xlsx.dv.customFormula')).toMatchObject({ count: 2, where: ['V'] });
    expect(item(report, 'xlsx.dv.messages')).toMatchObject({ count: 3 });
    expect(item(report, 'xlsx.dv.crossSheet')).toMatchObject({ count: 3 });
    expect(item(await detectXlsxLoss(bytes), 'xlsx.dataValidation')?.count).toBe(4);
  });

  it('reports threaded comments and formatted comment text per sheet', async () => {
    const bytes = await build({
      sheets: [
        { name: 'N', rels: [rel('rId1', 'comments', '../comments1.xml'), rel('rId2', 'threadedComment', '../threadedComments/t.xml')] },
        { name: 'M', rels: [rel('rId1', 'comments', '../comments-missing.xml')] },
      ],
      files: {
        'xl/comments1.xml':
          '<comments><commentList><comment ref="A1"><text><r><rPr><b/></rPr><t>x</t></r></text></comment><comment ref="A2"><text><t>plain</t></text></comment><comment ref="A3"><text><r><t>y</t></r></text></comment></commentList></comments>',
        'xl/threadedComments/t.xml': '<ThreadedComments><threadedComment id="1"/><threadedComment id="2" parentId="1"/></ThreadedComments>',
      },
    });
    const report = await full(bytes);
    expect(item(report, 'xlsx.comments.threaded')).toMatchObject({ count: 2, where: ['N'] });
    expect(item(report, 'xlsx.comments.richFormatting')).toMatchObject({ count: 1, where: ['N'] });
    expect(item(report, 'xlsx.comments')).toMatchObject({ count: 2, where: ['N', 'M'] });
  });

  it('reports internal links and script schemes, and not ordinary external links', async () => {
    const bytes = await build({
      sheets: [
        {
          name: 'L',
          tail:
            '<hyperlinks>' +
            '<hyperlink ref="A1" location="Sheet2!A1"/>' +
            '<hyperlink ref="A2" r:id="rId1"/><hyperlink ref="A3" r:id="rId2"/><hyperlink ref="A4" r:id="rId3"/><hyperlink ref="A5" r:id="rId4"/>' +
            '<hyperlink ref="A6" r:id="rId9" location="X!A1"/><hyperlink ref="A7" r:id="rId9"/><hyperlink ref="A8"/>' +
            '<hyperlink ref="A9" r:id="rId5"/><hyperlink ref="A10" r:id="rId6" location="Z!A1"/></hyperlinks>',
          rels: [
            rel('rId1', 'hyperlink', 'https://example.com', true),
            rel('rId2', 'hyperlink', 'JavaScript:alert(1)', true),
            rel('rId3', 'hyperlink', '  vbscript:x', true),
            rel('rId4', 'hyperlink', '&#1;&#9;data:text/html,hi', true),
            rel('rId5', 'hyperlink', 'file:///etc/passwd', true),
            rel('rId6', 'hyperlink', 'javascript:alert(2)', true),
          ],
        },
      ],
    });
    const report = await full(bytes);
    // A1 and A6 have a location and no usable relationship (A6's rId9 is missing): internal; A7/A8 have neither
    expect(item(report, 'xlsx.links.internal')).toMatchObject({ count: 2, where: ['L'] });
    expect(item(report, 'xlsx.links.unsafeScheme')).toMatchObject({ count: 5, where: ['L'] });
    expect(item(report, 'xlsx.links.unsafeScheme')?.effect).toMatch(/^kept: /);
  });

  it('checks the target as written, so a script link without TargetMode="External" is still reported', async () => {
    const bytes = await build({
      sheets: [{ tail: '<hyperlinks><hyperlink ref="A1" r:id="rId1"/><hyperlink ref="A2" r:id="rId2"/></hyperlinks>', rels: [rel('rId1', 'hyperlink', 'javascript:x'), rel('rId2', 'hyperlink', 'other.xlsx')] }],
    });
    expect(item(await full(bytes), 'xlsx.links.unsafeScheme')?.count).toBe(1);
  });

  it('ignores an exclamation mark inside a string literal when looking for references to another sheet', async () => {
    const bytes = await build({
      sheets: [{ tail: '<dataValidations><dataValidation type="custom" sqref="A1"><formula1>A2="Hi!"</formula1></dataValidation><dataValidation type="list" sqref="A2"><formula1>Other!A1:A3</formula1></dataValidation></dataValidations>' }],
    });
    expect(item(await full(bytes), 'xlsx.dv.crossSheet')?.count).toBe(1);
  });

  it('does not call a note formatted for the author line and plain run properties that every note has', async () => {
    const note = (runs: string) => `<comment ref="A1"><text>${runs}</text></comment>`;
    const author = '<r><rPr><b/><sz val="9"/><rFont val="Tahoma"/></rPr><t>Author:\n</t></r>';
    const body = (extra = '') => `<r><rPr><sz val="9"/>${extra}<rFont val="Tahoma"/></rPr><t>text</t></r>`;
    const bytes = await build({
      sheets: [{ name: 'N', rels: [rel('rId1', 'comments', '../comments1.xml')] }],
      files: {
        'xl/comments1.xml': `<comments><commentList>${[
          note(author + body()), // Excel's shape: not formatted
          note(body()), // LibreOffice's shape: not formatted
          note(author + body('<i/>')), // italic text: formatted
          note(body() + '<r><rPr><sz val="14"/><rFont val="Tahoma"/></rPr><t>big</t></r>'), // two different run formats
          note(body() + body()), // the same format twice: not formatted
          '<comment ref="B1"/>',
        ].join('')}</commentList></comments>`,
      },
    });
    expect(item(await full(bytes), 'xlsx.comments.richFormatting')?.count).toBe(2);
  });
});

describe('drawing objects', () => {
  const anchor = (inner: string) => `<xdr:twoCellAnchor>${inner}</xdr:twoCellAnchor>`;
  const sp = (id: number, nvSpPr = '') => anchor(`<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="${id}" name="S"/>${nvSpPr || '<xdr:cNvSpPr/>'}</xdr:nvSpPr></xdr:sp>`);
  const drawing = (inner: string) => `<xdr:wsDr ${XDR}>${inner}</xdr:wsDr>`;
  const sheetWith = (inner: string, tail = '') => build({
    sheets: [{ name: 'D', tail, rels: [rel('rId1', 'drawing', '../drawings/drawing1.xml')] }],
    files: { 'xl/drawings/drawing1.xml': drawing(inner) },
  });

  it('reports shapes, connectors and text boxes separately, and shapes nested in groups', async () => {
    const bytes = await sheetWith(
      sp(2) +
        sp(3, '<xdr:cNvSpPr txBox="1"/>') +
        sp(4, '<xdr:cNvSpPr txBox="0"/>') +
        anchor('<xdr:cxnSp><xdr:nvCxnSpPr><xdr:cNvPr id="5" name="C"/></xdr:nvCxnSpPr></xdr:cxnSp>') +
        anchor('<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="6" name="G"/></xdr:nvGrpSpPr><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="7" name="in"/><xdr:cNvSpPr/></xdr:nvSpPr></xdr:sp></xdr:grpSp>') +
        anchor('<xdr:graphicFrame><a:graphic><a:graphicData uri="dgm"/></a:graphic></xdr:graphicFrame>') +
        anchor('<xdr:sp><xdr:txBody/></xdr:sp>')
    );
    const report = await full(bytes);
    expect(item(report, 'xlsx.objects.shapes')).toMatchObject({ count: 5, where: ['D'] });
    expect(item(report, 'xlsx.objects.textBoxes')).toMatchObject({ count: 1 });
    expect(ids(report)).not.toContain('xlsx.charts');
  });

  it('counts a chart once even when its markup is repeated in a compatibility fallback', async () => {
    const chart = '<xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="C"/></xdr:nvGraphicFramePr><a:graphic><a:graphicData uri="c"><c:chart/></a:graphicData></a:graphic></xdr:graphicFrame>';
    const bytes = await sheetWith(`<mc:AlternateContent><mc:Choice Requires="cx">${anchor(chart)}</mc:Choice><mc:Fallback>${anchor(chart)}</mc:Fallback></mc:AlternateContent>`);
    expect(item(await detectXlsxLoss(bytes), 'xlsx.charts')).toMatchObject({ count: 1 });
  });

  it('does not count a control or an embedded object a second time as a shape or picture', async () => {
    const bytes = await build({
      sheets: [
        {
          name: 'D',
          tail:
            '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="x14"><controls>' +
            '<mc:AlternateContent><mc:Choice Requires="x14"><control shapeId="10" r:id="rId2"/></mc:Choice><mc:Fallback><control shapeId="10"/></mc:Fallback></mc:AlternateContent>' +
            '<mc:AlternateContent><mc:Choice Requires="x14"><control r:id="rId3"/></mc:Choice></mc:AlternateContent>' +
            '</controls></mc:Choice></mc:AlternateContent>' +
            '<oleObjects><oleObject progId="Word.Document" shapeId="11" r:id="rId4"/><oleObject progId="x"/></oleObjects>',
          rels: [rel('rId1', 'drawing', '../drawings/drawing1.xml')],
        },
      ],
      files: {
        'xl/drawings/drawing1.xml': drawing(
          sp(10) + anchor('<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="11" name="ole"/></xdr:nvPicPr></xdr:pic>') + sp(12) + anchor('<xdr:pic><xdr:nvPicPr><xdr:cNvPr name="noid"/></xdr:nvPicPr></xdr:pic>')
        ),
      },
    });
    const report = await full(bytes);
    expect(item(report, 'xlsx.objects.formControls')).toMatchObject({ count: 2, where: ['D'] });
    expect(item(report, 'xlsx.objects.embeddedObjects')).toMatchObject({ count: 2, where: ['D'] });
    // what is left of the drawing: one shape and the one picture with no id
    expect(item(report, 'xlsx.objects.shapes')).toMatchObject({ count: 1 });
    expect(item(report, 'xlsx.images')).toMatchObject({ count: 1 });
  });

  it('falls back to the control and embedding parts when the sheet does not list them', async () => {
    const bytes = await build({ files: { 'xl/ctrlProps/ctrlProp1.xml': '<formControlPr/>', 'xl/ctrlProps/ctrlProp2.xml': '<formControlPr/>', 'xl/embeddings/oleObject1.bin': 'x' } });
    const report = await full(bytes);
    expect(item(report, 'xlsx.objects.formControls')?.count).toBe(2);
    expect(item(report, 'xlsx.objects.embeddedObjects')?.count).toBe(1);
  });

  it('reports pictures in cells and other rich values from the rich data parts', async () => {
    const bytes = await build({
      files: {
        'xl/richData/rdrichvaluestructure.xml': '<rvStructures count="2"><s t="_localImage"><k n="x" t="i"/></s><s t="_stock"/></rvStructures>',
        'xl/richData/rdrichvalue.xml': '<rvData><rv s="0"/><rv s="0"/><rv s="1"/><rv s="7"/></rvData>',
      },
    });
    const report = await full(bytes);
    expect(item(report, 'xlsx.objects.cellImages')?.count).toBe(2);
    expect(item(report, 'xlsx.workbook.richData')?.count).toBe(2);
    expect(item(report, 'xlsx.images')?.count).toBe(2);
    const unreadable = await full(await build({ files: { 'xl/richData/other.xml': '<x/>' } }));
    expect(item(unreadable, 'xlsx.workbook.richData')?.count).toBe(1);
    const legacy = await full(await build({ files: { 'xl/cellimages.xml': '<etc:cellImages xmlns:etc="e"><etc:cellImage/><etc:cellImage/></etc:cellImages>' } }));
    expect(item(legacy, 'xlsx.objects.cellImages')?.count).toBe(2);
  });
});

// ---------------------------------------------------------------------------
// Package level
// ---------------------------------------------------------------------------

describe('package level', () => {
  const OLE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
  const utf16 = (s: string) => [...s].flatMap((c) => [c.charCodeAt(0), 0]);

  it('names the parts nothing reads or reports, once, and not the known ones', async () => {
    const bytes = await build({
      files: {
        'docProps/core.xml': '<x/>',
        'xl/theme/theme1.xml': '<x/>',
        'xl/calcChain.xml': '<x/>',
        'customXml/item1.xml': '<x/>',
        'xl/dialogsheets/sheet1.xml': '<x/>',
        'xl/diagrams/data1.xml': '<x/>',
      },
    });
    const report = await full(bytes);
    expect(item(report, 'xlsx.package.unaccountedParts')).toMatchObject({
      count: 3,
      where: ['customXml/item1.xml', 'xl/dialogsheets/sheet1.xml', 'xl/diagrams/data1.xml'],
    });
    expect(ids(await detectXlsxLoss(bytes))).toEqual([]);
  });

  it('caps the part names at five but counts them all', async () => {
    const files = Object.fromEntries(Array.from({ length: 8 }, (_, i) => [`extra/p${i}.xml`, '<x/>']));
    const report = await full(await build({ files }));
    expect(item(report, 'xlsx.package.unaccountedParts')?.count).toBe(8);
    expect(item(report, 'xlsx.package.unaccountedParts')?.where).toHaveLength(5);
  });

  it('does not call a zip without a workbook an unaccounted package (it is not a workbook at all)', async () => {
    const zip = new JSZip();
    zip.file('mimetype', 'application/vnd.oasis.opendocument.spreadsheet');
    zip.file('content.xml', '<office:document-content/>');
    zip.file('dir/', '');
    expect(await full(await zip.generateAsync({ type: 'uint8array' }))).toEqual({ format: 'xlsx', items: [] });
  });

  it('recognises an encrypted package and a legacy binary workbook by their OLE container', async () => {
    const encrypted = await full(new Uint8Array([...OLE, 0, 0, ...utf16('EncryptionInfo'), ...utf16('EncryptedPackage')]));
    expect(ids(encrypted)).toEqual(['xlsx.package.encrypted']);
    expect(encrypted.items[0].effect).toMatch(/^lost: /);
    const legacy = await full(new Uint8Array([...OLE, 0, 0, ...utf16('Workbook')]));
    expect(ids(legacy)).toEqual(['xlsx.package.legacyBinary']);
    // summary detail stays quiet: the importer reports these files by failing with a named error
    expect(await detectXlsxLoss(new Uint8Array([...OLE, 0, 0, ...utf16('Workbook')]))).toEqual({ format: 'xlsx', items: [] });
  });

  it('recognises a binary (.xlsb) workbook', async () => {
    const report = await full(await build({ files: { 'xl/workbook.bin': 'x' } }));
    expect(item(report, 'xlsx.package.legacyBinary')).toMatchObject({ count: 1 });
  });

  it('never throws on empty, garbage, truncated or non-workbook input', async () => {
    const empty = { format: 'xlsx', items: [] };
    expect(await detectXlsxLoss(new Uint8Array(0))).toEqual(empty);
    expect(await full(new Uint8Array(0))).toEqual(empty);
    expect(await full(new Uint8Array([1, 2, 3]))).toEqual(empty);
    expect(await full(new TextEncoder().encode('PK\u0003\u0004 not really a zip'))).toEqual(empty);
    expect(await full(new Uint8Array(4096).map((_, i) => (i * 131) & 255))).toEqual(empty);
    const bytes = await build({ sheets: [{ tail: '<autoFilter ref="A1"/>' }] });
    expect(await full(bytes.slice(0, Math.floor(bytes.length / 2)))).toEqual(empty);
    expect(await detectXlsxLoss(bytes.slice(0, 30))).toEqual(empty);
  });

  it('reports what it can when a part is malformed, a relationship dangles or a stream is corrupt', async () => {
    const bytes = await build({
      sheets: [
        { name: 'Bad', raw: 'this is <not xml', rels: ['<Relationship Id="r1" Type="x/drawing" Target="../drawings/missing.xml"/>'] },
        { name: 'Good', tail: '<autoFilter ref="A1"/>' },
      ],
      workbook: '<definedNames><definedName name="N">x</definedName></definedNames>',
      styles: '<<<',
      sharedStrings: '</si>',
    });
    const report = await full(bytes);
    expect(ids(report)).toEqual(['xlsx.autoFilter', 'xlsx.definedNames']);

    const zip = new JSZip();
    zip.file('xl/workbook.xml', `<workbook ${MAIN}><sheets><sheet name="A" sheetId="1" r:id="rId1"/></sheets><definedNames><definedName name="N">x</definedName></definedNames></workbook>`);
    zip.file('xl/_rels/workbook.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`);
    zip.file('xl/worksheets/sheet1.xml', `<worksheet ${MAIN}>${'<sheetData/>'.repeat(200)}</worksheet>`);
    const good = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
    // corrupt the compressed stream of the sheet part: flip bytes in the middle of the archive
    const damaged = good.slice();
    const at = good.length >> 1;
    for (let i = 0; i < 16; i++) damaged[at + i] ^= 0xff;
    const out = await full(damaged);
    expect(out.format).toBe('xlsx');
  });

  it('goes on when a part cannot be read from the archive', async () => {
    const bytes = await build({ sheets: [{ name: 'Broken', tail: '<autoFilter ref="A1"/>' }, { name: 'Fine', tail: '<autoFilter ref="A1"/>' }] });
    const real = JSZip.loadAsync.bind(JSZip);
    const spy = jest.spyOn(JSZip, 'loadAsync').mockImplementation(async (data: Parameters<typeof JSZip.loadAsync>[0]) => {
      const zip = await real(data);
      const broken = zip.file('xl/worksheets/sheet1.xml') as unknown as { internalStream: () => unknown };
      const failing = { on(event: string, fn: () => void) { if (event === 'error') setTimeout(fn, 0); return failing; }, pause: () => failing, resume: () => failing };
      broken.internalStream = () => failing;
      return zip;
    });
    try {
      const report = await detectXlsxLoss(bytes);
      expect(item(report, 'xlsx.autoFilter')).toMatchObject({ count: 1, where: ['Fine'] });
    } finally {
      spy.mockRestore();
    }
  });

  it('reads a relationship with missing attributes and resolves parent-relative targets', async () => {
    const bytes = await build({
      sheets: [{ name: 'D', rels: ['<Relationship Id="rId1" Type="x/drawing" Target="../drawings/./drawing1.xml"/>', '<Relationship/>', `<Relationship Id="rId5" Type="${REL}/drawing" Target="/xl/drawings/drawing1.xml"/>`] }],
      files: { 'xl/drawings/drawing1.xml': `<xdr:wsDr ${XDR}><xdr:twoCellAnchor><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="P"/></xdr:nvPicPr></xdr:pic></xdr:twoCellAnchor></xdr:wsDr>` },
    });
    // the same drawing is reachable through two relationships: it is counted once per relationship
    expect(item(await detectXlsxLoss(bytes), 'xlsx.images')?.count).toBe(2);
  });

  it('refuses entity expansion and external entities: the markup is read, the entities are not', async () => {
    const doctype =
      '<?xml version="1.0"?><!DOCTYPE worksheet [<!ENTITY lol "lol"><!ENTITY lol2 "&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;"><!ENTITY ext SYSTEM "file:///etc/passwd">]>';
    const bytes = await build({
      sheets: [{ name: '&lol2;', raw: `${doctype}<worksheet ${MAIN}><sheetData/><sheetProtection sheet="1"/><autoFilter ref="&ext;"/></worksheet>` }],
      workbook: '<definedNames><definedName name="&lol2;">x</definedName></definedNames>',
    });
    const report = await detectXlsxLoss(bytes);
    expect(ids(report)).toEqual(['xlsx.autoFilter', 'xlsx.definedNames', 'xlsx.sheetProtection']);
    // the name comes back exactly as written, not expanded to 'lollollol...'
    expect(item(report, 'xlsx.definedNames')?.where).toEqual(['&lol2;']);
  });

  it('does not stack-overflow or hang on deeply nested XML', async () => {
    const deep = `<worksheet ${MAIN}>${'<a>'.repeat(5000)}<autoFilter ref="A1"/>${'</a>'.repeat(5000)}</worksheet>`;
    const report = await detectXlsxLoss(await build({ sheets: [{ raw: deep }], workbook: '<workbookProtection lockStructure="1"/>' }));
    expect(ids(report)).toEqual(['xlsx.workbookProtection']);
  });

  it('counts a very large number of features quickly', async () => {
    const rules = '<conditionalFormatting sqref="A1"><cfRule type="cellIs" priority="1"/></conditionalFormatting>'.repeat(20000);
    const started = Date.now();
    const report = await detectXlsxLoss(await build({ sheets: [{ tail: rules }] }));
    expect(item(report, 'xlsx.conditionalFormatting')?.count).toBe(20000);
    expect(Date.now() - started).toBeLessThan(10000);
  });

  it('skips a part that is over the size limit and names it at the full level', async () => {
    const saved = { ...LIMITS };
    LIMITS.partBytes = 600;
    try {
      const bytes = await build({
        sheets: [{ name: 'Huge', tail: `<autoFilter ref="A1"/><!-- ${'x'.repeat(2000)} -->` }, { name: 'Small', tail: '<autoFilter ref="A1"/>' }],
      });
      const summary = await detectXlsxLoss(bytes);
      expect(item(summary, 'xlsx.autoFilter')).toMatchObject({ count: 1, where: ['Small'] });
      expect(ids(summary)).not.toContain('xlsx.package.sizeLimit');
      const report = await full(bytes);
      expect(item(report, 'xlsx.package.sizeLimit')).toMatchObject({ count: 1, where: ['xl/worksheets/sheet1.xml'] });
    } finally {
      Object.assign(LIMITS, saved);
    }
  });

  it('stops reading once the total size budget is spent', async () => {
    const saved = { ...LIMITS };
    const filler = `<!-- ${'x'.repeat(700)} -->`;
    const bytes = await build({
      sheets: ['A', 'B', 'C'].map((name) => ({ name, tail: `<autoFilter ref="A1"/>${filler}` })),
    });
    LIMITS.totalBytes = 2000;
    try {
      const report = await full(bytes);
      expect(item(report, 'xlsx.autoFilter')).toMatchObject({ count: 1, where: ['A'] });
      expect(item(report, 'xlsx.package.sizeLimit')).toMatchObject({ count: 2, where: ['xl/worksheets/sheet2.xml', 'xl/worksheets/sheet3.xml'] });
    } finally {
      Object.assign(LIMITS, saved);
    }
  });

  it('stops a forged-header bomb: the inflated size is counted as the part comes out', async () => {
    const saved = LIMITS.partBytes;
    LIMITS.partBytes = 100_000;
    try {
      const zip = new JSZip();
      zip.file('xl/workbook.xml', `<workbook ${MAIN}><sheets><sheet name="A" sheetId="1" r:id="rId1"/></sheets></workbook>`);
      zip.file('xl/_rels/workbook.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`);
      zip.file('xl/worksheets/sheet1.xml', `<worksheet ${MAIN}><!--${' '.repeat(8_000_000)}--></worksheet>`);
      const bytes = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
      // forge the uncompressed size (central directory and local header) of the sheet part down to 1000 bytes
      const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
      let forged = 0;
      for (let i = 0; i + 30 < bytes.length; i++) {
        const sig = view.getUint32(i, true);
        if (sig === 0x02014b50 && view.getUint32(i + 24, true) > 1_000_000) { view.setUint32(i + 24, 1000, true); forged++; }
        if (sig === 0x04034b50 && view.getUint32(i + 22, true) > 1_000_000) { view.setUint32(i + 22, 1000, true); forged++; }
      }
      expect(forged).toBeGreaterThan(0);
      const started = Date.now();
      const report = await full(bytes);
      expect(item(report, 'xlsx.package.sizeLimit')).toMatchObject({ count: 1, where: ['xl/worksheets/sheet1.xml'] });
      expect(Date.now() - started).toBeLessThan(5000);
    } finally {
      LIMITS.partBytes = saved;
    }
  });

  it('reads and counts a part shared by several sheets once', async () => {
    const saved = LIMITS.totalBytes;
    const pic = '<xdr:oneCellAnchor><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="P"/></xdr:nvPicPr></xdr:pic></xdr:oneCellAnchor>';
    const drawing = `<xdr:wsDr ${XDR}>${pic}<!-- ${'x'.repeat(2000)} --></xdr:wsDr>`;
    const bytes = await build({
      sheets: ['A', 'B', 'C', 'D'].map((name) => ({ name, rels: [rel('rId1', 'drawing', '../drawings/drawing1.xml')] })),
      files: { 'xl/drawings/drawing1.xml': drawing },
    });
    LIMITS.totalBytes = 8000; // the workbook parts and one copy of the drawing fit; four copies would not
    try {
      const report = await full(bytes);
      expect(item(report, 'xlsx.images')?.count).toBe(4);
      expect(item(report, 'xlsx.package.sizeLimit')).toBeUndefined();
    } finally {
      LIMITS.totalBytes = saved;
    }
  });

  it('inflates nothing from a zip bomb: the declared size is checked before the part is read', async () => {
    const zip = new JSZip();
    zip.file('xl/workbook.xml', `<workbook ${MAIN}><sheets><sheet name="A" sheetId="1" r:id="rId1"/></sheets></workbook>`);
    zip.file('xl/_rels/workbook.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`);
    zip.file('xl/worksheets/sheet1.xml', `<worksheet ${MAIN}><!--${' '.repeat(LIMITS.partBytes + 1024)}--></worksheet>`);
    const bytes = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
    expect(bytes.length).toBeLessThan(200_000);
    const started = Date.now();
    const report = await full(bytes);
    expect(item(report, 'xlsx.package.sizeLimit')).toMatchObject({ count: 1, where: ['xl/worksheets/sheet1.xml'] });
    expect(Date.now() - started).toBeLessThan(5000);
  });

  it('checks the size after reading when the archive gives no declared size up front', async () => {
    const saved = LIMITS.partBytes;
    const real = JSZip.loadAsync.bind(JSZip);
    const spy = jest.spyOn(JSZip, 'loadAsync').mockImplementation(async (data: Parameters<typeof JSZip.loadAsync>[0]) => {
      const zip = await real(data);
      for (const file of Object.values(zip.files)) {
        const record = (file as unknown as { _data?: { uncompressedSize?: number } })._data;
        if (!record) continue;
        // the first look (the detector's) finds nothing; JSZip's own consistency check afterwards sees the real size
        const real = record.uncompressedSize;
        let looks = 0;
        Object.defineProperty(record, 'uncompressedSize', { get: () => (looks++ === 0 ? undefined : real) });
      }
      return zip;
    });
    try {
      const bytes = await build({
        sheets: [{ name: 'Huge', tail: `<autoFilter ref="A1"/><!-- ${'x'.repeat(4000)} -->` }, { name: 'Small', tail: '<autoFilter ref="A1"/>' }],
      });
      LIMITS.partBytes = 1500;
      const report = await full(bytes);
      expect(item(report, 'xlsx.autoFilter')).toMatchObject({ count: 1, where: ['Small'] });
      expect(item(report, 'xlsx.package.sizeLimit')).toMatchObject({ count: 1, where: ['xl/worksheets/sheet1.xml'] });
    } finally {
      LIMITS.partBytes = saved;
      spy.mockRestore();
    }
  });
});

// ---------------------------------------------------------------------------
// Each full-level id has a test above; this table keeps the list of ids honest.
// ---------------------------------------------------------------------------

describe('full id coverage', () => {
  const FULL_IDS = [
    'xlsx.formulas.array', 'xlsx.formulas.dynamicArray', 'xlsx.formulas.externalRef', 'xlsx.formulas.volatile',
    'xlsx.styles.themeColors', 'xlsx.styles.patternFill', 'xlsx.styles.gradientFill', 'xlsx.styles.diagonalBorder', 'xlsx.styles.namedStyles', 'xlsx.styles.columnRowStyles',
    'xlsx.sheet.tabColor', 'xlsx.sheet.zoom', 'xlsx.sheet.rtl', 'xlsx.sheet.gridlines', 'xlsx.sheet.veryHidden', 'xlsx.sheet.customViews',
    'xlsx.print.setup', 'xlsx.print.areas', 'xlsx.print.titles', 'xlsx.print.breaks', 'xlsx.print.headersFooters',
    'xlsx.cf.dataBarExt', 'xlsx.cf.iconSetExt', 'xlsx.cf.formulaRule', 'xlsx.cf.stopIfTrue',
    'xlsx.dv.customFormula', 'xlsx.dv.messages', 'xlsx.dv.crossSheet',
    'xlsx.comments.threaded', 'xlsx.comments.richFormatting', 'xlsx.links.internal', 'xlsx.links.unsafeScheme',
    'xlsx.objects.shapes', 'xlsx.objects.textBoxes', 'xlsx.objects.formControls', 'xlsx.objects.embeddedObjects',
    'xlsx.workbook.date1904', 'xlsx.workbook.calcSettings', 'xlsx.workbook.externalData',
    'xlsx.package.unaccountedParts',
  ];

  it('the kitchen-sink workbook reports every full id that a plain zip can show', async () => {
    const report = await full(await kitchenSink(), { knownFunctions: KNOWN, isNumberFormatExact: () => false });
    for (const id of FULL_IDS) expect(ids(report)).toContain(id);
  });
});
