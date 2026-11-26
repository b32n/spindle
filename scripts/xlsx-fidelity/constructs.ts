import { isWellFormed, relationshipTargets } from './ooxml-package';
import type { OoxmlParts } from './ooxml-package';
import type { Construct } from './types';

/** XLSX construct registry: each probe counts one construct in the package XML (0 = absent). */
export type XlsxConstruct = Construct<OoxmlParts>;

const count = (s: string, re: RegExp) => s.match(re)?.length ?? 0;
const part = (p: OoxmlParts, name: string) => p.xml[name] ?? '';
const SHEET = /^xl\/worksheets\/[^/]+\.xml$/;
const sheetXml = (p: OoxmlParts) =>
  Object.entries(p.xml)
    .filter(([n]) => SHEET.test(n))
    .map(([, x]) => x);
const inSheets = (p: OoxmlParts, re: RegExp) =>
  sheetXml(p).reduce((sum, x) => sum + count(x, re), 0);
const inParts = (p: OoxmlParts, names: RegExp, re: RegExp) =>
  Object.entries(p.xml).reduce((sum, [n, x]) => sum + (names.test(n) ? count(x, re) : 0), 0);
const countNames = (p: OoxmlParts, re: RegExp) => p.names.filter((n) => re.test(n)).length;
const workbook = (p: OoxmlParts) => part(p, 'xl/workbook.xml');
const styles = (p: OoxmlParts) => part(p, 'xl/styles.xml');
const block = (xml: string, tag: string) =>
  new RegExp(`<${tag}\\b[^>]*>[\\s\\S]*?</${tag}>`).exec(xml)?.[0] ?? '';
const drawings = /^xl\/drawings\/drawing\d*\.xml$/;
const charts = /^xl\/charts\/chart\d*\.xml$/;
const textParts = (p: OoxmlParts) => [part(p, 'xl/sharedStrings.xml'), ...sheetXml(p)].join('\n');

interface Formula {
  cell: string;
  cellAttrs: string;
  attrs: string;
  text: string;
}
function formulas(p: OoxmlParts): Formula[] {
  const out: Formula[] = [];
  for (const x of sheetXml(p)) {
    for (const m of x.matchAll(/<c\b([^>]*)>\s*<f\b([^>]*?)(?:\/>|>([\s\S]*?)<\/f>)/g)) {
      out.push({
        cell: /\br="([A-Z]+\d+)"/.exec(m[1])?.[1] ?? '',
        cellAttrs: m[1],
        attrs: m[2],
        text: (m[3] ?? '')
          .replace(/&lt;/g, '<')
          .replace(/&gt;/g, '>')
          .replace(/&quot;/g, '"')
          .replace(/&apos;/g, "'")
          .replace(/&amp;/g, '&')
          .replace(/"(?:[^"]|"")*"/g, '""'), // string literals are not references
      });
    }
  }
  return out;
}
const withText = (p: OoxmlParts, re: RegExp) => formulas(p).filter((f) => re.test(f.text)).length;

const DATE_IDS = new Set([
  14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 45, 46, 47, 50, 51,
  52, 53, 54, 55, 56, 57, 58,
]);
const attr = (tag: string, name: string) => new RegExp(`\\b${name}="([^"]*)"`).exec(tag)?.[1];
/** Number of cells whose cell format is a date or time format (built-in id or a custom code with date tokens). */
function dateCells(p: OoxmlParts): number {
  const st = styles(p);
  const codes = new Map<number, string>();
  for (const m of st.matchAll(/<numFmt\b[^>]*>/g)) {
    const id = attr(m[0], 'numFmtId');
    const code = attr(m[0], 'formatCode');
    if (id !== undefined && code !== undefined) codes.set(Number(id), code);
  }
  const xfs = [...block(st, 'cellXfs').matchAll(/<xf\b[^>]*>/g)].map((m) =>
    Number(attr(m[0], 'numFmtId') ?? 0)
  );
  const isDate = xfs.map(
    (id) =>
      DATE_IDS.has(id) ||
      /[ymdhs]/i.test(
        (codes.get(id) ?? '').replace(/&quot;/g, '"').replace(/"[^"]*"|\\.|\[[^\]]*\]|General/g, '')
      )
  );
  return sheetXml(p).reduce(
    (sum, x) =>
      sum + [...x.matchAll(/<c\b[^>]*\bs="(\d+)"/g)].filter((m) => isDate[Number(m[1])]).length,
    0
  );
}

const colNumber = (letters: string) =>
  [...letters].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
/** Cells covered by the largest `<dimension ref>` in the package. */
function maxDimension(p: OoxmlParts): { rows: number; cells: number } {
  let best = { rows: 0, cells: 0 };
  for (const x of sheetXml(p)) {
    const m = /<dimension\b[^>]*\bref="([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?"/.exec(x);
    if (!m) continue;
    const rows = Number(m[4] ?? m[2]) - Number(m[2]) + 1;
    const cells = rows * (colNumber(m[3] ?? m[1]) - colNumber(m[1]) + 1);
    if (cells > best.cells) best = { rows, cells };
  }
  return best;
}

const cf = (type: string) => (p: OoxmlParts) =>
  inSheets(p, new RegExp(`<(?:x14:)?cfRule\\b[^>]*\\btype="(?:${type})"`, 'g'));
const dv = (type: string) => (p: OoxmlParts) =>
  inSheets(p, new RegExp(`<(?:x14:)?dataValidation\\b[^>]*\\btype="${type}"`, 'g'));
const chart = (tag: string) => (p: OoxmlParts) =>
  inParts(p, charts, new RegExp(`<(?:c:)?${tag}\\b`, 'g'));
const media = (ext: string) => (p: OoxmlParts) =>
  countNames(p, new RegExp(`^xl/media/[^/]+\\.(?:${ext})$`, 'i'));
const anchor = (tag: string) => (p: OoxmlParts) =>
  inParts(p, drawings, new RegExp(`<xdr:${tag}\\b(?:(?!</xdr:${tag}>)[\\s\\S])*?<xdr:pic\\b`, 'g'));
const relTargets = (p: OoxmlParts, re: RegExp) =>
  Object.entries(p.xml).reduce(
    (sum, [n, x]) => sum + (/^xl\/worksheets\/_rels\//.test(n) ? count(x, re) : 0),
    0
  );
const linkRel = (target: string) =>
  new RegExp(
    `<Relationship\\b[^>]*\\bTarget="${target}[^"]*"[^>]*TargetMode="External"|<Relationship\\b[^>]*TargetMode="External"[^>]*\\bTarget="${target}`,
    'gi'
  );
const nameBody = /<definedName\b([^>]*)>([^<]*)<\/definedName>/g;

export const CONSTRUCTS: XlsxConstruct[] = [
  {
    id: 'cells.numbers',
    label: 'Numeric cells',
    count: (p) => inSheets(p, /<c\b(?![^>]*\bt=")[^>]*>\s*<v>/g),
  },
  {
    id: 'cells.sharedStrings',
    label: 'Shared-string cells',
    count: (p) => inSheets(p, /<c\b[^>]*\bt="s"/g),
  },
  {
    id: 'cells.inlineStrings',
    label: 'Inline-string cells',
    count: (p) => inSheets(p, /<c\b[^>]*\bt="inlineStr"/g),
  },
  { id: 'cells.booleans', label: 'Boolean cells', count: (p) => inSheets(p, /<c\b[^>]*\bt="b"/g) },
  { id: 'cells.errors', label: 'Error cells', count: (p) => inSheets(p, /<c\b[^>]*\bt="e"/g) },
  { id: 'cells.dates', label: 'Date / time formatted cells', count: dateCells },
  {
    id: 'cells.unicode',
    label: 'Non-ASCII text',
    count: (p) => count(textParts(p), /[^\x00-\x7f]/g),
  },
  {
    id: 'cells.rtl',
    label: 'Right-to-left text (Hebrew, Arabic)',
    count: (p) => count(textParts(p), /[֐-ࣿ]/g),
  },
  {
    id: 'cells.richText',
    label: 'Rich-text strings (runs)',
    count: (p) =>
      count(part(p, 'xl/sharedStrings.xml'), /<si>\s*<r>/g) + inSheets(p, /<is>\s*<r>/g),
  },
  {
    id: 'workbook.date1904',
    label: '1904 date system',
    count: (p) => count(workbook(p), /<workbookPr\b[^>]*\bdate1904="(?:1|true)"/g),
  },
  {
    id: 'workbook.calcSettings',
    label: 'Calculation settings (calcPr)',
    count: (p) => count(workbook(p), /<calcPr\b/g),
  },
  {
    id: 'workbook.externalLinks',
    label: 'External workbook links',
    count: (p) => countNames(p, /^xl\/externalLinks\/externalLink\d+\.xml$/),
  },
  {
    id: 'styles.basic',
    label: 'Cell formats beyond the default',
    count: (p) => Math.max(0, count(block(styles(p), 'cellXfs'), /<xf\b/g) - 1),
  },
  {
    id: 'styles.alignment',
    label: 'Alignment (wrap, rotation, indent)',
    count: (p) => count(styles(p), /<alignment\b/g),
  },
  {
    id: 'styles.customNumFmts',
    label: 'Custom number formats',
    count: (p) => count(styles(p), /<numFmt\b/g),
  },
  {
    id: 'styles.themeColors',
    label: 'Theme colours',
    count: (p) => count(styles(p), /<(?:color|fgColor|bgColor)\b[^>]*\btheme="/g),
  },
  {
    id: 'styles.indexedColors',
    label: 'Indexed colours (other than the default pen 64)',
    count: (p) => count(styles(p), /<(?:color|fgColor|bgColor)\b[^>]*\bindexed="(?!64")/g),
  },
  { id: 'styles.tint', label: 'Colour tints', count: (p) => count(styles(p), /\btint="/g) },
  {
    id: 'styles.namedStyles',
    label: 'Named cell styles',
    count: (p) => count(styles(p), /<cellStyle\b(?![^>]*\bname="Normal")/g),
  },
  {
    id: 'styles.patternFill',
    label: 'Pattern fills (other than solid)',
    count: (p) => count(styles(p), /<patternFill\b[^>]*\bpatternType="(?!none"|solid"|gray125")/g),
  },
  {
    id: 'styles.gradientFill',
    label: 'Gradient fills',
    count: (p) => count(styles(p), /<gradientFill\b/g),
  },
  {
    id: 'styles.borderStyles',
    label: 'Distinct border styles (13 exist)',
    count: (p) =>
      new Set(
        [
          ...block(styles(p), 'borders').matchAll(
            /<(?:left|right|top|bottom|diagonal)\b[^>]*\bstyle="(\w+)"/g
          ),
        ].map((m) => m[1])
      ).size,
  },
  {
    id: 'styles.diagonalBorder',
    label: 'Diagonal borders',
    count: (p) => count(styles(p), /<diagonal\b[^>]*\bstyle=/g),
  },
  { id: 'formulas.basic', label: 'Formulas', count: (p) => formulas(p).length },
  {
    id: 'formulas.cachedValues',
    label: 'Formulas with a cached value',
    count: (p) => inSheets(p, /<f\b[^>]*(?:\/>|>[^<]*<\/f>)\s*<v>/g),
  },
  {
    id: 'formulas.shared',
    label: 'Shared formulas',
    count: (p) => formulas(p).filter((f) => /\bt="shared"/.test(f.attrs)).length,
  },
  {
    id: 'formulas.array',
    label: 'Array formulas (t="array")',
    count: (p) => formulas(p).filter((f) => /\bt="array"/.test(f.attrs)).length,
  },
  {
    id: 'formulas.dynamicArray',
    label: 'Dynamic-array formulas (cm metadata)',
    count: (p) =>
      formulas(p).filter((f) => /\bcm="\d+"/.test(f.cellAttrs) && /\bt="array"/.test(f.attrs))
        .length,
  },
  {
    id: 'formulas.xlfn',
    label: 'Future functions (_xlfn. prefix)',
    count: (p) => withText(p, /_xlfn\./),
  },
  {
    id: 'formulas.lambda',
    label: 'LET / LAMBDA parameters (_xlpm.)',
    count: (p) => withText(p, /_xlpm\./),
  },
  {
    id: 'formulas.wholeColumn',
    label: 'Whole-column references (A:A)',
    count: (p) => withText(p, /(?<![A-Za-z0-9_$.])\$?[A-Z]{1,3}:\$?[A-Z]{1,3}(?![A-Za-z0-9_(])/),
  },
  {
    id: 'formulas.wholeRow',
    label: 'Whole-row references (1:1)',
    count: (p) => withText(p, /(?<![A-Za-z0-9_$.])\$?\d+:\$?\d+(?![A-Za-z0-9_(])/),
  },
  { id: 'formulas.crossSheet', label: 'Cross-sheet references', count: (p) => withText(p, /!/) },
  {
    id: 'formulas.threeD',
    label: '3D references (Sheet1:Sheet3!A1)',
    count: (p) => withText(p, /'[^']+:[^']+'!|[A-Za-z0-9_.]+:[A-Za-z0-9_.]+!/),
  },
  {
    id: 'formulas.structuredRef',
    label: 'Structured table references',
    count: (p) => withText(p, /[A-Za-z_]\w*\[/),
  },
  {
    id: 'formulas.externalRef',
    label: 'External workbook references ([1]Sheet!A1)',
    count: (p) => withText(p, /\[\d+\]/),
  },
  {
    id: 'formulas.circular',
    label: 'Self-referencing formulas',
    count: (p) =>
      formulas(p).filter(
        (f) =>
          f.cell !== '' &&
          new RegExp(`(?<![A-Z$])\\$?${f.cell.replace(/(\d+)/, '\\$?$1')}(?!\\d)`).test(f.text)
      ).length,
  },
  {
    id: 'names.defined',
    label: 'Defined names',
    count: (p) => count(workbook(p), /<definedName\b/g),
  },
  {
    id: 'names.sheetScoped',
    label: 'Sheet-scoped names',
    count: (p) => count(workbook(p), /<definedName\b[^>]*\blocalSheetId=/g),
  },
  {
    id: 'names.printArea',
    label: 'Print areas',
    count: (p) => count(workbook(p), /name="_xlnm\.Print_Area"/g),
  },
  {
    id: 'names.printTitles',
    label: 'Print titles',
    count: (p) => count(workbook(p), /name="_xlnm\.Print_Titles"/g),
  },
  {
    id: 'names.constants',
    label: 'Names holding constants',
    count: (p) =>
      [...workbook(p).matchAll(nameBody)].filter((m) => /^(?:[-+]?\d|&quot;|"|\{)/.test(m[2]))
        .length,
  },
  {
    id: 'layout.colWidths',
    label: 'Custom column widths',
    count: (p) => inSheets(p, /<col\b[^>]*\bcustomWidth="1"/g),
  },
  {
    id: 'layout.rowHeights',
    label: 'Custom row heights',
    count: (p) => inSheets(p, /<row\b[^>]*\bcustomHeight="1"/g),
  },
  {
    id: 'layout.defaultSizes',
    label: 'Default column / row sizes',
    count: (p) =>
      inSheets(p, /<sheetFormatPr\b[^>]*\b(?:defaultColWidth|defaultRowHeight|baseColWidth)=/g),
  },
  {
    id: 'layout.hidden',
    label: 'Hidden rows and columns',
    count: (p) => inSheets(p, /<(?:col|row)\b[^>]*\bhidden="1"/g),
  },
  {
    id: 'layout.outline',
    label: 'Outline groups',
    count: (p) => inSheets(p, /<(?:col|row)\b[^>]*\boutlineLevel="[1-9]/g),
  },
  {
    id: 'layout.columnStyles',
    label: 'Column default styles',
    count: (p) => inSheets(p, /<col\b[^>]*\bstyle="[1-9]/g),
  },
  {
    id: 'layout.frozenPanes',
    label: 'Frozen panes',
    count: (p) => inSheets(p, /<pane\b[^>]*\bstate="frozen/g),
  },
  { id: 'layout.merges', label: 'Merged cells', count: (p) => inSheets(p, /<mergeCell\b/g) },
  {
    id: 'layout.zoom',
    label: 'Zoom other than 100%',
    count: (p) => inSheets(p, /<sheetView\b[^>]*\bzoomScale="(?!100")/g),
  },
  {
    id: 'layout.noGridlines',
    label: 'Gridlines off',
    count: (p) => inSheets(p, /<sheetView\b[^>]*\bshowGridLines="(?:0|false)"/g),
  },
  {
    id: 'layout.rtl',
    label: 'Right-to-left sheet',
    count: (p) => inSheets(p, /<sheetView\b[^>]*\brightToLeft="(?:1|true)"/g),
  },
  {
    id: 'sheets.multiple',
    label: 'More than one sheet',
    count: (p) => (count(workbook(p), /<sheet\b/g) > 1 ? count(workbook(p), /<sheet\b/g) : 0),
  },
  {
    id: 'sheets.hidden',
    label: 'Hidden sheets',
    count: (p) => count(workbook(p), /<sheet\b[^>]*\bstate="hidden"/g),
  },
  {
    id: 'sheets.veryHidden',
    label: 'Very hidden sheets',
    count: (p) => count(workbook(p), /<sheet\b[^>]*\bstate="veryHidden"/g),
  },
  { id: 'sheets.tabColor', label: 'Tab colours', count: (p) => inSheets(p, /<tabColor\b/g) },
  {
    id: 'sheets.activeTab',
    label: 'Active tab other than the first',
    count: (p) => count(workbook(p), /<workbookView\b[^>]*\bactiveTab="[1-9]/g),
  },
  { id: 'cf.cellIs', label: 'Conditional formatting: cell value', count: cf('cellIs') },
  { id: 'cf.expression', label: 'Conditional formatting: formula', count: cf('expression') },
  { id: 'cf.top10', label: 'Conditional formatting: top / bottom', count: cf('top10') },
  {
    id: 'cf.aboveAverage',
    label: 'Conditional formatting: above / below average',
    count: cf('aboveAverage'),
  },
  {
    id: 'cf.duplicates',
    label: 'Conditional formatting: duplicate / unique',
    count: cf('duplicateValues|uniqueValues'),
  },
  {
    id: 'cf.text',
    label: 'Conditional formatting: text rules',
    count: cf(
      'containsText|notContainsText|beginsWith|endsWith|containsBlanks|notContainsBlanks|containsErrors|notContainsErrors'
    ),
  },
  { id: 'cf.dates', label: 'Conditional formatting: dates occurring', count: cf('timePeriod') },
  { id: 'cf.colorScale', label: 'Conditional formatting: colour scales', count: cf('colorScale') },
  { id: 'cf.dataBar', label: 'Conditional formatting: data bars', count: cf('dataBar') },
  { id: 'cf.iconSet', label: 'Conditional formatting: icon sets', count: cf('iconSet') },
  {
    id: 'cf.stopIfTrue',
    label: 'Conditional formatting: stop if true',
    count: (p) => inSheets(p, /<(?:x14:)?cfRule\b[^>]*\bstopIfTrue="1"/g),
  },
  {
    id: 'cf.x14',
    label: 'Conditional formatting: x14 extensions',
    count: (p) => inSheets(p, /<x14:conditionalFormatting\b/g),
  },
  { id: 'dv.list', label: 'Validation: list', count: dv('list') },
  { id: 'dv.whole', label: 'Validation: whole number', count: dv('whole') },
  { id: 'dv.decimal', label: 'Validation: decimal', count: dv('decimal') },
  { id: 'dv.date', label: 'Validation: date', count: dv('date') },
  { id: 'dv.time', label: 'Validation: time', count: dv('time') },
  { id: 'dv.textLength', label: 'Validation: text length', count: dv('textLength') },
  { id: 'dv.custom', label: 'Validation: custom formula', count: dv('custom') },
  {
    id: 'dv.messages',
    label: 'Validation: input / error messages',
    count: (p) => inSheets(p, /<dataValidation\b[^>]*\b(?:errorTitle|promptTitle|error|prompt)="/g),
  },
  {
    id: 'comments.notes',
    label: 'Comments (notes)',
    count: (p) => inParts(p, /^xl\/comments(?:\/[^/]+|\d*)\.xml$/, /<comment\b/g),
  },
  {
    id: 'comments.threaded',
    label: 'Threaded comments',
    count: (p) => inParts(p, /^xl\/threadedComments\/[^/]+\.xml$/, /<threadedComment\b/g),
  },
  {
    id: 'links.external',
    label: 'External hyperlinks (http, https)',
    count: (p) => relTargets(p, linkRel('https?:')),
  },
  {
    id: 'links.mailto',
    label: 'mailto hyperlinks',
    count: (p) => relTargets(p, linkRel('mailto:')),
  },
  {
    id: 'links.internal',
    label: 'Internal hyperlinks (location)',
    count: (p) => inSheets(p, /<hyperlink\b[^>]*\blocation=/g),
  },
  {
    id: 'links.unsafeScheme',
    label: 'Hyperlinks with a script scheme',
    count: (p) => relTargets(p, linkRel('(?:javascript|vbscript|data|file):')),
  },
  { id: 'images.png', label: 'Images: PNG', count: media('png') },
  { id: 'images.jpeg', label: 'Images: JPEG', count: media('jpe?g') },
  { id: 'images.gif', label: 'Images: GIF', count: media('gif') },
  { id: 'images.svg', label: 'Images: SVG', count: media('svg') },
  { id: 'images.emf', label: 'Images: EMF', count: media('emf') },
  { id: 'images.oneCell', label: 'Images: one-cell anchor', count: anchor('oneCellAnchor') },
  { id: 'images.twoCell', label: 'Images: two-cell anchor', count: anchor('twoCellAnchor') },
  { id: 'images.absolute', label: 'Images: absolute anchor', count: anchor('absoluteAnchor') },
  {
    id: 'images.crop',
    label: 'Cropped images',
    count: (p) => inParts(p, drawings, /<a:srcRect\b[^>]*\b[ltrb]="[1-9]/g),
  },
  { id: 'charts.bar', label: 'Charts: bar / column', count: chart('barChart') },
  { id: 'charts.line', label: 'Charts: line', count: chart('lineChart') },
  { id: 'charts.pie', label: 'Charts: pie', count: chart('pieChart') },
  { id: 'charts.scatter', label: 'Charts: scatter', count: chart('scatterChart') },
  {
    id: 'charts.combo',
    label: 'Charts: combination (several plot types)',
    count: (p) =>
      Object.entries(p.xml).filter(
        ([n, x]) =>
          charts.test(n) &&
          new Set(
            x
              .match(/<(?:c:)?(?:bar|line|pie|scatter|area)Chart\b/g)
              ?.map((t) => t.replace(/^<(?:c:)?/, ''))
          ).size > 1
      ).length,
  },
  {
    id: 'tables.listObjects',
    label: 'Tables (list objects)',
    count: (p) => countNames(p, /^xl\/tables\/table\d+\.xml$/),
  },
  {
    id: 'tables.totals',
    label: 'Tables with a totals row',
    count: (p) => inParts(p, /^xl\/tables\/table\d+\.xml$/, /\btotalsRowCount="[1-9]/g),
  },
  { id: 'filters.autoFilter', label: 'Auto-filters', count: (p) => inSheets(p, /<autoFilter\b/g) },
  {
    id: 'filters.criteria',
    label: 'Filter criteria',
    count: (p) => inSheets(p, /<filterColumn\b/g),
  },
  {
    id: 'filters.sort',
    label: 'Sort state',
    count: (p) => inParts(p, /^xl\/(?:worksheets|tables)\/[^/]+\.xml$/, /<sortState\b/g),
  },
  {
    id: 'pivots.tables',
    label: 'Pivot tables',
    count: (p) => countNames(p, /^xl\/pivotTables\/pivotTable\d+\.xml$/),
  },
  {
    id: 'pivots.slicers',
    label: 'Slicers',
    count: (p) => countNames(p, /^xl\/slicers\/slicer\d+\.xml$/),
  },
  { id: 'sparklines', label: 'Sparklines', count: (p) => inSheets(p, /<x14:sparkline>/g) },
  {
    id: 'protection.sheet',
    label: 'Sheet protection',
    count: (p) => inSheets(p, /<sheetProtection\b/g),
  },
  {
    id: 'protection.password',
    label: 'Sheet protection with a password hash',
    count: (p) => inSheets(p, /<sheetProtection\b[^>]*\b(?:password|hashValue)="/g),
  },
  {
    id: 'protection.noPassword',
    label: 'Sheet protection without a password',
    count: (p) => inSheets(p, /<sheetProtection\b(?![^>]*\b(?:password|hashValue)=)[^>]*>/g),
  },
  {
    id: 'protection.workbook',
    label: 'Workbook structure protection',
    count: (p) => count(workbook(p), /<workbookProtection\b/g),
  },
  {
    id: 'print.pageSetup',
    label: 'Page setup (orientation, paper, scale)',
    count: (p) => inSheets(p, /<pageSetup\b[^>]*\b(?:orientation|paperSize|scale|fitToWidth)=/g),
  },
  {
    id: 'print.fitToPage',
    label: 'Fit to page',
    count: (p) => inSheets(p, /<pageSetUpPr\b[^>]*\bfitToPage="(?:1|true)"/g),
  },
  {
    id: 'print.headerFooter',
    label: 'Print headers and footers',
    count: (p) => inSheets(p, /<headerFooter\b/g),
  },
  {
    id: 'print.breaks',
    label: 'Manual page breaks',
    count: (p) => inSheets(p, /<(?:row|col)Breaks\b/g),
  },
  {
    id: 'macros.vba',
    label: 'Macro project (vbaProject.bin)',
    count: (p) => countNames(p, /^xl\/vbaProject\.bin$/),
  },
  {
    id: 'edge.truncatedZip',
    label: 'Truncated zip (starts as a zip, cannot be read)',
    count: (p) =>
      p.kind === 'unreadable' &&
      p.raw[0] === 0x50 &&
      p.raw[1] === 0x4b &&
      p.raw[2] === 0x03 &&
      p.raw[3] === 0x04
        ? 1
        : 0,
  },
  {
    id: 'edge.encrypted',
    label: 'Encrypted package (OLE with EncryptionInfo)',
    count: (p) => (p.kind === 'ole' && ole(p, 'EncryptionInfo') ? 1 : 0),
  },
  {
    id: 'edge.legacyBinary',
    label: 'Legacy .xls binary (OLE with a Workbook stream)',
    count: (p) => (p.kind === 'ole' && ole(p, 'Workbook') ? 1 : 0),
  },
  {
    id: 'edge.noWorkbook',
    label: 'Package without xl/workbook.xml',
    count: (p) =>
      p.kind === 'zip' &&
      p.names.includes('[Content_Types].xml') &&
      !p.names.includes('xl/workbook.xml')
        ? 1
        : 0,
  },
  { id: 'edge.brokenRels', label: 'Relationships that point nowhere', count: brokenRels },
  {
    id: 'edge.invalidXml',
    label: 'Part that is not well-formed XML',
    count: (p) => Object.values(p.xml).filter((x) => !isWellFormed(x)).length,
  },
  {
    id: 'edge.xmlEntities',
    label: 'XML entity declarations (DOCTYPE)',
    count: (p) => Object.values(p.xml).reduce((sum, x) => sum + count(x, /<!ENTITY\b/g), 0),
  },
  { id: 'edge.duplicateSheetNames', label: 'Duplicate sheet names', count: duplicateSheetNames },
  {
    id: 'edge.indexOutOfRange',
    label: 'Shared-string or style index out of range',
    count: indexOutOfRange,
  },
  {
    id: 'edge.emptyWorkbook',
    label: 'Workbook whose sheets hold no cells',
    count: (p) =>
      sheetXml(p).length > 0 &&
      sheetXml(p).every((x) => /<sheetData\s*\/>|<sheetData>\s*<\/sheetData>/.test(x))
        ? 1
        : 0,
  },
  {
    id: 'edge.hugeDimension',
    label: 'Sheet dimension of 10 million cells or more',
    count: (p) => (maxDimension(p).cells >= 10_000_000 ? 1 : 0),
  },
  {
    id: 'edge.zipBomb',
    label: 'Zip entry with a compression ratio of 200:1 or more (1 MB+)',
    count: (p) =>
      Object.values(p.sizes).filter(
        (s) => s.uncompressed >= 1_000_000 && s.uncompressed >= 200 * Math.max(1, s.compressed)
      ).length,
  },
  {
    id: 'perf.largeSheet',
    label: 'Sheet of 100,000 rows or more',
    count: (p) => (maxDimension(p).rows >= 100_000 ? 1 : 0),
  },
];

/** True when the OLE container holds a directory entry with this name (stored as UTF-16LE). */
function ole(p: OoxmlParts, name: string): boolean {
  const needle = Buffer.from(name, 'utf16le');
  return Buffer.from(p.raw).includes(needle);
}

function brokenRels(p: OoxmlParts): number {
  let broken = 0;
  for (const [name, x] of Object.entries(p.xml)) {
    if (!name.endsWith('.rels')) continue;
    for (const r of relationshipTargets(name, x))
      if (r.target !== null && !p.names.includes(r.target)) broken++;
  }
  const ids = new Set(
    relationshipTargets('xl/_rels/workbook.xml.rels', part(p, 'xl/_rels/workbook.xml.rels')).map(
      (r) => r.id
    )
  );
  for (const m of workbook(p).matchAll(/<sheet\b[^>]*\br:id="([^"]*)"/g))
    if (!ids.has(m[1])) broken++;
  return broken;
}

function duplicateSheetNames(p: OoxmlParts): number {
  const names = [...workbook(p).matchAll(/<sheet\b[^>]*\bname="([^"]*)"/g)].map((m) =>
    m[1].toLowerCase()
  );
  return names.length - new Set(names).size;
}

function indexOutOfRange(p: OoxmlParts): number {
  const strings = count(part(p, 'xl/sharedStrings.xml'), /<si\b/g);
  const xfs = count(block(styles(p), 'cellXfs'), /<xf\b/g);
  let bad = 0;
  for (const x of sheetXml(p)) {
    for (const m of x.matchAll(
      /<c\b([^>]*)>(?:\s*<f\b[^>]*(?:\/>|>[^<]*<\/f>))?(?:\s*<v>(\d+)<\/v>)?/g
    )) {
      if (/\bt="s"/.test(m[1]) && m[2] !== undefined && Number(m[2]) >= strings) bad++;
      const s = /\bs="(\d+)"/.exec(m[1])?.[1];
      if (s !== undefined && Number(s) >= xfs) bad++;
    }
  }
  return bad;
}

export function findConstruct(id: string): XlsxConstruct | undefined {
  return CONSTRUCTS.find((c) => c.id === id);
}
