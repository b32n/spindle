import JSZip from 'jszip';
import { createLossReport, type ImportLossItem, type ImportLossReport } from '@b32nio/spindle-shared';
import type { DetectXlsxLossOptions } from '@b32nio/spindle-shared/export/xlsx';
import { scanFormula, VOLATILE_FUNCTIONS, withoutStrings, type FormulaFacts } from './formula-scan';
import { isEncryptedOle, isKnownPart, isOleFile } from './package-kind';
import { parseStyles } from './styles-scan';
import { descendants, isMain, isOn, isTrue, kid, kids, parseXml, walk, type XmlElement } from './xml';

export type { DetectXlsxLossOptions };

const MAX_WHERE = 5;

/**
 * Safety limits (uncompressed bytes). A part over `partBytes`, or one that would take the total over `totalBytes`, is
 * not read; a worksheet's cell data is scanned up to `cellBytes` (parsing costs roughly a second per 10 MB) in
 * `chunkBytes` slices, so a huge sheet is never one tree.
 */
export const LIMITS = { partBytes: 64 * 1024 * 1024, totalBytes: 256 * 1024 * 1024, chunkBytes: 1024 * 1024, cellBytes: 48 * 1024 * 1024 };

/** Schemes a spreadsheet link has no business using: they run code, embed data or reach the local disk. */
const UNSAFE_SCHEME = /^(?:javascript|vbscript|data|file):/i;

/** Browsers ignore leading whitespace and control characters in a URL, so `\u0001 javascript:` is still a script link. */
function isUnsafeTarget(target: string): boolean {
  let start = 0;
  while (start < target.length && target.charCodeAt(start) <= 0x20) start++;
  return UNSAFE_SCHEME.test(target.slice(start));
}

interface Rel {
  id: string;
  /** Last segment of the relationship type URI: `drawing`, `comments`, `threadedComment`, `hyperlink`... */
  type: string;
  /** Resolved part name, or the raw URL for an external target. */
  target: string;
  external: boolean;
  /** The `Target` exactly as written. */
  raw: string;
}

interface SheetEntry {
  name: string;
  state: string;
  /** `worksheet`, `chartsheet`, ... from the relationship type; empty when it does not resolve. */
  kind: string;
  path?: string;
}

/** Counts for one sheet, summed into the report (`where` lists the sheets with a non-zero count). */
class Tally {
  private readonly perSheet = new Map<string, number>();
  total = 0;
  add(sheet: string, n = 1): void {
    if (n <= 0) return;
    this.total += n;
    this.perSheet.set(sheet, (this.perSheet.get(sheet) ?? 0) + n);
  }
  get where(): string[] {
    return [...this.perSheet.keys()];
  }
}

const TALLY_KEYS = [
  'cf', 'dv', 'filter', 'protect', 'outline', 'spark', 'comments', 'charts', 'images',
  'structured', 'unsupported', 'array', 'dynamic', 'external', 'volatile',
  'columnRowStyles', 'tabColor', 'zoom', 'rtl', 'gridlines', 'customViews',
  'printSetup', 'breaks', 'headerFooter',
  'dataBarExt', 'iconSetExt', 'formulaRule', 'stopIfTrue',
  'dvCustom', 'dvMessages', 'dvCrossSheet', 'commentsThreaded', 'commentsRich',
  'linksInternal', 'linksUnsafe', 'shapes', 'textBoxes', 'formControls', 'embedded',
] as const;
type Tallies = Record<(typeof TALLY_KEYS)[number], Tally>;
const newTallies = (): Tallies => Object.fromEntries(TALLY_KEYS.map((k) => [k, new Tally()])) as Tallies;

function resolveTarget(base: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const parts = base.split('/').slice(0, -1);
  for (const seg of target.split('/')) {
    if (seg === '..') parts.pop();
    else if (seg !== '.' && seg !== '') parts.push(seg);
  }
  return parts.join('/');
}

const relsPart = (part: string): string => {
  const slash = part.lastIndexOf('/');
  return `${part.slice(0, slash + 1)}_rels/${part.slice(slash + 1)}.rels`;
};

/** The part of JSZip's (untyped) streaming helper used here. */
interface Stream {
  on(event: 'data', fn: (chunk: Uint8Array) => void): Stream;
  on(event: 'error' | 'end', fn: () => void): Stream;
  pause(): Stream;
  resume(): Stream;
}

/**
 * Inflate a zip entry, giving up (`'limit'`) as soon as more than `limit` bytes have come out. `value` is `undefined` when it
 * cannot be read; `used` is how many bytes were inflated either way, so the caller can charge them to the total budget.
 */
function inflate(file: JSZip.JSZipObject, limit: number): Promise<{ value: Uint8Array | 'limit' | undefined; used: number }> {
  return new Promise((resolve) => {
    const chunks: Uint8Array[] = [];
    let total = 0;
    let done = false;
    const finish = (value: Uint8Array | 'limit' | undefined) => {
      if (done) return;
      done = true;
      resolve({ value, used: total });
    };
    const stream = (file as unknown as { internalStream(type: 'uint8array'): Stream }).internalStream('uint8array');
    stream
      .on('data', (chunk: Uint8Array) => {
        total += chunk.length;
        if (total > limit) {
          stream.pause();
          finish('limit');
        } else chunks.push(chunk);
      })
      .on('error', () => finish(undefined))
      .on('end', () => {
        const out = new Uint8Array(total);
        let at = 0;
        for (const c of chunks) {
          out.set(c, at);
          at += c.length;
        }
        finish(out);
      })
      .resume();
  });
}

class Package {
  read = 0;
  readonly skipped: string[] = [];
  constructor(private readonly zip: JSZip) {}

  has(name: string): boolean {
    return this.zip.file(name) !== null;
  }

  /**
   * The text of a part, or `undefined` when it is absent, over the limits or unreadable. Never throws. The declared size
   * is a cheap first check; the inflated bytes are counted as they come out and the stream is stopped at the limit, so a
   * forged header cannot make a bomb inflate.
   */
  async text(name: string): Promise<string | undefined> {
    const file = this.zip.file(name);
    if (!file) return undefined;
    const size = (file as unknown as { _data?: { uncompressedSize?: number } })._data?.uncompressedSize;
    const limit = Math.min(LIMITS.partBytes, LIMITS.totalBytes - this.read);
    if (typeof size === 'number' && size > limit) {
      this.skipped.push(name);
      return undefined;
    }
    const { value: bytes, used } = await inflate(file, limit);
    // charge every inflated byte, also those of a part that hit the limit: forged sizes on many parts must exhaust the budget
    this.read += used;
    if (bytes === 'limit') this.skipped.push(name);
    if (typeof bytes === 'string' || bytes === undefined) return undefined;
    return new TextDecoder().decode(bytes);
  }

  private readonly parsed = new Map<string, XmlElement | undefined>();

  /** A parsed part; each part is read, counted against the budget and parsed once however many sheets refer to it. */
  async xml(name: string): Promise<XmlElement | undefined> {
    if (this.parsed.has(name)) return this.parsed.get(name);
    const text = await this.text(name);
    const root = text === undefined ? undefined : (parseXml(text) ?? undefined);
    this.parsed.set(name, root);
    return root;
  }

  async rels(part: string): Promise<Rel[]> {
    const root = await this.xml(relsPart(part));
    return kids(root, 'Relationship').map((r) => {
      const external = r.attrs.TargetMode === 'External';
      const target = r.attrs.Target ?? '';
      return {
        id: r.attrs.Id ?? '',
        type: (r.attrs.Type ?? '').slice((r.attrs.Type ?? '').lastIndexOf('/') + 1),
        target: external ? target : resolveTarget(part, target),
        external,
        raw: target,
      };
    });
  }
}

/** What a worksheet's cell data holds that matters here. Rows are parsed in chunks so a large sheet is never one tree. */
interface CellFacts {
  structuredRefs: number;
  arrays: number;
  dynamicArrays: number;
  externalRefs: number;
  volatile: number;
  unsupported: number;
  unsupportedNames: string[];
  outlineRows: number;
  styledRows: number;
  styleIds: Set<number>;
  richInline: number;
  /** Facts of each shared formula's master by `si`, applied to its followers (cells with no formula text). */
  shared: Map<string, FormulaFacts>;
}

async function scanSheetData(xml: string, known: ReadonlySet<string> | undefined): Promise<{ outer: string; cells: CellFacts; limited: boolean }> {
  const cells: CellFacts = {
    structuredRefs: 0,
    arrays: 0,
    dynamicArrays: 0,
    externalRefs: 0,
    volatile: 0,
    unsupported: 0,
    unsupportedNames: [],
    outlineRows: 0,
    styledRows: 0,
    styleIds: new Set(),
    richInline: 0,
    shared: new Map(),
  };
  const open = /<((?:[A-Za-z_][\w.-]*:)?)sheetData\b[^>]*?(\/?)>/.exec(xml);
  if (!open) return { outer: xml, cells, limited: false };
  const prefix = open[1];
  const bodyStart = open.index + open[0].length;
  const closeTag = `</${prefix}sheetData>`;
  let bodyEnd = open[2] ? bodyStart : xml.indexOf(closeTag, bodyStart);
  let after = bodyEnd + closeTag.length;
  if (bodyEnd < 0) {
    // truncated: no closing tag, so everything after the opening tag is cell data
    bodyEnd = xml.length;
    after = xml.length;
  }
  if (open[2]) after = bodyStart;
  const outer = xml.slice(0, open.index) + xml.slice(after);

  const root = /<(?:[A-Za-z_][\w.-]*:)?worksheet\b[^>]*>/.exec(xml);
  const head = root ? root[0] : '<worksheet>';
  const tail = root ? `</${root[0].slice(1).split(/[\s>]/)[0]}>` : '</worksheet>';
  const rowTag = `<${prefix}row`;
  let limited = false;
  for (let pos = bodyStart; pos < bodyEnd; ) {
    if (pos - bodyStart >= LIMITS.cellBytes) {
      limited = true;
      break;
    }
    let end = Math.min(pos + LIMITS.chunkBytes, bodyEnd);
    if (end < bodyEnd) {
      const next = xml.indexOf(rowTag, end);
      end = next >= 0 && next < bodyEnd ? next : bodyEnd;
    }
    const slice = xml.slice(pos, end);
    pos = end;
    // cells with no formula, format, outline or rich text have nothing to report
    if (!/<(?:[A-Za-z_][\w.-]*:)?f[\s>/]| s=["']|outlineLevel|<(?:[A-Za-z_][\w.-]*:)?rPr/.test(slice)) continue;
    const chunk = parseXml(head + slice + tail) ?? undefined;
    for (const row of kids(chunk, 'row')) {
      const level = Number(row.attrs.outlineLevel);
      if (level >= 1) cells.outlineRows++;
      const rowStyle = Number(row.attrs.s);
      if (rowStyle >= 1 && isTrue(row.attrs.customFormat)) {
        cells.styledRows++;
        cells.styleIds.add(rowStyle);
      }
      for (const c of kids(row, 'c')) readCell(c, cells, known);
    }
  }
  return { outer, cells, limited };
}

function readCell(c: XmlElement, cells: CellFacts, known: ReadonlySet<string> | undefined): void {
  const style = Number(c.attrs.s);
  if (style >= 0) cells.styleIds.add(style);
  if (kid(kid(c, 'is'), 'r')?.children.some((e) => e.name === 'rPr')) cells.richInline++;
  const f = kid(c, 'f');
  if (!f) return;
  if (f.attrs.t === 'array') {
    if (c.attrs.cm !== undefined) cells.dynamicArrays++;
    else cells.arrays++;
  }
  // a shared formula's followers have no text of their own: they count with the facts of their master
  let facts: FormulaFacts | undefined;
  if (f.text !== '') {
    facts = scanFormula(f.text);
    if (f.attrs.t === 'shared' && f.attrs.si !== undefined) cells.shared.set(f.attrs.si, facts);
  } else if (f.attrs.t === 'shared') facts = cells.shared.get(f.attrs.si ?? '');
  if (!facts) return;
  if (facts.structuredRef) cells.structuredRefs++;
  if (facts.externalRef) cells.externalRefs++;
  if (facts.functions.some((fn) => VOLATILE_FUNCTIONS.has(fn))) cells.volatile++;
  if (known) {
    const missing = facts.functions.filter((fn) => !known.has(fn));
    if (missing.length > 0) {
      cells.unsupported++;
      for (const fn of missing) if (!cells.unsupportedNames.includes(fn)) cells.unsupportedNames.push(fn);
    }
  }
}

/** True when a page-setup, print-options or margins element asks for something other than Excel's defaults. */
function customPrintSetup(root: XmlElement): boolean {
  const setup = kid(root, 'pageSetup');
  const a = setup?.attrs ?? {};
  const fit = isTrue(kid(kid(root, 'sheetPr'), 'pageSetUpPr')?.attrs.fitToPage);
  const options = kid(root, 'printOptions')?.attrs ?? {};
  const margins = kid(root, 'pageMargins')?.attrs ?? {};
  const defaults: Record<string, number> = { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 };
  return (
    fit ||
    a.orientation === 'landscape' ||
    (a.scale !== undefined && Number(a.scale) !== 100) ||
    (a.paperSize !== undefined && a.paperSize !== '1' && a.paperSize !== '9') ||
    isTrue(a.useFirstPageNumber) ||
    Object.values(options).some(isTrue) ||
    Object.entries(defaults).some(([k, v]) => margins[k] !== undefined && Math.abs(Number(margins[k]) - v) > 1e-9)
  );
}

const hasText = (e: XmlElement | undefined): boolean => !!e && e.text.trim() !== '';

export async function detectXlsxLoss(bytes: Uint8Array, options: DetectXlsxLossOptions = {}): Promise<ImportLossReport> {
  const items: ImportLossItem[] = [];
  try {
    await detect(bytes, options, items);
  } catch {
    // never throw: report whatever was detected before the failure
  }
  return createLossReport('xlsx', items);
}

async function detect(bytes: Uint8Array, options: DetectXlsxLossOptions, items: ImportLossItem[]): Promise<void> {
  const full = options.detail === 'full';
  const add = (id: string, label: string, count: number, effect: string, where?: string[], onlyFull = false) => {
    if (count > 0 && (full || !onlyFull))
      items.push({ id, label, count, effect, ...(where && where.length ? { where: where.slice(0, MAX_WHERE) } : {}) });
  };
  /** An id that only the `full` level reports. */
  const addFull = (id: string, label: string, count: number, effect: string, where?: string[]) =>
    add(id, label, count, effect, where, true);

  if (isOleFile(bytes)) {
    if (isEncryptedOle(bytes))
      addFull('xlsx.package.encrypted', 'password-protected workbook (encrypted package)', 1, 'lost: cannot be opened; remove the password and save the file again');
    else addFull('xlsx.package.legacyBinary', 'legacy binary workbook (.xls)', 1, 'lost: not supported; save it as .xlsx first');
    return;
  }
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(bytes);
  } catch {
    return;
  }
  const pkg = new Package(zip);
  const names = Object.keys(zip.files).filter((n) => !n.endsWith('/'));
  if (pkg.has('xl/workbook.bin')) addFull('xlsx.package.legacyBinary', 'binary workbook (.xlsb)', 1, 'lost: not supported; save it as .xlsx first');

  const workbook = await pkg.xml('xl/workbook.xml');
  const sheets = await readSheets(pkg, workbook);
  // ---- per-sheet scan (also tells which cell formats are in use) ----------
  const t = newTallies();
  const scan = await scanSheets(pkg, sheets, options, t);
  const styles = parseStyles(await pkg.xml('xl/styles.xml'));

  // ---- workbook level ----------------------------------------------------
  const wb = scanWorkbook(workbook, sheets);

  // ---- summary ids that exist today (semantics kept, counts fixed) --------
  add('xlsx.conditionalFormatting', 'conditional formatting rule(s)', t.cf.total, 'lost: not imported', t.cf.where);
  add('xlsx.dataValidation', 'data validation rule(s) (dropdowns, limits)', t.dv.total, 'lost: not imported', t.dv.where);
  add('xlsx.autoFilter', 'filter(s)', t.filter.total, 'lost: not imported', t.filter.where);
  add('xlsx.sheetProtection', 'protected sheet(s)', t.protect.total, 'lost: protection is dropped', t.protect.where);
  add('xlsx.outline', 'grouped (outlined) row/column(s)', t.outline.total, 'approximated: groups are flattened', t.outline.where);
  add('xlsx.sparklines', 'sparkline group(s)', t.spark.total, 'lost: not imported', t.spark.where);
  add('xlsx.hiddenSheets', 'hidden sheet(s)', wb.hidden.length, 'approximated: become visible', wb.hidden);
  add('xlsx.definedNames', 'named range(s)/name(s)', wb.names.length, 'lost: not imported; formulas using them break', wb.names);
  add('xlsx.workbookProtection', 'workbook protection', wb.protected ? 1 : 0, 'lost: protection is dropped');

  // ---- parts that exist only when the feature is used ---------------------
  const count = (re: RegExp) => names.filter((n) => re.test(n)).length;
  const cellImages = await scanRichData(pkg, names);
  const mediaParts = count(/^xl\/media\//);
  const pictures = t.images.total + scan.backgroundPictures + cellImages.images;
  add('xlsx.charts', 'chart(s)', t.charts.total > 0 ? t.charts.total : count(/^xl\/charts\/chart\d*\.xml$/), 'lost: not imported', t.charts.where);
  add('xlsx.images', 'image(s)', pictures > 0 ? pictures : mediaParts, 'lost: not imported', t.images.where);
  const comments = t.comments.total > 0 ? t.comments.total : Math.max(count(/^xl\/comments[^/]*\.xml$/), count(/^xl\/threadedComments\/[^/]+\.xml$/));
  add('xlsx.comments', 'sheet(s) with comments/notes', comments, 'lost: comments are dropped', t.comments.where);
  add('xlsx.tables', 'table(s)', count(/^xl\/tables\/table\d*\.xml$/), 'approximated: imported as plain cells (no table styling or structured references)');
  add('xlsx.pivotTables', 'pivot table(s)', count(/^xl\/pivotTables\/pivotTable\d*\.xml$/), 'approximated: only the last computed values are kept');
  add('xlsx.macros', 'macro project', count(/^xl\/vbaProject\.bin$/), 'lost: macros are dropped');
  add('xlsx.externalLinks', 'external workbook link(s)', count(/^xl\/externalLinks\/externalLink\d*\.xml$/), 'lost: not imported');
  add('xlsx.slicers', 'slicer(s)', await countSlicers(pkg, names), 'lost: not imported');

  // ---- rich text inside cells --------------------------------------------
  const sst = await pkg.xml('xl/sharedStrings.xml');
  const richShared = kids(sst, 'si').filter((si) => kids(si, 'r').some((r) => kid(r, 'rPr'))).length;
  add('xlsx.richText', 'cell(s) with mixed formatting inside the text', richShared + scan.richInline, 'approximated: imported as plain text');

  // ---- new summary ids ----------------------------------------------------
  add('xlsx.formulas.structuredRefs', 'formula(s) using table (structured) references', t.structured.total, 'lost: table references are not calculated', t.structured.where);
  if (options.knownFunctions)
    add('xlsx.formulas.unsupported', 'formula(s) using functions the editor does not have', t.unsupported.total, 'lost: cannot be calculated (the cells show #NAME?)', scan.unsupportedNames);
  if (options.isNumberFormatExact) {
    const inexact = styles.numberFormats(scan.usedStyles).filter((f) => !f.code || !options.isNumberFormatExact?.(f.code));
    add('xlsx.numberFormats.approximated', 'number format(s) the editor cannot reproduce exactly', inexact.length, 'approximated: shown with the closest format the editor has; the original code may not survive a save', inexact.map((f) => f.label));
  }

  if (!full) return;

  // ---- full detail: formulas ---------------------------------------------
  addFull('xlsx.formulas.array', 'array formula(s)', t.array.total, 'approximated: imported as an ordinary formula in the first cell; the other cells keep their saved values', t.array.where);
  addFull('xlsx.formulas.dynamicArray', 'dynamic-array formula(s)', t.dynamic.total, "approximated: imported as an ordinary formula; Excel's array metadata is dropped", t.dynamic.where);
  addFull('xlsx.formulas.externalRef', 'formula(s) referencing another workbook', t.external.total, 'lost: external references are not evaluated', t.external.where);
  addFull('xlsx.formulas.volatile', 'formula(s) using volatile functions (NOW, RAND, OFFSET...)', t.volatile.total, "approximated: recalculated on import; the file's saved value is not kept", t.volatile.where);

  // ---- styles -------------------------------------------------------------
  const used = styles.usedFacts(scan.usedStyles);
  addFull('xlsx.styles.themeColors', 'cell format(s) using theme or indexed colours', used.themeColors, 'lost: the colour is dropped (only RGB colours are imported)');
  addFull('xlsx.styles.patternFill', 'cell format(s) with a pattern fill', used.patternFill, 'lost: pattern fills are dropped (only solid fills are imported)');
  addFull('xlsx.styles.gradientFill', 'cell format(s) with a gradient fill', used.gradientFill, 'lost: gradient fills are dropped');
  addFull('xlsx.styles.diagonalBorder', 'cell format(s) with a diagonal border', used.diagonalBorder, 'lost: diagonal borders are dropped');
  addFull('xlsx.styles.namedStyles', 'named cell style(s) in use', used.namedStyles.length, 'approximated: formatting is kept per cell; the style names are not', used.namedStyles);
  addFull('xlsx.styles.columnRowStyles', 'column/row default style(s)', t.columnRowStyles.total, 'lost: only cells present in the file keep their formatting', t.columnRowStyles.where);

  // ---- sheet --------------------------------------------------------------
  addFull('xlsx.sheet.tabColor', 'tab colour(s)', t.tabColor.total, 'lost: not imported', t.tabColor.where);
  addFull('xlsx.sheet.zoom', 'sheet(s) with a zoom other than 100%', t.zoom.total, 'lost: shown at 100%', t.zoom.where);
  addFull('xlsx.sheet.rtl', 'right-to-left sheet(s)', t.rtl.total, 'lost: shown left to right', t.rtl.where);
  addFull('xlsx.sheet.gridlines', 'sheet(s) with gridlines turned off', t.gridlines.total, 'lost: gridlines are shown', t.gridlines.where);
  addFull('xlsx.sheet.veryHidden', 'very hidden sheet(s)', wb.veryHidden.length, 'approximated: become visible', wb.veryHidden);
  addFull('xlsx.sheet.customViews', 'custom view(s)', t.customViews.total + wb.customViews, 'lost: not imported', t.customViews.where);

  // ---- print --------------------------------------------------------------
  addFull('xlsx.print.setup', 'sheet(s) with page setup (orientation, paper size, scale, margins, fit to page)', t.printSetup.total, 'lost: not imported', t.printSetup.where);
  addFull('xlsx.print.areas', 'print area(s)', wb.printAreas.length, 'lost: not imported', wb.printAreas);
  addFull('xlsx.print.titles', 'print title row/column(s)', wb.printTitles.length, 'lost: not imported', wb.printTitles);
  addFull('xlsx.print.breaks', 'manual page break(s)', t.breaks.total, 'lost: not imported', t.breaks.where);
  addFull('xlsx.print.headersFooters', 'sheet(s) with header or footer text', t.headerFooter.total, 'lost: not imported', t.headerFooter.where);

  // ---- conditional formatting, validation, comments, links ----------------
  addFull('xlsx.cf.dataBarExt', 'data bar(s) with Excel 2010 extensions', t.dataBarExt.total, 'lost: not imported', t.dataBarExt.where);
  addFull('xlsx.cf.iconSetExt', 'icon set(s) with Excel 2010 extensions', t.iconSetExt.total, 'lost: not imported', t.iconSetExt.where);
  addFull('xlsx.cf.formulaRule', 'formula-based conditional formatting rule(s)', t.formulaRule.total, 'lost: not imported', t.formulaRule.where);
  addFull('xlsx.cf.stopIfTrue', 'conditional formatting rule(s) with stop-if-true', t.stopIfTrue.total, 'lost: not imported', t.stopIfTrue.where);
  addFull('xlsx.dv.customFormula', 'custom-formula validation rule(s)', t.dvCustom.total, 'lost: not imported', t.dvCustom.where);
  addFull('xlsx.dv.messages', 'validation rule(s) with an input or error message', t.dvMessages.total, 'lost: not imported', t.dvMessages.where);
  addFull('xlsx.dv.crossSheet', 'validation rule(s) referring to another sheet', t.dvCrossSheet.total, 'lost: not imported', t.dvCrossSheet.where);
  addFull('xlsx.comments.threaded', 'threaded comment(s)', t.commentsThreaded.total, 'lost: not imported, replies included', t.commentsThreaded.where);
  addFull('xlsx.comments.richFormatting', 'comment(s) with formatted text', t.commentsRich.total, 'lost: comments are not imported, so their formatting is not either', t.commentsRich.where);
  addFull('xlsx.links.internal', 'internal hyperlink(s) (to a cell or a name)', t.linksInternal.total, 'lost: not imported', t.linksInternal.where);
  addFull('xlsx.links.unsafeScheme', 'hyperlink(s) with a script or file scheme (javascript:, vbscript:, data:, file:)', t.linksUnsafe.total, 'kept: the target is imported as written and is never opened or run by the library', t.linksUnsafe.where);

  // ---- drawing objects ----------------------------------------------------
  addFull('xlsx.objects.shapes', 'shape(s) and connector(s)', t.shapes.total, 'lost: not imported', t.shapes.where);
  addFull('xlsx.objects.textBoxes', 'text box(es)', t.textBoxes.total, 'lost: not imported', t.textBoxes.where);
  addFull('xlsx.objects.formControls', 'form control(s) and ActiveX control(s)', Math.max(t.formControls.total, count(/^xl\/ctrlProps\//)), 'lost: not imported', t.formControls.where);
  addFull('xlsx.objects.embeddedObjects', 'embedded object(s) (OLE)', t.embedded.total > 0 ? t.embedded.total : count(/^xl\/embeddings\//), 'lost: not imported', t.embedded.where);
  addFull('xlsx.objects.cellImages', 'picture(s) placed in cells', cellImages.images, 'lost: not imported');

  // ---- workbook -----------------------------------------------------------
  addFull('xlsx.workbook.date1904', '1904 date system', wb.date1904 ? 1 : 0, 'approximated: dates are converted to the 1900 system; the setting itself is not kept');
  addFull('xlsx.workbook.calcSettings', 'non-default calculation setting(s) (manual mode, iteration, precision as displayed, R1C1)', wb.calcSettings, 'lost: not imported');
  const connections = descendants(await pkg.xml('xl/connections.xml'), 'connection').length;
  const queryTables = count(/^xl\/queryTables\/queryTable\d*\.xml$/);
  const dataModel = count(/^xl\/model\//) > 0 ? 1 : 0;
  addFull('xlsx.workbook.externalData', 'data connection(s) and data model', Math.max(connections, queryTables) + dataModel, 'lost: not imported');
  addFull('xlsx.workbook.richData', 'rich data value(s) (linked data types)', cellImages.other, 'lost: not imported');

  // ---- package ------------------------------------------------------------
  if (workbook) {
    const unaccounted = names.filter((n) => !isKnownPart(n));
    addFull('xlsx.package.unaccountedParts', 'package part(s) the importer neither reads nor reports', unaccounted.length, 'lost: dropped without a more specific report', unaccounted);
  }
  addFull('xlsx.package.sizeLimit', 'package part(s) over the size limit', pkg.skipped.length, 'lost: not scanned, so their content is not reported', pkg.skipped);
}

async function readSheets(pkg: Package, workbook: XmlElement | undefined): Promise<SheetEntry[]> {
  const rels = new Map((await pkg.rels('xl/workbook.xml')).map((r) => [r.id, r]));
  return kids(kid(workbook, 'sheets'), 'sheet').map((s) => {
    const rel = rels.get(s.attrs.id ?? '');
    return {
      name: s.attrs.name ?? '?',
      state: s.attrs.state ?? 'visible',
      kind: rel && !rel.external ? rel.type : '',
      path: rel && !rel.external ? rel.target : undefined,
    };
  });
}

interface WorkbookFacts {
  hidden: string[];
  veryHidden: string[];
  names: string[];
  printAreas: string[];
  printTitles: string[];
  protected: boolean;
  date1904: boolean;
  calcSettings: number;
  customViews: number;
}

function scanWorkbook(workbook: XmlElement | undefined, sheets: SheetEntry[]): WorkbookFacts {
  const names: string[] = [];
  const printAreas: string[] = [];
  const printTitles: string[] = [];
  for (const n of kids(kid(workbook, 'definedNames'), 'definedName')) {
    const name = n.attrs.name ?? '';
    const scope = n.attrs.localSheetId === undefined ? undefined : sheets[Number(n.attrs.localSheetId)]?.name;
    if (name === '_xlnm.Print_Area') printAreas.push(scope ?? name);
    else if (name === '_xlnm.Print_Titles') printTitles.push(scope ?? name);
    else if (name && !name.startsWith('_xlnm.')) names.push(name);
  }
  const protection = kid(workbook, 'workbookProtection');
  const calc = kid(workbook, 'calcPr')?.attrs ?? {};
  const calcSettings = [
    calc.calcMode === 'manual' || calc.calcMode === 'autoNoTable',
    isTrue(calc.iterate),
    calc.fullPrecision !== undefined && !isOn(calc.fullPrecision),
    calc.refMode === 'R1C1',
  ].filter(Boolean).length;
  return {
    hidden: sheets.filter((s) => s.state === 'hidden' || s.state === 'veryHidden').map((s) => s.name),
    veryHidden: sheets.filter((s) => s.state === 'veryHidden').map((s) => s.name),
    names,
    printAreas,
    printTitles,
    protected: isTrue(protection?.attrs.lockStructure) || isTrue(protection?.attrs.lockWindows),
    date1904: isTrue(kid(workbook, 'workbookPr')?.attrs.date1904),
    calcSettings,
    customViews: kids(kid(workbook, 'customWorkbookViews'), 'customWorkbookView').length,
  };
}


interface SheetsScan {
  unsupportedNames: string[];
  /** Cell format ids used by any cell, row or column (the default format 0 included). */
  usedStyles: Set<number>;
  richInline: number;
  backgroundPictures: number;
}

async function scanSheets(pkg: Package, sheets: SheetEntry[], options: DetectXlsxLossOptions, t: Tallies): Promise<SheetsScan> {
  const scan: SheetsScan = { unsupportedNames: [], usedStyles: new Set([0]), richInline: 0, backgroundPictures: 0 };
  for (const sheet of sheets) {
    if (!sheet.path || (sheet.kind !== 'worksheet' && sheet.kind !== 'chartsheet')) continue;
    const name = sheet.name;
    const rels = await pkg.rels(sheet.path);
    /** Shape ids that belong to a control or an embedded object, so the drawing does not count them a second time. */
    const owned = new Set<string>();

    const text = sheet.kind === 'worksheet' ? await pkg.text(sheet.path) : undefined;
    if (text !== undefined) {
      const { outer, cells, limited } = await scanSheetData(text, options.knownFunctions);
      if (limited) pkg.skipped.push(sheet.path);
      t.structured.add(name, cells.structuredRefs);
      t.array.add(name, cells.arrays);
      t.dynamic.add(name, cells.dynamicArrays);
      t.external.add(name, cells.externalRefs);
      t.volatile.add(name, cells.volatile);
      t.unsupported.add(name, cells.unsupported);
      t.outline.add(name, cells.outlineRows);
      t.columnRowStyles.add(name, cells.styledRows);
      for (const fn of cells.unsupportedNames) if (!scan.unsupportedNames.includes(fn)) scan.unsupportedNames.push(fn);
      for (const id of cells.styleIds) scan.usedStyles.add(id);
      scan.richInline += cells.richInline;
      const root = parseXml(outer);
      if (root) {
        scanWorksheet(root, name, t, owned, scan.usedStyles);
        scan.backgroundPictures += kids(root, 'picture').length;
        scanLinks(root, rels, name, t);
      }
    }

    let hasComments = false;
    for (const rel of rels) {
      if (rel.type === 'drawing') await scanDrawing(pkg, rel.target, name, owned, t);
      else if (rel.type === 'comments') {
        hasComments = true;
        await scanComments(pkg, rel.target, name, t);
      } else if (rel.type === 'threadedComment') {
        hasComments = true;
        t.commentsThreaded.add(name, descendants(await pkg.xml(rel.target), 'threadedComment').length);
      }
    }
    if (hasComments) t.comments.add(name);
  }
  return scan;
}

/** Add the findings for one validation rule (main or x14 markup) to the tallies. */
function scanValidation(dv: XmlElement, name: string, t: Tallies, forceCrossSheet: boolean): void {
  if (dv.attrs.type === 'custom') t.dvCustom.add(name);
  if (['error', 'errorTitle', 'prompt', 'promptTitle'].some((a) => dv.attrs[a] !== undefined && dv.attrs[a] !== '')) t.dvMessages.add(name);
  const refersElsewhere = dv.children
    .filter((c) => c.name.startsWith('formula'))
    .some((c) => withoutStrings(c.text + c.children.map((g) => g.text).join('')).includes('!'));
  if (forceCrossSheet || refersElsewhere) t.dvCrossSheet.add(name);
}

function scanRule(rule: XmlElement, name: string, t: Tallies, extension: boolean): void {
  const type = rule.attrs.type;
  if (type === 'expression') t.formulaRule.add(name);
  if (isTrue(rule.attrs.stopIfTrue)) t.stopIfTrue.add(name);
  if (extension && type === 'dataBar') t.dataBarExt.add(name);
  if (extension && type === 'iconSet') t.iconSetExt.add(name);
}

function scanWorksheet(root: XmlElement, name: string, t: Tallies, owned: Set<string>, usedStyles: Set<number>): void {
  if (kid(kid(root, 'sheetPr'), 'tabColor')) t.tabColor.add(name);
  const view = kid(kid(root, 'sheetViews'), 'sheetView')?.attrs ?? {};
  if (view.zoomScale !== undefined && Number(view.zoomScale) !== 100) t.zoom.add(name);
  if (isTrue(view.rightToLeft)) t.rtl.add(name);
  if (view.showGridLines !== undefined && !isOn(view.showGridLines)) t.gridlines.add(name);
  t.customViews.add(name, kids(kid(root, 'customSheetViews'), 'customSheetView').length);

  for (const col of kids(kid(root, 'cols'), 'col')) {
    const style = Number(col.attrs.style);
    if (style >= 1) {
      t.columnRowStyles.add(name);
      usedStyles.add(style);
    }
    if (Number(col.attrs.outlineLevel) >= 1) t.outline.add(name);
  }

  for (const cf of kids(root, 'conditionalFormatting')) {
    if (!isMain(cf)) continue;
    t.cf.add(name);
    for (const rule of kids(cf, 'cfRule')) scanRule(rule, name, t, false);
  }
  for (const cf of descendants(root, 'conditionalFormatting', 'x14'))
    for (const rule of kids(cf, 'cfRule')) scanRule(rule, name, t, true);

  const validations = kid(root, 'dataValidations');
  if (validations && isMain(validations)) {
    for (const dv of kids(validations, 'dataValidation')) {
      t.dv.add(name);
      scanValidation(dv, name, t, false);
    }
  }
  for (const dv of descendants(root, 'dataValidation', 'x14')) scanValidation(dv, name, t, true);

  if (kids(root, 'autoFilter').some(isMain)) t.filter.add(name);
  if (isTrue(kid(root, 'sheetProtection')?.attrs.sheet)) t.protect.add(name);
  t.spark.add(name, descendants(root, 'sparklineGroup').length);

  if (customPrintSetup(root)) t.printSetup.add(name);
  t.breaks.add(name, descendants(kid(root, 'rowBreaks'), 'brk').length + descendants(kid(root, 'colBreaks'), 'brk').length);
  const hf = kid(root, 'headerFooter');
  if (hf && hf.children.some((c) => c.name !== 'extLst' && hasText(c))) t.headerFooter.add(name);

  for (const o of descendants(root, 'oleObject', 'main')) {
    t.embedded.add(name);
    if (o.attrs.shapeId) owned.add(o.attrs.shapeId);
  }
  for (const c of descendants(root, 'control', 'main')) {
    t.formControls.add(name);
    if (c.attrs.shapeId) owned.add(c.attrs.shapeId);
  }
}

function scanLinks(root: XmlElement, rels: Rel[], name: string, t: Tallies): void {
  for (const link of kids(kid(root, 'hyperlinks'), 'hyperlink')) {
    const rel = link.attrs.id === undefined ? undefined : rels.find((r) => r.id === link.attrs.id);
    if (rel) {
      if (isUnsafeTarget(rel.raw)) t.linksUnsafe.add(name);
    } else if (link.attrs.location !== undefined) t.linksInternal.add(name);
  }
}

async function scanDrawing(pkg: Package, part: string, name: string, owned: ReadonlySet<string>, t: Tallies): Promise<void> {
  const drawing = await pkg.xml(part);
  walk(drawing, (e) => {
    if (e.ns !== 'xdr') return;
    const id = descendants(e, 'cNvPr')[0]?.attrs.id;
    if (id !== undefined && owned.has(id)) return;
    if (e.name === 'pic') t.images.add(name);
    else if (e.name === 'cxnSp') t.shapes.add(name);
    else if (e.name === 'graphicFrame') {
      if (descendants(e, 'chart').length > 0) t.charts.add(name);
    } else if (e.name === 'sp') {
      if (isTrue(descendants(e, 'cNvSpPr')[0]?.attrs.txBox)) t.textBoxes.add(name);
      else t.shapes.add(name);
    }
  });
}

const EMPHASIS = ['b', 'i', 'u', 'strike', 'vertAlign'];
const signature = (rPr: XmlElement | undefined): string =>
  JSON.stringify((rPr?.children ?? []).map((e) => [e.name, e.attrs]));

/**
 * True when a note's text is formatted beyond what every note has. Excel and LibreOffice write a bold run for the author
 * line (`Author:` + newline) and plain run properties (font, size, colour) on the rest; neither counts.
 */
function hasFormattedRuns(comment: XmlElement): boolean {
  let runs = descendants(kid(comment, 'text'), 'r');
  const first = runs[0];
  if (first && kid(kid(first, 'rPr'), 'b') && kid(first, 't')?.text.endsWith('\n')) runs = runs.slice(1);
  const props = runs.map((r) => kid(r, 'rPr'));
  if (props.some((p) => EMPHASIS.some((e) => kid(p, e)))) return true;
  return new Set(props.map(signature)).size > 1;
}

async function scanComments(pkg: Package, part: string, name: string, t: Tallies): Promise<void> {
  const comments = descendants(await pkg.xml(part), 'comment');
  t.commentsRich.add(name, comments.filter(hasFormattedRuns).length);
}

/** Pictures placed in cells (`IMAGE`, "Place in cell") and other rich values, from the richData parts. */
async function scanRichData(pkg: Package, names: string[]): Promise<{ images: number; other: number }> {
  let images = descendants(await pkg.xml('xl/cellimages.xml'), 'cellImage').length;
  let other = 0;
  if (names.some((n) => n.startsWith('xl/richData/'))) {
    const structures = descendants(await pkg.xml('xl/richData/rdrichvaluestructure.xml'), 's').map((s) => s.attrs.t);
    for (const rv of descendants(await pkg.xml('xl/richData/rdrichvalue.xml'), 'rv')) {
      if (structures[Number(rv.attrs.s)] === '_localImage') images++;
      else other++;
    }
    // parts that exist but hold no readable value still mean rich data
    if (images + other === 0) other = 1;
  }
  return { images, other };
}

async function countSlicers(pkg: Package, names: string[]): Promise<number> {
  let n = 0;
  for (const part of names.filter((p) => /^xl\/slicers\/slicer\d*\.xml$/.test(p)))
    n += Math.max(1, descendants(await pkg.xml(part), 'slicer').length);
  return n;
}
