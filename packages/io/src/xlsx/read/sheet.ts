import type {
  SpreadsheetCell,
  SpreadsheetCellValue,
  SpreadsheetDefaultFormat,
  SpreadsheetMergedRange,
  SpreadsheetSheet,
  XlsxColumnInfo,
  XlsxFormulaInfo,
  XlsxHyperlink,
  XlsxRowInfo,
  XlsxSharedString,
  XlsxSheetDetails,
  XlsxSheetFormat,
  XlsxSheetProperties,
  XlsxSheetRef,
  XlsxSheetView,
  XlsxStyles,
} from '@b32nio/spindle-shared/export/xlsx';
import { isOn, isTrue, kid, kids, type XmlElement } from '../xml';
import { abbreviate, CappedWarnings, flat } from './capped';
import { LookCache, type CellLook } from './cell-style';
import { countCover, indexCells, lastCover, sweepCover } from './cell-ranges';
import { MAX_COL, MAX_ROW, parseCellRef, parseRangeRef, type RangeBounds } from './cell-ref';
import { readColor, type ColorContext } from './colors';
import { scanXml, type SaxAttrs } from './sax';
import { parseTree } from './tree';
import { decodeStringEscapes } from './shared-strings';
import { LimitLoss } from './limit-loss';
import { shiftFormula } from './shared-formula';
import { boolAttr, intAttr, numAttr, parseDecimal, parseIndex, put } from './values';
import { DEFAULT_IMPORT_LIMITS, type ImportLimits } from '../../common/safe-zip';

// ---------------------------------------------------------------------------
// One worksheet part. The part's text is held as one string (the package hands it over whole); a single forward pass (`scanXml`,
// events, no tree) walks it: `sheetData` is read cell by cell, so besides that string only the cells kept are held. `<col>`,
// `<mergeCell>` and `<hyperlink>` are read as events too, up to a count; the few elements around them that are cut out by offset
// as they close and read as small trees (`sheetPr`, `dimension`, `sheetViews`, `sheetFormatPr`) are read only up to
// `MAX_TREE_CHARS` characters each, and the raw `extLst` kept up to `MAX_EXT_CHARS`: nothing a file names is read without a bound.
//
// What a cell becomes:
//   - value by type (`t`): number (plain decimals only), shared string, formula string result, boolean, error text, inline
//     string, ISO date (validated, as the serial); `_xHHHH_` escapes are decoded in inline and formula strings as in shared
//     strings; an empty cell (no value and nothing of a formula) is not kept, as before
//   - formula: text of `<f>`; shared formulas are expanded (each cell gets the master's formula moved by its distance from the
//     master), array formulas keep their text on the anchor cell; the cached result is the value. A formula over 8192
//     characters (Excel's own limit) is dropped, and so are the shared-formula cells past a per-sheet and a per-package budget
//     of expanded text (`ImportLimits.maxSheetFormulaChars`, `maxFormulaChars`; a formula costs its length plus
//     `FORMULA_OVERHEAD`, its string's own weight): the value stays, one warning per sheet counts them. `<f/>` with no
//     text (and no `ref`/`si` meaning) is no formula: with a value the cell is a plain value, without one it is not kept
//   - position: a cell or row without a usable `r` follows the previous one silently; only a position that has to be clamped to
//     the sheet limits (past the last column or row) is warned about
//   - hyperlink: the `hyperlinks` entry for the cell, by relationship target (a `location` link is only in `details`); at most
//     `MAX_HYPERLINKS` per sheet, and a link whose relationship the sheet does not have is warned about (the first few, then one count)
//   - cells covered by a merge but its first are dropped, as before (the merge itself is `mergedRanges`); merges and hyperlinks
//     are applied by one sweep over the cells in row order (`cell-ranges.ts`), so a range of a million cells costs no more than
//     one cell and a sheet with thousands of ranges stays near-linear; at most `MAX_MERGES` merges per sheet, and merges and
//     hyperlinks together draw on the package's `maxRanges`, the rest being left out with one counted warning
//   - panes: `frozen` and `frozenSplit` (ECMA-376 18.3.1.66: frozen, but split before) both freeze rows and columns
//
// Cells past the package's cell budget (`ImportLimits.maxTotalCells`) end the scan of the part, with one warning. `<col>` elements
// are read as they come (no tree) up to `MAX_COLS` and applied by a sweep (`applyColumns`), so a file of thousands of
// full-width ranges costs no more than a few.
//
// Never throws: a damaged part gives what precedes the damage and `details.wellFormed = false`; a cell that cannot be read is a
// warning. Everything the file names (shared string index, style index, `si`) is looked up in `Map`s or arrays by number.
// ---------------------------------------------------------------------------

/** What the sheet reader needs from the workbook-level parts. */
export interface SheetContext {
  styles: XlsxStyles;
  sharedStrings: readonly XlsxSharedString[];
  colors: ColorContext;
  date1904: boolean;
  /** Cost of the formula text the sheet may gain by expanding shared formulas (`ImportLimits.maxSheetFormulaChars`). */
  maxSheetFormulaChars?: number;
  /** What the package may still spend, shared by all its sheets (`readBudget`); a sheet read alone gets the defaults. */
  budget?: ReadBudget;
  /** The cell formats' looks, shared by all the sheets of the workbook (`readFormats`); a sheet read alone makes its own. */
  formats?: LookCache;
  /** Receives non-fatal problems. */
  warn: (message: string) => void;
}

/**
 * What a package may still spend over all its worksheets: cells kept, cost of expanded shared-formula text, merged ranges and hyperlinks
 * kept, rows with settings kept, column settings, characters of raw extension lists and cell and row elements scanned. What a budget
 * cuts is counted in `loss` (the package's own, when the sheets are read through one), with the limits that name it.
 */
export interface ReadBudget {
  cells: number;
  formulaChars: number;
  ranges: number;
  rows: number;
  columns: number;
  extensionChars: number;
  scanned: number;
  /** Sheets so far that a cell or scan limit cut (the reader of a package drops a sheet that is cut before it keeps anything). */
  cuts: number;
  readonly limits: ImportLimits;
  readonly loss: LimitLoss;
}

export const readBudget = (limits: ImportLimits, loss: LimitLoss = new LimitLoss(limits)): ReadBudget => ({
  cells: limits.maxTotalCells,
  formulaChars: limits.maxFormulaChars,
  ranges: limits.maxRanges,
  rows: limits.maxRows,
  columns: limits.maxColumns,
  extensionChars: limits.maxExtensionChars,
  scanned: limits.maxScannedElements,
  cuts: 0,
  limits,
  loss,
});

/** The looks of a workbook's cell formats, for all its sheets: the first number format code over 255 characters is warned about, and each is counted as a loss. */
export function readFormats(styles: XlsxStyles, budget: ReadBudget, warn: (message: string) => void): LookCache {
  return new LookCache(styles, () => {
    if (budget.loss.add('numberFormats')) warn('A number format code is over 255 characters (Excel allows no more); it is kept as written and not read as a date or time.');
  });
}

/** What an expanded formula costs beyond its characters: the string's own weight (about 32 bytes), counted as characters of 2 bytes. */
export const FORMULA_OVERHEAD = 16;

/** Most `<col>` elements read from one sheet (the sheet has 16,384 columns, and a valid file has no overlapping ones); the rest are left out with a warning, and all sheets together draw on `maxColumns`. */
export const MAX_COLS = MAX_COL;

/** Most merged ranges read from one sheet (a sheet has 1,048,576 rows, so no sound file has more merges than that). */
export const MAX_MERGES = MAX_ROW;

/** Most hyperlinks read from one sheet: Excel's own limit (65,530 per worksheet). */
export const MAX_HYPERLINKS = 65_530;

/** Longest `sheetPr`, `dimension`, `sheetViews` or `sheetFormatPr` element (characters) that is read: real ones are a few hundred, and an element is read as a tree. A longer one is left out with a warning. */
export const MAX_TREE_CHARS = 64 * 1024;

/** Longest `extLst` (characters) kept raw in `details.extLst`: extensions of a large sheet (conditional formats, sparklines) run to hundreds of KB; a longer one is left out with a warning. */
export const MAX_EXT_CHARS = 8 * 1024 * 1024;

/** Most warnings one sheet adds for single cells; the rest are counted in one closing message. */
const MAX_CELL_WARNINGS = 20;

/** Longest formula Excel accepts (characters); a longer one in a file is not read. */
export const MAX_FORMULA_CHARS = 8192;

const ISO_DATE = /^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2})(?::(\d{2}(?:\.\d+)?))?)?Z?$/;
const EXCEL_EPOCH_UTC = Date.UTC(1899, 11, 30);
const DAY_MS = 86_400_000;

const DAYS_IN_MONTH = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];

/**
 * The serial number of an ISO 8601 date or date-time (`t="d"` cells), or `undefined` when it is not one: the date must exist
 * (`2020-02-30` does not), the time be in range, and the day not lie before 1900-01-01 (Excel has no serial for an earlier one).
 * A year under 100 is that year, not 19xx. Serials count as Excel does, including its 1900 leap-year bug (it has a 1900-02-29,
 * serial 60, that never existed): 1900-01-01 is 1, 1900-02-28 is 59 and 1900-03-01 is 61, from where a serial is the days
 * after 1899-12-30.
 */
export function isoToSerial(text: string): number | undefined {
  const m = ISO_DATE.exec(text.trim());
  if (m === null) return undefined;
  const [year, month, day, hour, minute] = [m[1], m[2], m[3], m[4] ?? '0', m[5] ?? '0'].map(Number);
  const second = Number(m[6] ?? 0);
  if (month < 1 || month > 12 || day < 1 || hour > 23 || minute > 59 || second >= 60) return undefined;
  const leap = year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0);
  if (day > DAYS_IN_MONTH[month - 1] + (month === 2 && leap ? 1 : 0)) return undefined;
  const date = new Date(0);
  date.setUTCFullYear(year, month - 1, day);
  const days = (date.getTime() - EXCEL_EPOCH_UTC) / DAY_MS;
  if (days < 2) return undefined;
  // the days before the fake 1900-02-29 sit one serial lower than the days since the epoch say
  const shift = days < 61 ? DAY_MS : 0;
  return (date.getTime() + ((hour * 60 + minute) * 60 + second) * 1000 - EXCEL_EPOCH_UTC - shift) / DAY_MS;
}

/** Days between the 1904 and the 1900 date systems. */
const DATE1904_OFFSET = 1462;

/** Width of a column in px, as the editor measures it (the inverse of the writer's Calibri-11 approximation). */
const excelColumnWidthToPx = (width: number): number => Math.round(width * 7) + 5;
const pointsToPx = (points: number): number => points / 0.75;

/** The sheet's default column width in pixels (`sheetFormatPr defaultColWidth`, read as a column's `width` is), or `undefined` when the file gives none. */
export const defaultColumnWidthPx = (format: XlsxSheetFormat): number | undefined =>
  format.defaultColWidth !== undefined && format.defaultColWidth > 0 ? excelColumnWidthToPx(format.defaultColWidth) : undefined;

function readViews(root: XmlElement | undefined): XlsxSheetView[] {
  return kids(root, 'sheetView').map((v) => {
    const a = v.attrs;
    const view: XlsxSheetView = { activeCells: kids(v, 'selection').flatMap((s) => (s.attrs.activeCell === undefined ? [] : [s.attrs.activeCell])) };
    put(view, 'showGridLines', boolAttr(a.showGridLines));
    put(view, 'showRowColHeaders', boolAttr(a.showRowColHeaders));
    put(view, 'showZeros', boolAttr(a.showZeros));
    put(view, 'rightToLeft', boolAttr(a.rightToLeft));
    put(view, 'tabSelected', boolAttr(a.tabSelected));
    put(view, 'zoomScale', numAttr(a.zoomScale));
    put(view, 'zoomScaleNormal', numAttr(a.zoomScaleNormal));
    put(view, 'view', a.view);
    put(view, 'topLeftCell', a.topLeftCell);
    const pane = kid(v, 'pane');
    if (pane) {
      view.pane = {};
      put(view.pane, 'xSplit', numAttr(pane.attrs.xSplit));
      put(view.pane, 'ySplit', numAttr(pane.attrs.ySplit));
      put(view.pane, 'topLeftCell', pane.attrs.topLeftCell);
      put(view.pane, 'activePane', pane.attrs.activePane);
      put(view.pane, 'state', pane.attrs.state);
    }
    return view;
  });
}

function readFormat(el: XmlElement | undefined): XlsxSheetFormat {
  const format: XlsxSheetFormat = {};
  if (!el) return format;
  const a = el.attrs;
  put(format, 'defaultColWidth', numAttr(a.defaultColWidth));
  put(format, 'baseColWidth', numAttr(a.baseColWidth));
  put(format, 'defaultRowHeight', numAttr(a.defaultRowHeight));
  put(format, 'customHeight', boolAttr(a.customHeight));
  put(format, 'zeroHeight', boolAttr(a.zeroHeight));
  put(format, 'outlineLevelRow', intAttr(a.outlineLevelRow));
  put(format, 'outlineLevelCol', intAttr(a.outlineLevelCol));
  return format;
}

function readProperties(el: XmlElement | undefined, colors: ColorContext): XlsxSheetProperties {
  const props: XlsxSheetProperties = {};
  if (!el) return props;
  put(props, 'codeName', el.attrs.codeName);
  put(props, 'filterMode', boolAttr(el.attrs.filterMode));
  put(props, 'tabColor', readColor(kid(el, 'tabColor'), colors));
  const outline = kid(el, 'outlinePr');
  put(props, 'summaryBelow', outline && outline.attrs.summaryBelow !== undefined ? isOn(outline.attrs.summaryBelow) : undefined);
  put(props, 'summaryRight', outline && outline.attrs.summaryRight !== undefined ? isOn(outline.attrs.summaryRight) : undefined);
  put(props, 'fitToPage', boolAttr(kid(el, 'pageSetUpPr')?.attrs.fitToPage));
  return props;
}

/** One `<col>` element, or `undefined` when its range is not one. */
function readColumn(attrs: SaxAttrs): XlsxColumnInfo | undefined {
  const min = intAttr(attrs.min);
  const max = intAttr(attrs.max);
  if (min === undefined || max === undefined || min < 1 || max < min) return undefined;
  const col: XlsxColumnInfo = { min, max };
  put(col, 'width', numAttr(attrs.width));
  put(col, 'hidden', boolAttr(attrs.hidden));
  put(col, 'customWidth', boolAttr(attrs.customWidth));
  put(col, 'bestFit', boolAttr(attrs.bestFit));
  put(col, 'style', intAttr(attrs.style));
  put(col, 'outlineLevel', intAttr(attrs.outlineLevel));
  put(col, 'collapsed', boolAttr(attrs.collapsed));
  return col;
}

/**
 * The value each column up to `maxCol` ends with for one attribute: where ranges overlap the last one that names the attribute
 * wins. The ranges are taken last to first and each column takes the first value it meets, so a column is visited once however
 * many ranges cover it (a "next column not yet taken" pointer with path compression skips the taken ones): near-linear in
 * ranges plus columns, where applying every range to every column it covers is their product.
 */
function lastWriter<T>(columns: readonly XlsxColumnInfo[], pick: (c: XlsxColumnInfo) => T | undefined, maxCol: number): Array<T | undefined> {
  const out = new Array<T | undefined>(maxCol + 1);
  const next = new Int32Array(maxCol + 2).map((_, i) => i);
  const find = (from: number): number => {
    let root = from;
    while (next[root] !== root) root = next[root];
    for (let i = from; next[i] !== root; ) {
      const after = next[i];
      next[i] = root;
      i = after;
    }
    return root;
  };
  for (let k = columns.length - 1; k >= 0; k--) {
    const c = columns[k];
    const value = pick(c);
    if (value === undefined || c.min - 1 > maxCol) continue;
    const last = Math.min(c.max - 1, maxCol);
    for (let i = find(c.min - 1); i <= last; i = find(i)) {
      out[i] = value;
      next[i] = i + 1;
    }
  }
  return out;
}

/** The editor's column widths and hidden columns from the `<col>` ranges, over the columns up to `maxCol` (in column order). */
function applyColumns(columns: readonly XlsxColumnInfo[], maxCol: number): { columnWidths: Map<number, number>; hiddenCols: Set<number> } {
  const widths = lastWriter(columns, (c) => (c.width === undefined ? undefined : excelColumnWidthToPx(c.width)), maxCol);
  const hidden = lastWriter(columns, (c) => c.hidden, maxCol);
  const columnWidths = new Map<number, number>();
  const hiddenCols = new Set<number>();
  for (let i = 0; i <= maxCol; i++) {
    const w = widths[i];
    if (w !== undefined) columnWidths.set(i, w);
    if (hidden[i]) hiddenCols.add(i);
  }
  return { columnWidths, hiddenCols };
}

/**
 * The default format (looks of a `cellXfs` entry) of each column up to `span`, from the `<col style>` ranges: the last range that names a
 * style wins (style 0 is the workbook default and clears an earlier range). Only the columns that have a format are in the result.
 */
function columnFormats(columns: readonly XlsxColumnInfo[], span: number, formats: LookCache): Map<number, SpreadsheetDefaultFormat> {
  const out = new Map<number, SpreadsheetDefaultFormat>();
  if (span <= 0) return out;
  const styles = lastWriter(columns, (c) => c.style, span - 1);
  for (let i = 0; i < span; i++) {
    const index = styles[i];
    const look = index === undefined || index < 1 ? undefined : formats.get(index);
    if (look?.style || look?.numberFormat) out.set(i, { style: look.style, numberFormat: look.numberFormat });
  }
  return out;
}

/** One `<hyperlink>` element, or `undefined` when it has no range. A missing relationship is reported through `missing`. */
function readLink(attrs: SaxAttrs, targets: ReadonlyMap<string, string>, missing: (message: string) => void): XlsxHyperlink | undefined {
  const ref = attrs.ref;
  if (ref === undefined) return undefined;
  const link: XlsxHyperlink = { ref };
  const id = attrs.id;
  if (id !== undefined) {
    const target = targets.get(id);
    if (target !== undefined) link.target = target;
    else missing(`The hyperlink at ${abbreviate(ref)} names the relationship ${abbreviate(id)}, which the sheet does not have.`);
  }
  put(link, 'location', attrs.location);
  put(link, 'display', attrs.display);
  put(link, 'tooltip', attrs.tooltip);
  return link;
}

interface RawFormula {
  text: string;
  /** The text went over `MAX_FORMULA_CHARS` (it is not kept past that point). */
  tooLong: boolean;
  type: string | undefined;
  ref: string | undefined;
  si: string | undefined;
}

interface Master {
  /** `undefined` when the master's formula was too long to keep (its cells lose theirs too). */
  formula: string | undefined;
  row: number;
  col: number;
}

/** The top-level elements of a worksheet that are cut out as small trees. */
const SMALL = new Set(['sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr']);

/** Read one worksheet part. `text` is the part's text (`undefined` when it could not be read: an empty sheet and a warning from the package). */
export function readSheet(text: string | undefined, ref: XlsxSheetRef, ctx: SheetContext): SpreadsheetSheet {
  const { styles, sharedStrings, date1904 } = ctx;
  const formulaBudget = ctx.maxSheetFormulaChars ?? DEFAULT_IMPORT_LIMITS.maxSheetFormulaChars;
  const budget = ctx.budget ?? readBudget(DEFAULT_IMPORT_LIMITS);
  const name = ref.name;
  // the file's own text (the sheet's name, a cell's address or value) is quoted in warnings shortened: a warning must not grow with it
  const label = abbreviate(name);
  const cells = new Map<string, SpreadsheetCell>();
  const formats = ctx.formats ?? readFormats(styles, budget, ctx.warn);
  const masters = new Map<string, Master>();
  const rowInfo = new Map<number, XlsxRowInfo>();
  const trees = new Map<string, XmlElement>();
  const tooBig = new Set<string>();
  const columns: XlsxColumnInfo[] = [];
  const merges: SpreadsheetMergedRange[] = [];
  const hyperlinks: XlsxHyperlink[] = [];
  // only the first `mergeCells` and the first `hyperlinks` element are read (as for the other single elements)
  const listsSeen = new Set<string>();
  let targets: Map<string, string> | undefined;
  let mergesLeftOut = 0;
  let linksLeftOut = 0;
  let rowsLeftOut = 0;
  const linkWarnings = new CappedWarnings(ctx.warn, `${label} hyperlink warnings`);
  let columnsLeftOut = 0;
  let truncated = false;
  let scanCut = false;
  let extLst: string | undefined;
  let maxRow = 0;
  let maxCol = 0;
  let cellWarnings = 0;
  const warnCell = (address: string, message: string): void => {
    if (cellWarnings++ < MAX_CELL_WARNINGS) ctx.warn(`${label}!${abbreviate(address)}: ${message}`);
  };
  let missingStrings = 0;
  let missingStyles = 0;
  let formulaChars = 0;
  let formulasTooLong = 0;
  let formulasOverBudget = 0;

  // sheetData state
  let depth = 0;
  let top = '';
  let row = -1;
  let col = -1;
  let cellAttrs: SaxAttrs | undefined;
  let value = '';
  let formula: RawFormula | undefined;
  let inline = '';
  let hasValue = false;
  let sink: 'v' | 'f' | 't' | undefined;
  let inInline = false;
  let phonetic = 0;
  let smallStart = -1;

  /** Read a small element as a tree, unless it is over `MAX_TREE_CHARS` (then it is left out, with a warning); the first of its name counts. */
  const keepTree = (tag: string, from: number, to: number): void => {
    if (trees.has(tag) || tooBig.has(tag)) return;
    if (to - from > MAX_TREE_CHARS) {
      tooBig.add(tag);
      budget.loss.add('elements', name);
      ctx.warn(`${label}: the ${tag} element is over ${MAX_TREE_CHARS} characters; it was not read.`);
      return;
    }
    // (a tree built from events: its attribute values are copies of their own, so a 13-character `dimension` does not keep the sheet's text alive)
    const { root } = parseTree((text as string).slice(from, to), MAX_TREE_CHARS);
    if (root) trees.set(tag, root);
  };

  // a shifted formula is about as long as its master (a few characters more or less where a reference gains or loses digits)
  const affordable = (cost: number): boolean => formulaChars + cost <= formulaBudget && cost <= budget.formulaChars;

  const finishCell = (): void => {
    const attrs = cellAttrs as SaxAttrs;
    cellAttrs = undefined;
    const address = attrs.r ?? `#${row + 1}:${col + 1}`;
    const type = attrs.t ?? 'n';
    const out: SpreadsheetCell = { value: null };

    let result: SpreadsheetCellValue = null;
    let sharedIndex: number | undefined;
    if (type === 's') {
      const idx = parseIndex(value);
      const item = idx === undefined ? undefined : sharedStrings[idx];
      if (item !== undefined) {
        result = item.text;
        sharedIndex = idx;
      } else if (value.trim() !== '') missingStrings++;
    } else if (type === 'inlineStr') {
      result = hasValue || inline !== '' ? flat(decodeStringEscapes(inline !== '' ? inline : value)) : null;
    } else if (type === 'str') {
      result = hasValue ? flat(decodeStringEscapes(value)) : null;
    } else if (type === 'b') {
      result = hasValue ? isTrue(value.trim()) : null;
    } else if (type === 'e') {
      result = hasValue ? flat(value) : null;
    } else if (type === 'd') {
      const serial = hasValue ? isoToSerial(value) : undefined;
      if (hasValue && serial === undefined) warnCell(address, `"${abbreviate(value)}" is not a date; imported as text`);
      result = serial ?? (hasValue ? flat(value) : null);
    } else if (hasValue && value.trim() !== '') {
      const n = parseDecimal(value);
      if (n === undefined) {
        warnCell(address, `"${abbreviate(value)}" is not a number; imported as text`);
        result = flat(value);
      } else result = n;
    }

    const styleIndex = attrs.s === undefined ? 0 : parseIndex(attrs.s);
    const look: CellLook | undefined = styleIndex === undefined ? undefined : formats.get(styleIndex);
    if (look === undefined && attrs.s !== undefined) missingStyles++;

    // the 1904 date system: dates are moved to 1900 serials, as the editor only knows that one, until it models the system
    if (date1904 && typeof result === 'number' && look?.isDate) result += DATE1904_OFFSET;

    let formulaText: string | undefined;
    let info: XlsxFormulaInfo | undefined;
    if (formula !== undefined) {
      const f = formula;
      const si = f.si === undefined ? undefined : intAttr(f.si);
      if (f.type === 'shared' && si !== undefined) {
        const key = String(si);
        if (f.tooLong) {
          masters.set(key, { formula: undefined, row, col });
          formulasTooLong++;
        } else if (f.text !== '') {
          formulaText = flat(f.text);
          masters.set(key, { formula: formulaText, row, col });
          info = { type: 'shared', si };
          if (f.ref !== undefined) info.ref = f.ref;
        } else {
          const master = masters.get(key);
          if (master === undefined) warnCell(address, 'shared formula without a resolvable expression; imported as a plain value');
          else if (master.formula === undefined) formulasOverBudget++;
          else if (!affordable(master.formula.length + FORMULA_OVERHEAD)) formulasOverBudget++;
          else {
            formulaText = shiftFormula(master.formula, row - master.row, col - master.col);
            const cost = formulaText.length + FORMULA_OVERHEAD;
            formulaChars += cost;
            budget.formulaChars -= cost;
            info = { type: 'shared', si };
          }
        }
      } else {
        if (f.tooLong) formulasTooLong++;
        else if (f.text !== '') formulaText = flat(f.text);
        if (f.type === 'array') {
          info = { type: 'array' };
          if (f.ref !== undefined) info.ref = f.ref;
          // `cm` (cell metadata) on the cell makes it a dynamic-array formula: it spills instead of being a legacy (CSE) array
          if (attrs.cm !== undefined) info.dynamic = true;
        } else if (f.type === 'dataTable') {
          info = { type: 'dataTable' };
          if (f.ref !== undefined) info.ref = f.ref;
        }
      }
    }

    // an empty cell is not kept: no value and nothing of a formula (a formula with text and no result still is, as is an array or
    // data-table anchor; `<f/>` alone is not one)
    if (result === null && formulaText === undefined && info === undefined) return;
    if (budget.cells <= 0) {
      truncated = true;
      return;
    }
    budget.cells--;
    out.value = result;
    if (formulaText !== undefined) out.formula = formulaText;
    if (look?.style) out.style = look.style;
    if (look?.numberFormat) out.numberFormat = look.numberFormat;
    if (attrs.s !== undefined && look !== undefined) out.styleIndex = styleIndex as number;
    if (sharedIndex !== undefined) out.sharedString = sharedIndex;
    if (info !== undefined) out.formulaInfo = info;
    cells.set(`${row}:${col}`, out);
    if (row > maxRow) maxRow = row;
    if (col > maxCol) maxCol = col;
  };

  const wellFormed =
    text === undefined
      ? true
      : scanXml(text, {
          open(tag, attrs, selfClosing, start, end) {
            depth++;
            if (depth === 2) {
              top = tag;
              if (tag === 'mergeCells' || tag === 'hyperlinks') {
                if (listsSeen.has(tag)) top = '';
                listsSeen.add(tag);
              }
              if (SMALL.has(tag) || tag === 'extLst') smallStart = start;
              if (selfClosing && SMALL.has(tag)) {
                keepTree(tag, start, end);
                smallStart = -1;
              }
            } else if (depth === 3 && tag === 'mergeCell' && top === 'mergeCells') {
              const bounds = attrs.ref === undefined ? undefined : parseRangeRef(attrs.ref);
              if (bounds !== undefined) {
                if (merges.length < MAX_MERGES && budget.ranges > 0) {
                  merges.push(bounds);
                  budget.ranges--;
                } else mergesLeftOut++;
              }
            } else if (depth === 3 && tag === 'hyperlink' && top === 'hyperlinks') {
              if (targets === undefined) {
                targets = new Map();
                for (const rel of ref.rels) if (!targets.has(rel.id)) targets.set(rel.id, rel.target);
              }
              if (hyperlinks.length < MAX_HYPERLINKS && budget.ranges > 0) {
                const link = readLink(attrs, targets, (m) => linkWarnings.add(m));
                if (link !== undefined) {
                  hyperlinks.push(link);
                  budget.ranges--;
                }
              } else if (attrs.ref !== undefined) linksLeftOut++;
            } else if (depth === 3 && tag === 'col' && top === 'cols') {
              if (budget.scanned-- <= 0) scanCut = true;
              const col = readColumn(attrs);
              if (col !== undefined) {
                if (columns.length < MAX_COLS && budget.columns > 0) {
                  columns.push(col);
                  budget.columns--;
                } else columnsLeftOut++;
              }
            } else if (depth === 3 && tag === 'row' && top === 'sheetData') {
              if (budget.scanned-- <= 0) scanCut = true;
              const r = intAttr(attrs.r);
              if (r !== undefined && r >= 1 && r <= MAX_ROW) row = r - 1;
              else {
                if (attrs.r !== undefined) warnCell(`row ${attrs.r}`, 'the row number is not a row of the sheet; the row follows the previous one');
                if (row + 1 >= MAX_ROW) {
                  warnCell(`row ${row + 2}`, 'the row is past the last row of the sheet; its cells are placed on the last row');
                  row = MAX_ROW - 1;
                } else row++;
              }
              col = -1;
              const info: XlsxRowInfo = { row };
              put(info, 'height', numAttr(attrs.ht));
              put(info, 'hidden', boolAttr(attrs.hidden));
              put(info, 'customHeight', boolAttr(attrs.customHeight));
              put(info, 'style', intAttr(attrs.s));
              put(info, 'customFormat', boolAttr(attrs.customFormat));
              put(info, 'outlineLevel', intAttr(attrs.outlineLevel));
              put(info, 'collapsed', boolAttr(attrs.collapsed));
              if (Object.keys(info).length > 1) {
                // a row given twice replaces its entry, so only a new row costs the budget
                if (rowInfo.has(row)) rowInfo.set(row, info);
                else if (budget.rows > 0) {
                  rowInfo.set(row, info);
                  budget.rows--;
                } else rowsLeftOut++;
              }
            } else if (depth === 4 && tag === 'c' && top === 'sheetData') {
              if (budget.scanned-- <= 0) scanCut = true;
              const pos = attrs.r === undefined ? undefined : parseCellRef(attrs.r);
              if (pos !== undefined) {
                col = pos.col;
                row = pos.row;
              } else if (col + 1 >= MAX_COL) {
                warnCell(`row ${row + 1}`, 'a cell is past the last column of the sheet; it is placed in the last column');
                col = MAX_COL - 1;
              } else col++;
              cellAttrs = attrs;
              value = '';
              inline = '';
              hasValue = false;
              formula = undefined;
              inInline = false;
              phonetic = 0;
            } else if (cellAttrs !== undefined) {
              if (depth === 5 && tag === 'v') {
                sink = 'v';
                hasValue = true;
              } else if (depth === 5 && tag === 'f') {
                sink = 'f';
                formula = { text: '', tooLong: false, type: attrs.t, ref: attrs.ref, si: attrs.si };
              } else if (depth === 5 && tag === 'is') inInline = true;
              else if (inInline && tag === 'rPh') phonetic++;
              else if (inInline && phonetic === 0 && tag === 't') sink = 't';
            }
          },
          close(tag, end) {
            if (depth === 2 && smallStart >= 0) {
              if (tag !== 'extLst') keepTree(tag, smallStart, end);
              else if (end - smallStart > MAX_EXT_CHARS) {
                budget.loss.add('extLst', name);
                ctx.warn(`${label}: the extension list (extLst) is over ${MAX_EXT_CHARS} characters; it was not kept.`);
              } else if (end - smallStart > budget.extensionChars) {
                // the package's budget: the list is not kept, and the budget is not spent on it (a smaller list in a later sheet may still fit)
                if (budget.loss.add('extLst', name)) ctx.warn(`${label}: the extension lists of the workbook are over the import limit maxExtensionChars (${budget.limits.maxExtensionChars} characters); from this sheet on those that do not fit were not kept.`);
              } else {
                budget.extensionChars -= end - smallStart;
                extLst = flat((text as string).slice(smallStart, end));
              }
              smallStart = -1;
            } else if (cellAttrs !== undefined) {
              if (depth === 4 && tag === 'c') finishCell();
              else if (tag === 'v' || tag === 'f' || tag === 't') sink = undefined;
              else if (tag === 'is') inInline = false;
              else if (tag === 'rPh') phonetic--;
            }
            depth--;
          },
          text(chunk) {
            if (sink === 'v') value += chunk;
            else if (sink === 'f') {
              const f = formula as RawFormula;
              if (!f.tooLong) {
                f.text += chunk;
                if (f.text.length > MAX_FORMULA_CHARS) {
                  f.tooLong = true;
                  f.text = '';
                }
              }
            }
            else if (sink === 't') inline += chunk;
          },
          halted: () => truncated || scanCut,
        });
  if (truncated || scanCut) budget.cuts++;
  if (truncated) {
    if (budget.loss.add('cells', name)) ctx.warn(`${label}: the workbook has more cells than the import limit allows (maxTotalCells, ${budget.limits.maxTotalCells}; raise it with the limits option); the rest of this sheet, and of the sheets after it, was not read (${cells.size} cells kept here).`);
  }
  if (scanCut) {
    if (budget.loss.add('scan', name)) ctx.warn(`${label}: the workbook has more cell, row and column elements than the import limit allows (maxScannedElements, ${budget.limits.maxScannedElements}; raise it with the limits option); the rest of this sheet, and of the sheets after it, was not read (${cells.size} cells kept here).`);
  }
  if (columnsLeftOut > 0 && budget.loss.add('columns', name, columnsLeftOut)) ctx.warn(`${label}: ${columnsLeftOut} column range(s) were left out (the limit is ${MAX_COLS} per sheet and maxColumns, ${budget.limits.maxColumns}, for the workbook); the other sheets that lose some are counted in the import report.`);
  if (!wellFormed) {
    // a cell cut off by the damage is not kept: its value may be partial
    ctx.warn(`The sheet "${label}" is not well-formed XML; it was read as far as the damage (${cells.size} cells).`);
  }
  if (cellWarnings > MAX_CELL_WARNINGS) ctx.warn(`${label}: ${cellWarnings - MAX_CELL_WARNINGS} more cell warnings were left out.`);
  if (formulasTooLong + formulasOverBudget > 0) {
    budget.loss.add('formulas', name, formulasTooLong + formulasOverBudget);
    const reasons = [
      formulasTooLong > 0 ? `${formulasTooLong} over ${MAX_FORMULA_CHARS} characters (and the cells of a shared formula with such a master)` : '',
      formulasOverBudget > 0 ? `${formulasOverBudget} past the limit of ${formulaBudget} characters of expanded shared formulas` : '',
    ].filter((r) => r !== '');
    ctx.warn(`${label}: formulas left out, their values kept: ${reasons.join('; ')}.`);
  }
  if (missingStrings > 0) ctx.warn(`${label}: ${missingStrings} cell(s) refer to a shared string the file does not have; imported as empty.`);
  if (missingStyles > 0) ctx.warn(`${label}: ${missingStyles} cell(s) use a format the file does not have; imported unformatted.`);

  linkWarnings.close();
  if (rowsLeftOut > 0) budget.loss.add('rows', name, rowsLeftOut);
  if (mergesLeftOut + linksLeftOut > 0) budget.loss.add('ranges', name, mergesLeftOut + linksLeftOut);
  if (rowsLeftOut > 0) ctx.warn(`${label}: the settings of ${rowsLeftOut} row(s) were left out (the import limit maxRows).`);
  if (mergesLeftOut > 0) ctx.warn(`${label}: ${mergesLeftOut} merged range(s) were left out (the limit is ${MAX_MERGES} per sheet and maxRanges for the workbook).`);
  if (linksLeftOut > 0) ctx.warn(`${label}: ${linksLeftOut} hyperlink(s) were left out (the limit is ${MAX_HYPERLINKS} per sheet and maxRanges for the workbook).`);
  const views = readViews(trees.get('sheetViews'));

  if (dropMergeCovered(cells, merges) > 0) {
    // the extent is that of the cells kept
    maxRow = 0;
    maxCol = 0;
    for (const key of cells.keys()) {
      const i = key.indexOf(':');
      maxRow = Math.max(maxRow, Number(key.slice(0, i)));
      maxCol = Math.max(maxCol, Number(key.slice(i + 1)));
    }
  }
  applyHyperlinks(cells, hyperlinks);

  // the columns the ranges set, up to the last cell: each draws on the package's column budget, and the ones past it are left out
  let columnCount = 0;
  if (columns.length > 0) {
    columnCount = Math.max(Math.min(maxCol + 1, budget.columns), 0);
    budget.columns -= columnCount;
    if (columnCount < maxCol + 1 && budget.loss.add('columns', name, maxCol + 1 - columnCount)) ctx.warn(`${label}: the widths and hidden state of ${maxCol + 1 - columnCount} column(s) were left out (the import limit maxColumns, ${budget.limits.maxColumns}, for the workbook); the other sheets that lose some are counted in the import report.`);
  }
  const { columnWidths, hiddenCols } = applyColumns(columns, columnCount - 1);
  // the columns that carry a default format, past the last cell too (a formatted column is mostly empty): they draw on the same budget
  let styleSpan = 0;
  for (const c of columns) if (c.style !== undefined && c.style >= 1) styleSpan = Math.max(styleSpan, Math.min(c.max, MAX_COL));
  const extra = Math.min(Math.max(styleSpan - columnCount, 0), Math.max(budget.columns, 0));
  budget.columns -= extra;
  if (columnCount + extra < styleSpan && budget.loss.add('columns', name, styleSpan - columnCount - extra)) ctx.warn(`${label}: the default formats of ${styleSpan - columnCount - extra} column(s) were left out (the import limit maxColumns, ${budget.limits.maxColumns}, for the workbook); the other sheets that lose some are counted in the import report.`);
  const columnStyles = columnFormats(columns, columnCount + extra, formats);
  const rowStyles = new Map<number, SpreadsheetDefaultFormat>();
  for (const [r, info] of rowInfo) {
    const look = info.customFormat && info.style !== undefined && info.style >= 1 ? formats.get(info.style) : undefined;
    if (look?.style || look?.numberFormat) rowStyles.set(r, { style: look.style, numberFormat: look.numberFormat });
  }
  const rowHeights = new Map<number, number>();
  const hiddenRows = new Set<number>();
  for (const [r, info] of rowInfo) {
    if (r > maxRow) continue;
    if (info.height !== undefined) rowHeights.set(r, pointsToPx(info.height));
    if (info.hidden) hiddenRows.add(r);
  }

  const frozen = views.find((v) => v.pane?.state === 'frozen' || v.pane?.state === 'frozenSplit')?.pane;
  const details: XlsxSheetDetails = {
    part: ref.part ?? '',
    state: ref.state,
    properties: readProperties(trees.get('sheetPr'), ctx.colors),
    views,
    format: readFormat(trees.get('sheetFormatPr')),
    columns,
    rows: [...rowInfo.values()],
    hyperlinks,
    wellFormed,
  };
  put(details, 'dimension', trees.get('dimension')?.attrs.ref);
  put(details, 'extLst', extLst);

  return {
    name,
    cells,
    columnWidths,
    rowHeights,
    hiddenCols,
    hiddenRows,
    columnStyles: columnStyles.size > 0 ? columnStyles : undefined,
    rowStyles: rowStyles.size > 0 ? rowStyles : undefined,
    defaultColumnWidth: defaultColumnWidthPx(details.format),
    defaultRowHeight: details.format.defaultRowHeight !== undefined ? pointsToPx(details.format.defaultRowHeight) : undefined,
    frozenRows: frozen?.ySplit || undefined,
    frozenCols: frozen?.xSplit || undefined,
    mergedRanges: merges.length > 0 ? merges : undefined,
    details,
  };
}

/** Remove the cells a merge covers but its first; returns how many were removed. */
function dropMergeCovered(cells: Map<string, SpreadsheetCell>, merges: readonly RangeBounds[]): number {
  if (merges.length === 0 || cells.size === 0) return 0;
  // a cell is covered beyond its own merges when more merges cover it than start on it
  const starts = new Map<string, number>();
  for (const m of merges) {
    const key = `${m.startRow}:${m.startCol}`;
    starts.set(key, (starts.get(key) ?? 0) + 1);
  }
  const index = indexCells(cells.keys());
  const covered: string[] = [];
  sweepCover(index, merges, countCover(), (cell, covering) => {
    if (covering > (starts.get(index.keys[cell]) ?? 0)) covered.push(index.keys[cell]);
  });
  for (const key of covered) cells.delete(key);
  return covered.length;
}

/** Give every cell inside a link's range its target; where links overlap the last one written wins. */
function applyHyperlinks(cells: Map<string, SpreadsheetCell>, links: readonly XlsxHyperlink[]): void {
  const ranges: RangeBounds[] = [];
  const targets: string[] = [];
  for (const link of links) {
    const range = link.target === undefined ? undefined : parseRangeRef(link.ref);
    if (range !== undefined) {
      ranges.push(range);
      targets.push(link.target as string);
    }
  }
  if (ranges.length === 0 || cells.size === 0) return;
  const index = indexCells(cells.keys());
  sweepCover(index, ranges, lastCover(ranges.length), (cell, last) => {
    if (last >= 0) (cells.get(index.keys[cell]) as SpreadsheetCell).hyperlink = targets[last];
  });
}
