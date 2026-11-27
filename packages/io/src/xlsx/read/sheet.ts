import type {
  SpreadsheetCell,
  SpreadsheetCellValue,
  SpreadsheetMergedRange,
  SpreadsheetSheet,
  XlsxColumnInfo,
  XlsxFormulaInfo,
  XlsxHyperlink,
  XlsxRelationship,
  XlsxRowInfo,
  XlsxSharedString,
  XlsxSheetDetails,
  XlsxSheetFormat,
  XlsxSheetProperties,
  XlsxSheetRef,
  XlsxSheetView,
  XlsxStyles,
} from '@b32nio/spindle-shared/export/xlsx';
import { isOn, isTrue, kid, kids, parseXml, type XmlElement } from '../xml';
import { lookOf, isDateFormat, type CellLook } from './cell-style';
import { countCover, indexCells, lastCover, sweepCover } from './cell-ranges';
import { MAX_COL, MAX_ROW, parseCellRef, parseRangeRef, type RangeBounds } from './cell-ref';
import { readColor, type ColorContext } from './colors';
import { scanXml, type SaxAttrs } from './sax';
import { decodeStringEscapes } from './shared-strings';
import { shiftFormula } from './shared-formula';
import { boolAttr, intAttr, numAttr, parseDecimal, parseIndex, put } from './values';
import { DEFAULT_IMPORT_LIMITS } from '../../common/safe-zip';

// ---------------------------------------------------------------------------
// One worksheet part. The part's text is held as one string (the package hands it over whole); a single forward pass (`scanXml`,
// events, no tree) walks it: `sheetData` is read cell by cell, so besides that string only the cells kept are held, and the small
// elements around it (`sheetPr`, `dimension`, `sheetViews`, `sheetFormatPr`, `cols`, `mergeCells`, `hyperlinks`, `extLst`) are cut
// out by offset as they close and read as little trees.
//
// What a cell becomes:
//   - value by type (`t`): number (plain decimals only), shared string, formula string result, boolean, error text, inline
//     string, ISO date (validated, as the serial); `_xHHHH_` escapes are decoded in inline and formula strings as in shared
//     strings; an empty cell (no value and nothing of a formula) is not kept, as before
//   - formula: text of `<f>`; shared formulas are expanded (each cell gets the master's formula moved by its distance from the
//     master), array formulas keep their text on the anchor cell; the cached result is the value. A formula over 8192
//     characters (Excel's own limit) is dropped, and so are the shared-formula cells past a per-sheet budget of expanded
//     characters (`ImportLimits.maxSheetFormulaChars`): the value stays, one warning per sheet counts them. `<f/>` with no
//     text (and no `ref`/`si` meaning) is no formula: with a value the cell is a plain value, without one it is not kept
//   - position: a cell or row without a usable `r` follows the previous one, clamped to the sheet limits (with a warning)
//   - hyperlink: the `hyperlinks` entry for the cell, by relationship target (a `location` link is only in `details`)
//   - cells covered by a merge but its first are dropped, as before (the merge itself is `mergedRanges`); merges and hyperlinks
//     are applied by one sweep over the cells in row order (`cell-ranges.ts`), so a range of a million cells costs no more than
//     one cell and a sheet with thousands of ranges stays near-linear
//   - panes: `frozen` and `frozenSplit` (ECMA-376 18.3.1.66: frozen, but split before) both freeze rows and columns
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
  /** Characters of formula text the sheet may gain by expanding shared formulas (`ImportLimits.maxSheetFormulaChars`). */
  maxFormulaChars?: number;
  /** Receives non-fatal problems. */
  warn: (message: string) => void;
}

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
 * (`2020-02-30` does not), the time be in range, and the day not lie before the epoch (1899-12-30: Excel has no serial for an
 * earlier one). A year under 100 is that year, not 19xx.
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
  const serial = (date.getTime() + ((hour * 60 + minute) * 60 + second) * 1000 - EXCEL_EPOCH_UTC) / DAY_MS;
  return serial < 0 ? undefined : serial;
}

/** Days between the 1904 and the 1900 date systems. */
const DATE1904_OFFSET = 1462;

/** Width of a column in px, as the editor measures it (the inverse of the writer's Calibri-11 approximation). */
const excelColumnWidthToPx = (width: number): number => Math.round(width * 7) + 5;
const pointsToPx = (points: number): number => points / 0.75;

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

function readColumns(root: XmlElement | undefined): XlsxColumnInfo[] {
  const out: XlsxColumnInfo[] = [];
  for (const el of kids(root, 'col')) {
    const min = intAttr(el.attrs.min);
    const max = intAttr(el.attrs.max);
    if (min === undefined || max === undefined || min < 1 || max < min) continue;
    const col: XlsxColumnInfo = { min, max };
    put(col, 'width', numAttr(el.attrs.width));
    put(col, 'hidden', boolAttr(el.attrs.hidden));
    put(col, 'customWidth', boolAttr(el.attrs.customWidth));
    put(col, 'bestFit', boolAttr(el.attrs.bestFit));
    put(col, 'style', intAttr(el.attrs.style));
    put(col, 'outlineLevel', intAttr(el.attrs.outlineLevel));
    put(col, 'collapsed', boolAttr(el.attrs.collapsed));
    out.push(col);
  }
  return out;
}

function readHyperlinks(root: XmlElement | undefined, rels: readonly XlsxRelationship[], warn: (m: string) => void): XlsxHyperlink[] {
  const out: XlsxHyperlink[] = [];
  const targets = new Map<string, string>();
  for (const rel of rels) if (!targets.has(rel.id)) targets.set(rel.id, rel.target);
  for (const el of kids(root, 'hyperlink')) {
    const ref = el.attrs.ref;
    if (ref === undefined) continue;
    const link: XlsxHyperlink = { ref };
    const id = el.attrs.id;
    if (id !== undefined) {
      const target = targets.get(id);
      if (target !== undefined) link.target = target;
      else warn(`The hyperlink at ${ref} names the relationship ${id}, which the sheet does not have.`);
    }
    put(link, 'location', el.attrs.location);
    put(link, 'display', el.attrs.display);
    put(link, 'tooltip', el.attrs.tooltip);
    out.push(link);
  }
  return out;
}

function readMerges(root: XmlElement | undefined): SpreadsheetMergedRange[] {
  const out: SpreadsheetMergedRange[] = [];
  for (const el of kids(root, 'mergeCell')) {
    const bounds = el.attrs.ref === undefined ? undefined : parseRangeRef(el.attrs.ref);
    if (bounds !== undefined) out.push(bounds);
  }
  return out;
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
const SMALL = new Set(['sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr', 'cols', 'mergeCells', 'hyperlinks']);

/** Read one worksheet part. `text` is the part's text (`undefined` when it could not be read: an empty sheet and a warning from the package). */
export function readSheet(text: string | undefined, ref: XlsxSheetRef, ctx: SheetContext): SpreadsheetSheet {
  const { styles, sharedStrings, date1904 } = ctx;
  const formulaBudget = ctx.maxFormulaChars ?? DEFAULT_IMPORT_LIMITS.maxSheetFormulaChars;
  const name = ref.name;
  const cells = new Map<string, SpreadsheetCell>();
  const looks = new Map<number, CellLook>();
  const masters = new Map<string, Master>();
  const rowInfo = new Map<number, XlsxRowInfo>();
  const trees = new Map<string, XmlElement>();
  let extLst: string | undefined;
  let maxRow = 0;
  let maxCol = 0;
  let cellWarnings = 0;
  const warnCell = (address: string, message: string): void => {
    if (cellWarnings++ < MAX_CELL_WARNINGS) ctx.warn(`${name}!${address}: ${message}`);
  };
  let missingStrings = 0;
  let missingStyles = 0;
  let formulaChars = 0;
  let formulasTooLong = 0;
  let formulasOverBudget = 0;

  // sheetData state
  let depth = 0;
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
      result = hasValue || inline !== '' ? decodeStringEscapes(inline !== '' ? inline : value) : null;
    } else if (type === 'str') {
      result = hasValue ? decodeStringEscapes(value) : null;
    } else if (type === 'b') {
      result = hasValue ? isTrue(value.trim()) : null;
    } else if (type === 'e') {
      result = hasValue ? value : null;
    } else if (type === 'd') {
      const serial = hasValue ? isoToSerial(value) : undefined;
      if (hasValue && serial === undefined) warnCell(address, `"${value}" is not a date; imported as text`);
      result = serial ?? (hasValue ? value : null);
    } else if (hasValue && value.trim() !== '') {
      const n = parseDecimal(value);
      if (n === undefined) {
        warnCell(address, `"${value}" is not a number; imported as text`);
        result = value;
      } else result = n;
    }

    const styleIndex = attrs.s === undefined ? 0 : parseIndex(attrs.s);
    let look: CellLook | undefined;
    if (styleIndex !== undefined) {
      look = looks.get(styleIndex);
      if (look === undefined) {
        const xf = styles.cellXfs[styleIndex];
        if (xf !== undefined) {
          look = lookOf(xf, styles, styleIndex === 0);
          looks.set(styleIndex, look);
        }
      }
    }
    if (look === undefined && attrs.s !== undefined) missingStyles++;

    // the 1904 date system: dates are moved to 1900 serials, as the editor only knows that one, until it models the system
    if (date1904 && typeof result === 'number' && isDateFormat(look?.numberFormat?.code)) result += DATE1904_OFFSET;

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
          masters.set(key, { formula: f.text, row, col });
          formulaText = f.text;
          info = { type: 'shared', si };
          if (f.ref !== undefined) info.ref = f.ref;
        } else {
          const master = masters.get(key);
          if (master === undefined) warnCell(address, 'shared formula without a resolvable expression; imported as a plain value');
          else if (master.formula === undefined) formulasOverBudget++;
          else if (formulaChars >= formulaBudget) formulasOverBudget++;
          else {
            formulaText = shiftFormula(master.formula, row - master.row, col - master.col);
            formulaChars += formulaText.length;
            info = { type: 'shared', si };
          }
        }
      } else {
        if (f.tooLong) formulasTooLong++;
        else if (f.text !== '') formulaText = f.text;
        if (f.type === 'array') {
          info = { type: 'array' };
          if (f.ref !== undefined) info.ref = f.ref;
        } else if (f.type === 'dataTable') {
          info = { type: 'dataTable' };
          if (f.ref !== undefined) info.ref = f.ref;
        }
      }
    }

    // an empty cell is not kept: no value and nothing of a formula (a formula with text and no result still is, as is an array or
    // data-table anchor; `<f/>` alone is not one)
    if (result === null && formulaText === undefined && info === undefined) return;
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
              if (SMALL.has(tag) || tag === 'extLst') smallStart = start;
              if (selfClosing && SMALL.has(tag)) {
                const el = trees.has(tag) ? undefined : parseXml(text.slice(start, end));
                if (el) trees.set(tag, el);
                smallStart = -1;
              }
            } else if (depth === 3 && tag === 'row') {
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
              if (Object.keys(info).length > 1) rowInfo.set(row, info);
            } else if (depth === 4 && tag === 'c') {
              const pos = attrs.r === undefined ? undefined : parseCellRef(attrs.r);
              if (pos !== undefined) {
                col = pos.col;
                if (pos.row !== row) row = pos.row;
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
              const xml = (text as string).slice(smallStart, end);
              if (tag === 'extLst') extLst = xml;
              else if (SMALL.has(tag) && !trees.has(tag)) {
                const el = parseXml(xml);
                if (el) trees.set(tag, el);
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
        });
  if (!wellFormed) {
    // a cell cut off by the damage is not kept: its value may be partial
    ctx.warn(`The sheet "${name}" is not well-formed XML; it was read as far as the damage (${cells.size} cells).`);
  }
  if (cellWarnings > MAX_CELL_WARNINGS) ctx.warn(`${name}: ${cellWarnings - MAX_CELL_WARNINGS} more cell warnings were left out.`);
  if (formulasTooLong + formulasOverBudget > 0) {
    const reasons = [
      formulasTooLong > 0 ? `${formulasTooLong} over ${MAX_FORMULA_CHARS} characters (and the cells of a shared formula with such a master)` : '',
      formulasOverBudget > 0 ? `${formulasOverBudget} past the limit of ${formulaBudget} characters of expanded shared formulas` : '',
    ].filter((r) => r !== '');
    ctx.warn(`${name}: formulas left out, their values kept: ${reasons.join('; ')}.`);
  }
  if (missingStrings > 0) ctx.warn(`${name}: ${missingStrings} cell(s) refer to a shared string the file does not have; imported as empty.`);
  if (missingStyles > 0) ctx.warn(`${name}: ${missingStyles} cell(s) use a format the file does not have; imported unformatted.`);

  const merges = readMerges(trees.get('mergeCells'));
  const hyperlinks = readHyperlinks(trees.get('hyperlinks'), ref.rels, ctx.warn);
  const columns = readColumns(trees.get('cols'));
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

  const columnWidths = new Map<number, number>();
  const hiddenCols = new Set<number>();
  for (const c of columns) {
    for (let i = c.min - 1; i <= Math.min(c.max - 1, maxCol); i++) {
      if (c.width !== undefined) columnWidths.set(i, excelColumnWidthToPx(c.width));
      if (c.hidden) hiddenCols.add(i);
      else if (c.hidden === false) hiddenCols.delete(i);
    }
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
