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
import { parseCellRef, parseRangeRef, type RangeBounds } from './cell-ref';
import { readColor, type ColorContext } from './colors';
import { scanXml, type SaxAttrs } from './sax';
import { shiftFormula } from './shared-formula';
import { boolAttr, intAttr, numAttr, put } from './values';

// ---------------------------------------------------------------------------
// One worksheet part. A single forward pass (`scanXml`, no tree) walks the part: `sheetData` is read cell by cell as events,
// so memory stays proportional to the cells kept, and the small elements around it (`sheetPr`, `dimension`, `sheetViews`,
// `sheetFormatPr`, `cols`, `mergeCells`, `hyperlinks`, `extLst`) are cut out by offset as they close and read as little trees.
//
// What a cell becomes:
//   - value by type (`t`): number, shared string, formula string result, boolean, error text, inline string, ISO date (as the
//     serial); an empty cell (no value, no formula) is not kept, as before
//   - formula: text of `<f>`; shared formulas are expanded (each cell gets the master's formula moved by its distance from the
//     master), array formulas keep their text on the anchor cell; the cached result is the value
//   - hyperlink: the `hyperlinks` entry for the cell, by relationship target (a `location` link is only in `details`)
//   - cells covered by a merge but its first are dropped, as before (the merge itself is `mergedRanges`)
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
  /** Receives non-fatal problems. */
  warn: (message: string) => void;
}

/** Most warnings one sheet adds for single cells; the rest are counted in one closing message. */
const MAX_CELL_WARNINGS = 20;

const ISO_DATE = /^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2})(?::(\d{2}(?:\.\d+)?))?)?Z?$/;
const EXCEL_EPOCH_UTC = Date.UTC(1899, 11, 30);
const DAY_MS = 86_400_000;

/** The serial number of an ISO 8601 date or date-time (`t="d"` cells), or `undefined` when it is not one. */
export function isoToSerial(text: string): number | undefined {
  const m = ISO_DATE.exec(text.trim());
  if (m === null) return undefined;
  const ms = Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3]), Number(m[4] ?? 0), Number(m[5] ?? 0), 0) + Number(m[6] ?? 0) * 1000;
  return (ms - EXCEL_EPOCH_UTC) / DAY_MS;
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
  for (const el of kids(root, 'hyperlink')) {
    const ref = el.attrs.ref;
    if (ref === undefined) continue;
    const link: XlsxHyperlink = { ref };
    const id = el.attrs.id;
    if (id !== undefined) {
      const rel = rels.find((r) => r.id === id);
      if (rel) link.target = rel.target;
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
  type: string | undefined;
  ref: string | undefined;
  si: string | undefined;
}

interface Master {
  formula: string;
  row: number;
  col: number;
}

/** The top-level elements of a worksheet that are cut out as small trees. */
const SMALL = new Set(['sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr', 'cols', 'mergeCells', 'hyperlinks']);

/** Read one worksheet part. `text` is the part's text (`undefined` when it could not be read: an empty sheet and a warning from the package). */
export function readSheet(text: string | undefined, ref: XlsxSheetRef, ctx: SheetContext): SpreadsheetSheet {
  const { styles, sharedStrings, date1904 } = ctx;
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
      const idx = Number(value);
      const item = Number.isInteger(idx) && value.trim() !== '' ? sharedStrings[idx] : undefined;
      if (item !== undefined) {
        result = item.text;
        sharedIndex = idx;
      } else if (value.trim() !== '') missingStrings++;
    } else if (type === 'inlineStr') {
      result = hasValue || inline !== '' ? (inline !== '' ? inline : value) : null;
    } else if (type === 'str') {
      result = hasValue ? value : null;
    } else if (type === 'b') {
      result = hasValue ? isTrue(value.trim()) : null;
    } else if (type === 'e') {
      result = hasValue ? value : null;
    } else if (type === 'd') {
      const serial = hasValue ? isoToSerial(value) : undefined;
      if (hasValue && serial === undefined) warnCell(address, `"${value}" is not a date; imported as text`);
      result = serial ?? (hasValue ? value : null);
    } else if (hasValue && value.trim() !== '') {
      const n = Number(value);
      if (Number.isNaN(n)) {
        warnCell(address, `"${value}" is not a number; imported as text`);
        result = value;
      } else result = n;
    }

    const styleIndex = attrs.s === undefined ? 0 : Number(attrs.s);
    let look: CellLook | undefined;
    if (Number.isInteger(styleIndex) && styleIndex >= 0) {
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
        if (f.text !== '') {
          masters.set(key, { formula: f.text, row, col });
          formulaText = f.text;
          info = { type: 'shared', si };
          if (f.ref !== undefined) info.ref = f.ref;
        } else {
          const master = masters.get(key);
          if (master === undefined) warnCell(address, 'shared formula without a resolvable expression; imported as a plain value');
          else {
            formulaText = shiftFormula(master.formula, row - master.row, col - master.col);
            info = { type: 'shared', si };
          }
        }
      } else if (f.type === 'array') {
        if (f.text !== '') formulaText = f.text;
        info = { type: 'array' };
        if (f.ref !== undefined) info.ref = f.ref;
      } else if (f.type === 'dataTable') {
        info = { type: 'dataTable' };
        if (f.ref !== undefined) info.ref = f.ref;
      } else if (f.text !== '') formulaText = f.text;
    }

    // an empty cell is not kept (a formula without a result still is)
    if (result === null && formula === undefined) return;
    out.value = result;
    if (formulaText !== undefined) out.formula = formulaText;
    if (look?.style) out.style = look.style;
    if (look?.numberFormat) out.numberFormat = look.numberFormat;
    if (attrs.s !== undefined && look !== undefined) out.styleIndex = styleIndex;
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
              row = r !== undefined && r >= 1 ? r - 1 : row + 1;
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
                formula = { text: '', type: attrs.t, ref: attrs.ref, si: attrs.si };
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
            else if (sink === 'f') (formula as RawFormula).text += chunk;
            else if (sink === 't') inline += chunk;
          },
        });
  if (!wellFormed) {
    // a cell cut off by the damage is not kept: its value may be partial
    ctx.warn(`The sheet "${name}" is not well-formed XML; it was read as far as the damage (${cells.size} cells).`);
  }
  if (cellWarnings > MAX_CELL_WARNINGS) ctx.warn(`${name}: ${cellWarnings - MAX_CELL_WARNINGS} more cell warnings were left out.`);
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

  const frozen = views.find((v) => v.pane?.state === 'frozen')?.pane;
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

/** Run `visit` for every kept cell inside `range`, walking whichever is smaller: the range or the cells. */
function forCellsIn(cells: Map<string, SpreadsheetCell>, range: RangeBounds, visit: (key: string, row: number, col: number) => void): void {
  const area = (range.endRow - range.startRow + 1) * (range.endCol - range.startCol + 1);
  if (area <= cells.size) {
    for (let r = range.startRow; r <= range.endRow; r++) {
      for (let c = range.startCol; c <= range.endCol; c++) {
        const key = `${r}:${c}`;
        if (cells.has(key)) visit(key, r, c);
      }
    }
  } else {
    for (const key of [...cells.keys()]) {
      const i = key.indexOf(':');
      const r = Number(key.slice(0, i));
      const c = Number(key.slice(i + 1));
      if (r >= range.startRow && r <= range.endRow && c >= range.startCol && c <= range.endCol) visit(key, r, c);
    }
  }
}

/** Remove the cells a merge covers but its first; returns how many were removed. */
function dropMergeCovered(cells: Map<string, SpreadsheetCell>, merges: readonly RangeBounds[]): number {
  let removed = 0;
  for (const m of merges) {
    forCellsIn(cells, m, (key, r, c) => {
      if (r !== m.startRow || c !== m.startCol) {
        cells.delete(key);
        removed++;
      }
    });
  }
  return removed;
}

function applyHyperlinks(cells: Map<string, SpreadsheetCell>, links: readonly XlsxHyperlink[]): void {
  for (const link of links) {
    if (link.target === undefined) continue;
    const range = parseRangeRef(link.ref);
    if (range === undefined) continue;
    const target = link.target;
    forCellsIn(cells, range, (key) => {
      (cells.get(key) as SpreadsheetCell).hyperlink = target;
    });
  }
}
