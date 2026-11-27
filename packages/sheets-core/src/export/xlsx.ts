import type { DetectXlsxLossOptions, ImportLossItem, ImportLossReport } from '@b32nio/spindle-shared/export/xlsx';
import type { ImportXlsxOptions } from '@b32nio/spindle-io/xlsx';

// The XLSX reader/writer lives in the separate @b32nio/spindle-io package; load it on first use so
// hosts that never open or save the format don't pay for its dependencies.
const loadIo = () => import('@b32nio/spindle-io/xlsx');
import type {
  SpreadsheetBorder,
  SpreadsheetCell,
  SpreadsheetCellStyle,
  SpreadsheetImportResult,
  SpreadsheetMergedRange,
  SpreadsheetSheet,
} from '@b32nio/spindle-shared/export/xlsx';
import type { Cell, CellStyle, SheetData, Sheet, WorkbookData } from '../types';
import type { WorkbookImpl } from '../workbook';
import { StylePool } from '../style-pool';
import { FormatPool } from '../format-pool';
import { generateId } from '../utils/id';
import { toExcelNumberFormatCode, fromExcelNumberFormatCode, normalizeExcelNumberFormatCode } from './xlsx-number-format';
import { isFaithfulFormatCode } from '../number-format';
import { getRegisteredFunctionNames } from '../formula-parser/functions';
import { isFormulaEvaluable, parseRange, spanRef } from './xlsx-formulas';

type IoXlsx = Awaited<ReturnType<typeof loadIo>>;

/** What `exportToXlsx` takes. */
export interface ExportToXlsxOptions {
  /**
   * The most cells of one array formula's result written to the file (default 1,000,000, `DEFAULT_MAX_ARRAY_CELLS`): a spilled
   * region over it is written as its first rows only, with the formula's range cut to match, so Excel (which recalculates
   * the formula) still spills it in full. Each region cut is reported to `onWarning`.
   */
  maxArrayCells?: number;
  /** Told about what the export left out (a spilled region over `maxArrayCells`), once for each. */
  onWarning?: (message: string) => void;
}

export async function exportToXlsx(workbook: WorkbookImpl, options: ExportToXlsxOptions = {}): Promise<Uint8Array> {
  const maxArrayCells = options.maxArrayCells ?? DEFAULT_MAX_ARRAY_CELLS;
  const sheetIds = Array.from(workbook.sheets.keys());
  const sheets = sheetIds.map((id) => toSpreadsheetSheet(workbook, workbook.sheets.get(id)!, maxArrayCells, options.onWarning));
  const activeSheetIndex = sheetIds.indexOf(workbook.activeSheetId);
  const { exportXlsx } = await loadIo();
  return exportXlsx({ sheets, activeSheetIndex: activeSheetIndex >= 0 ? activeSheetIndex : undefined });
}

function toSpreadsheetSheet(workbook: WorkbookImpl, sheet: Sheet, maxArrayCells: number, onWarning?: (message: string) => void): SpreadsheetSheet {
  const stylePool = workbook.getStylePool();
  const formatPool = workbook.getFormatPool();
  const cells = new Map<string, SpreadsheetCell>();

  // Bounded by actual cell data, same as CSV export — a hidden/resized
  // row or column past the last cell with data isn't carried over.
  let maxRow = 0;
  let maxCol = 0;
  for (const [row, col, cell] of sheet.entries()) {
    maxRow = Math.max(maxRow, row);
    maxCol = Math.max(maxCol, col);
    const out = toSpreadsheetCell(cell, stylePool, formatPool);
    cells.set(`${row}:${col}`, out);
    if (cell.formula) {
      // an array formula is written as one over its extent: the file's (cached-only) or the engine's live spill region
      const region = workbook.getSpillRegion(sheet.id, row, col);
      let span = region ?? cell.arraySpan;
      let written = region;
      if (region && region.rows * region.cols > maxArrayCells) {
        // a result over the cap is written as its first rows; the range says what is written, and Excel spills the rest when it recalculates
        const cols = Math.min(region.cols, Math.max(1, maxArrayCells));
        const rows = Math.min(region.rows, Math.max(1, Math.floor(maxArrayCells / cols)));
        span = { rows, cols };
        written = { ...region, rows, cols };
        onWarning?.(`${sheet.name.slice(0, 64)}!${spanRef(row, col, 1, 1)}: the array formula's result (${region.rows} x ${region.cols} cells) is over ${maxArrayCells} cells, so only its first ${rows} x ${cols} are written to the file.`);
      }
      if (span) {
        out.arrayRef = spanRef(row, col, span.rows, span.cols);
        // a formula that spills (or was a dynamic array in the file) goes back as one; a legacy array formula stays legacy
        if (region || cell.arraySpan?.spill) out.dynamicArray = true;
      }
      if (written) {
        // the spilled cells are not stored cells: they are written as the values the file keeps next to an array formula
        for (let r = 0; r < written.rows; r++) {
          for (let c = 0; c < written.cols; c++) {
            if (r === 0 && c === 0) continue;
            cells.set(`${row + r}:${col + c}`, { value: toFileValue(written.values[r][c] as SpreadsheetCell['value']) });
            maxRow = Math.max(maxRow, row + r);
            maxCol = Math.max(maxCol, col + c);
          }
        }
      }
    }
  }

  const columnWidths = new Map<number, number>();
  const hiddenCols = new Set<number>();
  for (let col = 0; col <= maxCol; col++) {
    columnWidths.set(col, sheet.getColWidth(col));
    if (sheet.isColHidden(col)) hiddenCols.add(col);
  }

  const rowHeights = new Map<number, number>();
  const hiddenRows = new Set<number>();
  for (let row = 0; row <= maxRow; row++) {
    rowHeights.set(row, sheet.getRowHeight(row));
    if (sheet.isRowHidden(row)) hiddenRows.add(row);
  }

  const mergedRanges: SpreadsheetMergedRange[] = sheet.getMergedRegions().map((r) => ({
    startRow: r.startRow,
    startCol: r.startCol,
    endRow: r.endRow,
    endCol: r.endCol,
  }));

  return {
    name: sheet.name,
    cells,
    columnWidths,
    rowHeights,
    hiddenCols,
    hiddenRows,
    frozenRows: sheet.getFrozenRows(),
    frozenCols: sheet.getFrozenCols(),
    mergedRanges,
  };
}

function toSpreadsheetCell(
  cell: Cell,
  stylePool: ReturnType<WorkbookImpl['getStylePool']>,
  formatPool: ReturnType<WorkbookImpl['getFormatPool']>
): SpreadsheetCell {
  const out: SpreadsheetCell = { value: cell.value };
  // A formula's saved result must be a value Excel knows. An engine message such as
  // "#NAME? Function X not found" is reduced to its error code, and the "#ERROR! ..." of a
  // formula that does not parse is not written at all (the formula has no result).
  if (cell.formula && !cell.cachedOnly) out.value = toFileValue(cell.value);
  if (cell.formula) out.formula = cell.formula.startsWith('=') ? cell.formula.slice(1) : cell.formula;
  // a formula the engine could not evaluate goes back as it came: its text exactly, next to the result the file saved
  if (cell.formula && cell.cachedOnly) out.cachedOnly = true;
  if (cell.hyperlink) out.hyperlink = cell.hyperlink;
  if (cell.styleId) {
    const style = stylePool.get(cell.styleId);
    if (style) out.style = toSpreadsheetStyle(style);
  }
  if (cell.formatId) {
    const format = formatPool.get(cell.formatId);
    if (format) out.numberFormat = { code: toExcelNumberFormatCode(format) };
  }
  return out;
}

/** A value as the file stores it: an engine error message (`#NAME? Function X not found`) is reduced to its code, a parse failure is no value. */
function toFileValue(value: SpreadsheetCell['value']): SpreadsheetCell['value'] {
  if (typeof value === 'string' && value.startsWith('#')) {
    if (value.startsWith('#ERROR!')) return null;
    if (value.includes(' ')) return value.split(' ')[0];
  }
  return value;
}

function toSpreadsheetStyle(style: CellStyle): SpreadsheetCellStyle {
  const out: SpreadsheetCellStyle = {};
  if (style.bold != null) out.bold = style.bold;
  if (style.italic != null) out.italic = style.italic;
  if (style.textDecoration === 'underline') out.underline = true;
  if (style.textDecoration === 'line-through') out.strikethrough = true;
  if (style.fontFamily != null) out.fontFamily = style.fontFamily;
  if (style.fontSize != null) out.fontSize = style.fontSize;
  if (style.fontColor != null) out.fontColor = style.fontColor;
  if (style.backgroundColor != null) out.backgroundColor = style.backgroundColor;
  if (style.textAlign != null) out.horizontalAlign = style.textAlign;
  if (style.verticalAlign != null) out.verticalAlign = style.verticalAlign;
  if (style.textWrap != null) out.wrapText = style.textWrap;
  if (style.textRotation != null) out.textRotation = style.textRotation;

  if (style.borderTop || style.borderRight || style.borderBottom || style.borderLeft) {
    out.border = {
      top: parseBorder(style.borderTop),
      right: parseBorder(style.borderRight),
      bottom: parseBorder(style.borderBottom),
      left: parseBorder(style.borderLeft),
    };
  }
  return out;
}

// CellStyle's border fields are CSS shorthand strings, e.g. "1px solid #000000".
const CSS_BORDER = /^(\d+)px\s+(\w+)\s+(#[0-9a-fA-F]{3,8}|\w+)$/;

function parseBorder(value: string | undefined): SpreadsheetBorder | undefined {
  if (!value) return undefined;
  const match = value.match(CSS_BORDER);
  if (!match) return { style: 'thin' };
  const [, widthStr, cssStyle, color] = match;
  return {
    style: toBorderStyle(Number(widthStr), cssStyle),
    color: color.startsWith('#') ? color : undefined,
  };
}

function toBorderStyle(widthPx: number, cssStyle: string): SpreadsheetBorder['style'] {
  if (cssStyle === 'dashed') return 'dashed';
  if (cssStyle === 'dotted') return 'dotted';
  if (cssStyle === 'double') return 'double';
  if (widthPx >= 3) return 'thick';
  if (widthPx === 2) return 'medium';
  return 'thin';
}

// ---------------------------------------------------------------------------
// Import
// ---------------------------------------------------------------------------

/** What `importFromXlsx` takes: the reader's options, and whether to recalculate every formula the engine can evaluate instead of keeping the results the file saved. */
export type ImportFromXlsxOptions = ImportXlsxOptions & {
  /**
   * Recalculate every formula the engine can evaluate on load (default `false`: the file's saved results are shown and a formula
   * is recalculated when a precedent changes or the user edits it). A formula the engine cannot evaluate keeps its saved
   * result either way.
   */
  recalculateOnLoad?: boolean;
  /**
   * The most cells the array formulas of one sheet may cover on load (default 1,000,000, `DEFAULT_MAX_ARRAY_CELLS`: the file chooses
   * the ranges). A sheet whose array formulas would cover more keeps the rest as cached-only formulas (one cell, the saved result),
   * reported as `xlsx.package.sizeLimit` (`array formula(s) kept as saved values`).
   */
  maxArrayCells?: number;
};

/**
 * Replace `workbook`'s entire contents with what's in the xlsx file at
 * `bytes` — every existing sheet is discarded, matching what "import a file"
 * should mean. Merges are applied after the bulk load via the live
 * `mergeCellsBulk` (the same result as `mergeCells` for each, in O(n log^2 n) and without history) rather than through setData's
 * config, since merged regions (unlike cells/widths/heights) only accept stable row/col IDs, not the numeric indices this adapter
 * otherwise works in. Overlapping merges (Excel writes none) are resolved as one by one would: the later wins.
 *
 * Throws an `ImportError` (from `@b32nio/spindle-io/xlsx`, with a `code`) for a file that cannot or must not be opened:
 * encrypted, legacy binary, not a package, over the limits (`options.limits`), or corrupt. The workbook is left untouched.
 *
 * `loss` is what the file holds that the editor cannot take (the detector's items) and, with id `xlsx.package.sizeLimit`, what the
 * import limits cut from it (cells past `maxTotalCells`, sheets past `maxSheets`...), each naming the option that raises the limit.
 */
export async function importFromXlsx(
  bytes: Uint8Array,
  workbook: WorkbookImpl,
  options: ImportFromXlsxOptions = {}
): Promise<{ warnings: string[]; loss: ImportLossReport }> {
  const io = await loadIo();
  const { recalculateOnLoad, maxArrayCells = DEFAULT_MAX_ARRAY_CELLS, ...readerOptions } = options;
  // the importer is the one that refuses unsafe files (an `ImportError`), so it goes first: no point scanning what is refused
  const result = await io.importXlsx(bytes, readerOptions);
  const detected = await io.detectXlsxLoss(bytes, lossOptions(io));
  // what the import limits cut is only known to the reader, so its items join the detector's: the host gets one report
  const loss: ImportLossReport = result.limitLoss?.length ? { ...detected, items: [...detected.items, ...result.limitLoss] } : detected;
  // setData() replaces the whole workbook; a file with no worksheets (e.g. a zip that isn't
  // really a workbook) would leave it with no sheets at all, so refuse it up front.
  if (result.sheets.length === 0) throw new Error('The file contains no worksheets.');
  const warnings = [...result.warnings];
  // one warning per array formula the file gave a range that does not start at its cell, counted after a small cap (the file chooses how many)
  let misplaced = 0;
  const { data, sheetIds, spanBudgetCut } = toWorkbookData(result, workbook, io, maxArrayCells, (message) => {
    if (misplaced++ < MAX_ARRAY_WARNINGS) warnings.push(message);
  });
  if (misplaced > MAX_ARRAY_WARNINGS) warnings.push(`${misplaced - MAX_ARRAY_WARNINGS} more array formulas whose range does not start at their cell were read as one-cell formulas.`);
  if (spanBudgetCut > 0) loss.items.push(spanBudgetItem(spanBudgetCut, maxArrayCells));
  // the file's saved results are the values (the engine recalculates a formula when its precedents change), unless the host asks for a recalculation
  workbook.setData(data, { keepCachedValues: recalculateOnLoad !== true });
  // all the merges of a sheet at once: one by one each looked at all the others (10,000 took 12 s, 20,000 51 s). No history: a load
  // replaces the document, and `setData` has just emptied the undo stack; one step per merge would let undo take the merges back one by one
  result.sheets.forEach((sheet, i) => {
    if (sheet.mergedRanges) workbook.mergeCellsBulk(sheet.mergedRanges, sheetIds[i], { history: false });
  });
  return { warnings, loss };
}

/**
 * True when a number format code survives import and export unchanged (a structured type or `custom` with the
 * verbatim code, see `fromExcelNumberFormatCode`) and the engine displays it the way Excel does. Only codes the
 * engine cannot read (an Excel-invalid code, a non-Gregorian calendar, `[DBNum1]`) or would show with another
 * language's month and day names are left: they are kept verbatim but shown in General or in English.
 */
function isNumberFormatExact(code: string): boolean {
  return isFaithfulFormatCode(normalizeExcelNumberFormatCode(code));
}

/** What the detector needs to know about this editor: its function catalog and its number-format mapping. */
function lossOptions(io: IoXlsx, detail?: DetectXlsxLossOptions['detail']): DetectXlsxLossOptions {
  return {
    detail,
    knownFunctions: getRegisteredFunctionNames(),
    isNumberFormatExact,
    isFormulaEvaluable: (formula) => isFormulaEvaluable(io.fromFileFormula(formula)),
  };
}

/**
 * What importing this .xlsx would not bring across (conditional formats, validation, names,
 * charts, comments, formulas the engine cannot calculate…). Read-only: nothing is imported or modified.
 * `detail: 'full'` (default `'summary'`, what the import dialog shows) adds the finer ids the fidelity harness records.
 */
export function inspectXlsx(bytes: Uint8Array, options: Pick<DetectXlsxLossOptions, 'detail'> = {}): Promise<ImportLossReport> {
  return loadIo().then((io) => io.detectXlsxLoss(bytes, lossOptions(io, options.detail)));
}

function toWorkbookData(
  result: SpreadsheetImportResult,
  workbook: WorkbookImpl,
  io: IoXlsx,
  maxArrayCells: number,
  warn: (message: string) => void
): { data: WorkbookData; sheetIds: string[]; spanBudgetCut: number } {
  const stylePool = new StylePool();
  const formatPool = new FormatPool();
  const sheetIds = result.sheets.map(() => generateId());
  let spanBudgetCut = 0;
  const sheets: SheetData[] = result.sheets.map((sheet, i) => {
    const converted = toSheetData(sheet, sheetIds[i], stylePool, formatPool, io, maxArrayCells, warn);
    spanBudgetCut += converted.spanBudgetCut;
    return converted.data;
  });

  const activeIndex = result.activeSheetIndex ?? 0;
  const data: WorkbookData = {
    id: workbook.id,
    name: workbook.name,
    activeSheetId: sheetIds[activeIndex] ?? sheetIds[0],
    defaultRowHeight: workbook.defaultRowHeight,
    defaultColWidth: workbook.defaultColWidth,
    stylePool: Object.fromEntries(stylePool.getAllStyles()),
    formatPool: Object.fromEntries(formatPool.getAllFormats()),
    sheets,
  };
  return { data, sheetIds, spanBudgetCut };
}

/**
 * What the array formulas of a sheet may cost on load, in cells: a dynamic-array formula clears the cells its saved result
 * occupies (at most as many as the sheet has), a legacy one keeps the extent the file gave it, in full. The file chooses the
 * extent (`ref="A1:XFD1048576"`), so a sheet whose array formulas would cost more keeps the rest as cached-only formulas. A million
 * cells is the 100,000 x 10 performance file; `ref="A1:J400000"` (4 million) cost 5 s and 1 GB to import and 44 s and 3 GB to write
 * back at the former 5 million. A host raises or lowers it with the `maxArrayCells` option of `importFromXlsx`.
 */
export const DEFAULT_MAX_ARRAY_CELLS = 1_000_000;

/** Warnings about array ranges that do not start at their cell that an import lists before it only counts the rest. */
const MAX_ARRAY_WARNINGS = 20;

function toSheetData(
  sheet: SpreadsheetSheet,
  id: string,
  stylePool: StylePool,
  formatPool: FormatPool,
  io: IoXlsx,
  maxArrayCells: number,
  warn: (message: string) => void
): { data: SheetData; spanBudgetCut: number } {
  const cells: SheetData['cells'] = [];
  let maxRow = 0;
  let maxCol = 0;
  const arrays: Array<{ row: number; col: number; rows: number; cols: number; cell: Cell }> = [];

  for (const [key, spreadsheetCell] of sheet.cells) {
    const [rowStr, colStr] = key.split(':');
    maxRow = Math.max(maxRow, Number(rowStr));
    maxCol = Math.max(maxCol, Number(colStr));

    const cell: Cell = { value: spreadsheetCell.value };
    if (spreadsheetCell.formula) {
      // the editor form of the text (no `_xlfn.` and friends) when the engine can evaluate it; otherwise the file's text as it is
      const text = io.fromFileFormula(spreadsheetCell.formula);
      if (isFormulaEvaluable(text)) cell.formula = `=${text}`;
      else {
        cell.formula = `=${spreadsheetCell.formula}`;
        cell.cachedOnly = true;
      }
      const info = spreadsheetCell.formulaInfo;
      const range = info?.type === 'array' && info.ref !== undefined ? parseRange(info.ref) : undefined;
      if (range) {
        const row = Number(rowStr);
        const col = Number(colStr);
        if (range.row === row && range.col === col) {
          const span = { rows: range.rows, cols: range.cols };
          cell.arraySpan = info?.dynamic ? { ...span, spill: true } : span;
          arrays.push({ row, col, ...span, cell });
        } else warn(`${sheet.name.slice(0, 64)}!${spanRef(row, col, 1, 1)}: the array formula's range ${info?.ref} does not start at the cell, so it is read as a one-cell formula.`);
      }
    }
    if (spreadsheetCell.hyperlink) cell.hyperlink = spreadsheetCell.hyperlink;
    if (spreadsheetCell.style) cell.styleId = stylePool.getOrCreate(fromSpreadsheetStyle(spreadsheetCell.style));
    if (spreadsheetCell.numberFormat) {
      const format = fromExcelNumberFormatCode(spreadsheetCell.numberFormat.code);
      if (format.type) cell.formatId = formatPool.getOrCreate(format);
    }
    cells.push({ key, cell });
  }

  // The cells an evaluable array formula covers hold the result the file saved; the engine spills its own there, so they go
  // (the anchor keeps the saved value until it is evaluated on load). Past the budget the formula stays cached-only instead.
  let spanBudgetCut = 0;
  const cleared = new Set<string>();
  let budget = maxArrayCells;
  for (const { row, col, rows, cols, cell } of arrays) {
    if (cell.cachedOnly || (rows === 1 && cols === 1)) continue;
    const area = rows * cols;
    const cost = cell.arraySpan?.spill ? Math.min(area, cells.length) : area;
    if (cost > budget) {
      cell.formula = `=${(sheet.cells.get(`${row}:${col}`) as SpreadsheetCell).formula}`;
      cell.cachedOnly = true;
      spanBudgetCut++;
      continue;
    }
    budget -= cost;
    if (area <= cells.length) {
      for (let r = 0; r < rows; r++) for (let c = 0; c < cols; c++) if (r + c > 0) cleared.add(`${row + r}:${col + c}`);
    } else {
      for (const { key } of cells) {
        const [r, c] = key.split(':').map(Number);
        if (r >= row && r < row + rows && c >= col && c < col + cols && r + c > row + col) cleared.add(key);
      }
    }
  }
  const kept = cleared.size === 0 ? cells : cells.filter(({ key, cell }) => !(cleared.has(key) && !cell.formula));

  const config: SheetData['config'] = {
    frozenRows: sheet.frozenRows,
    frozenCols: sheet.frozenCols,
  };
  if (sheet.rowHeights?.size) config.rowHeights = Array.from(sheet.rowHeights.entries());
  if (sheet.columnWidths?.size) config.colWidths = Array.from(sheet.columnWidths.entries());
  if (sheet.hiddenRows?.size) config.hiddenRows = Array.from(sheet.hiddenRows);
  if (sheet.hiddenCols?.size) config.hiddenCols = Array.from(sheet.hiddenCols);

  return {
    data: {
      id,
      name: sheet.name,
      cells: kept,
      config,
      // Same generous default as a freshly-created sheet, at least as large as
      // the imported data.
      rowCount: Math.max(maxRow + 1, 1000),
      colCount: Math.max(maxCol + 1, 100),
    },
    spanBudgetCut,
  };
}

/** The loss item for array formulas the load budget kept as saved values (see DEFAULT_MAX_ARRAY_CELLS). */
function spanBudgetItem(count: number, maxArrayCells: number): ImportLossItem {
  return {
    id: 'xlsx.package.sizeLimit',
    label: 'array formula(s) kept as saved values',
    count,
    effect: `approximated: the cells they cover are too many to take on load (over ${maxArrayCells} cells of array formulas on a sheet; the maxArrayCells option raises it), so they keep their saved results and are not recalculated`,
  };
}

function fromSpreadsheetStyle(style: SpreadsheetCellStyle): CellStyle {
  const out: CellStyle = {};
  if (style.bold != null) out.bold = style.bold;
  if (style.italic != null) out.italic = style.italic;
  if (style.underline) out.textDecoration = 'underline';
  else if (style.strikethrough) out.textDecoration = 'line-through';
  if (style.fontFamily != null) out.fontFamily = style.fontFamily;
  if (style.fontSize != null) out.fontSize = style.fontSize;
  if (style.fontColor != null) out.fontColor = style.fontColor;
  if (style.backgroundColor != null) out.backgroundColor = style.backgroundColor;
  if (style.horizontalAlign != null) out.textAlign = style.horizontalAlign;
  if (style.verticalAlign != null) out.verticalAlign = style.verticalAlign;
  if (style.wrapText != null) out.textWrap = style.wrapText;
  // an explicit 0 is the default: the file writer leaves it out, so the model does not keep it either (open -> save -> open is stable)
  if (style.textRotation) out.textRotation = style.textRotation;

  if (style.border) {
    if (style.border.top) out.borderTop = toCssBorder(style.border.top);
    if (style.border.right) out.borderRight = toCssBorder(style.border.right);
    if (style.border.bottom) out.borderBottom = toCssBorder(style.border.bottom);
    if (style.border.left) out.borderLeft = toCssBorder(style.border.left);
  }
  return out;
}

const BORDER_WIDTH_PX: Record<SpreadsheetBorder['style'], number> = {
  thin: 1, medium: 2, thick: 3, dashed: 1, dotted: 1, double: 3,
};
const BORDER_CSS_STYLE: Record<SpreadsheetBorder['style'], string> = {
  thin: 'solid', medium: 'solid', thick: 'solid', dashed: 'dashed', dotted: 'dotted', double: 'double',
};

function toCssBorder(border: SpreadsheetBorder): string {
  return `${BORDER_WIDTH_PX[border.style]}px ${BORDER_CSS_STYLE[border.style]} ${border.color ?? '#000000'}`;
}
