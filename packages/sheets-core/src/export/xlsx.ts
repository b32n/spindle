import { exportXlsx, importXlsx } from '@b32nio/spindle-shared/export/xlsx';
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
import { toExcelNumberFormatCode, fromExcelNumberFormatCode } from './xlsx-number-format';

export async function exportToXlsx(workbook: WorkbookImpl): Promise<Uint8Array> {
  const sheetIds = Array.from(workbook.sheets.keys());
  const sheets = sheetIds.map((id) => toSpreadsheetSheet(workbook, workbook.sheets.get(id)!));
  const activeSheetIndex = sheetIds.indexOf(workbook.activeSheetId);
  return exportXlsx({ sheets, activeSheetIndex: activeSheetIndex >= 0 ? activeSheetIndex : undefined });
}

function toSpreadsheetSheet(workbook: WorkbookImpl, sheet: Sheet): SpreadsheetSheet {
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
    cells.set(`${row}:${col}`, toSpreadsheetCell(cell, stylePool, formatPool));
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
  if (cell.formula) out.formula = cell.formula.startsWith('=') ? cell.formula.slice(1) : cell.formula;
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

/**
 * Replace `workbook`'s entire contents with what's in the xlsx file at
 * `bytes` — every existing sheet is discarded, matching what "import a file"
 * should mean. Merges are applied after the bulk load via the live
 * mergeCells API rather than through setData's config, since merged regions
 * (unlike cells/widths/heights) only accept stable row/col IDs, not the
 * numeric indices this adapter otherwise works in.
 */
export async function importFromXlsx(bytes: Uint8Array, workbook: WorkbookImpl): Promise<{ warnings: string[] }> {
  const result = await importXlsx(bytes);
  const { data, sheetIds } = toWorkbookData(result, workbook);
  workbook.setData(data);
  result.sheets.forEach((sheet, i) => {
    sheet.mergedRanges?.forEach((range) => workbook.mergeCells(range, sheetIds[i]));
  });
  return { warnings: result.warnings };
}

function toWorkbookData(
  result: SpreadsheetImportResult,
  workbook: WorkbookImpl
): { data: WorkbookData; sheetIds: string[] } {
  const stylePool = new StylePool();
  const formatPool = new FormatPool();
  const sheetIds = result.sheets.map(() => generateId());
  const sheets: SheetData[] = result.sheets.map((sheet, i) => toSheetData(sheet, sheetIds[i], stylePool, formatPool));

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
  return { data, sheetIds };
}

function toSheetData(sheet: SpreadsheetSheet, id: string, stylePool: StylePool, formatPool: FormatPool): SheetData {
  const cells: SheetData['cells'] = [];
  let maxRow = 0;
  let maxCol = 0;

  for (const [key, spreadsheetCell] of sheet.cells) {
    const [rowStr, colStr] = key.split(':');
    maxRow = Math.max(maxRow, Number(rowStr));
    maxCol = Math.max(maxCol, Number(colStr));

    const cell: Cell = { value: spreadsheetCell.value };
    if (spreadsheetCell.formula) cell.formula = `=${spreadsheetCell.formula}`;
    if (spreadsheetCell.hyperlink) cell.hyperlink = spreadsheetCell.hyperlink;
    if (spreadsheetCell.style) cell.styleId = stylePool.getOrCreate(fromSpreadsheetStyle(spreadsheetCell.style));
    if (spreadsheetCell.numberFormat) {
      const format = fromExcelNumberFormatCode(spreadsheetCell.numberFormat.code);
      if (format.type) cell.formatId = formatPool.getOrCreate(format);
    }
    cells.push({ key, cell });
  }

  const config: SheetData['config'] = {
    frozenRows: sheet.frozenRows,
    frozenCols: sheet.frozenCols,
  };
  if (sheet.rowHeights?.size) config.rowHeights = Array.from(sheet.rowHeights.entries());
  if (sheet.columnWidths?.size) config.colWidths = Array.from(sheet.columnWidths.entries());
  if (sheet.hiddenRows?.size) config.hiddenRows = Array.from(sheet.hiddenRows);
  if (sheet.hiddenCols?.size) config.hiddenCols = Array.from(sheet.hiddenCols);

  return {
    id,
    name: sheet.name,
    cells,
    config,
    // Same generous default as a freshly-created sheet, at least as large as
    // the imported data.
    rowCount: Math.max(maxRow + 1, 1000),
    colCount: Math.max(maxCol + 1, 100),
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
  if (style.textRotation != null) out.textRotation = style.textRotation;

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
