import { exportXlsx } from '@b32nio/spindle-shared/export/xlsx';
import type {
  SpreadsheetBorder,
  SpreadsheetCell,
  SpreadsheetCellStyle,
  SpreadsheetMergedRange,
  SpreadsheetSheet,
} from '@b32nio/spindle-shared/export/xlsx';
import type { Cell, CellStyle, Sheet } from '../types';
import type { WorkbookImpl } from '../workbook';
import { toExcelNumberFormatCode } from './xlsx-number-format';

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
