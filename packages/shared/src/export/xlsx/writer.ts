import ExcelJS from 'exceljs';
import type {
  SpreadsheetBorder,
  SpreadsheetCell,
  SpreadsheetCellStyle,
  SpreadsheetExportInput,
  SpreadsheetSheet,
} from './types';

export async function exportXlsx(input: SpreadsheetExportInput): Promise<Uint8Array> {
  const workbook = new ExcelJS.Workbook();
  input.sheets.forEach((sheet) => writeSheet(workbook.addWorksheet(sheet.name), sheet));
  const buffer = await workbook.xlsx.writeBuffer();
  return new Uint8Array(buffer);
}

function writeSheet(worksheet: ExcelJS.Worksheet, sheet: SpreadsheetSheet): void {
  for (const [key, cell] of sheet.cells) {
    const [row, col] = key.split(':').map(Number);
    // exceljs is 1-based; our grid is 0-based.
    writeCell(worksheet.getCell(row + 1, col + 1), cell);
  }

  sheet.columnWidths?.forEach((px, col) => {
    worksheet.getColumn(col + 1).width = pxToExcelColumnWidth(px);
  });
  sheet.rowHeights?.forEach((px, row) => {
    worksheet.getRow(row + 1).height = pxToPoints(px);
  });
  sheet.hiddenCols?.forEach((col) => {
    worksheet.getColumn(col + 1).hidden = true;
  });
  sheet.hiddenRows?.forEach((row) => {
    worksheet.getRow(row + 1).hidden = true;
  });

  if (sheet.frozenRows || sheet.frozenCols) {
    worksheet.views = [{ state: 'frozen', xSplit: sheet.frozenCols || 0, ySplit: sheet.frozenRows || 0 }];
  }

  sheet.mergedRanges?.forEach((range) => {
    worksheet.mergeCells(range.startRow + 1, range.startCol + 1, range.endRow + 1, range.endCol + 1);
  });
}

function writeCell(excelCell: ExcelJS.Cell, cell: SpreadsheetCell): void {
  if (cell.formula) {
    // exceljs caches a formula's last-known result alongside its expression;
    // `undefined` result is fine (recalculated on open) but we have it, so
    // pass it through.
    excelCell.value = { formula: cell.formula, result: cell.value ?? undefined };
  } else if (cell.hyperlink) {
    // exceljs's hyperlink cell value can't also carry a formula — a cell with
    // both is a rare combination we don't attempt to represent.
    excelCell.value = { text: cell.value != null ? String(cell.value) : cell.hyperlink, hyperlink: cell.hyperlink };
  } else if (cell.value != null) {
    excelCell.value = cell.value;
  }
  if (cell.numberFormat) excelCell.numFmt = cell.numberFormat.code;
  if (cell.style) applyStyle(excelCell, cell.style);
}

function applyStyle(excelCell: ExcelJS.Cell, style: SpreadsheetCellStyle): void {
  const font: Partial<ExcelJS.Font> = {};
  if (style.bold != null) font.bold = style.bold;
  if (style.italic != null) font.italic = style.italic;
  if (style.underline != null) font.underline = style.underline;
  if (style.strikethrough != null) font.strike = style.strikethrough;
  if (style.fontFamily != null) font.name = style.fontFamily;
  if (style.fontSize != null) font.size = style.fontSize;
  if (style.fontColor != null) font.color = { argb: toArgb(style.fontColor) };
  if (Object.keys(font).length > 0) excelCell.font = font;

  if (style.backgroundColor) {
    excelCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: toArgb(style.backgroundColor) } };
  }

  const alignment: Partial<ExcelJS.Alignment> = {};
  if (style.horizontalAlign) alignment.horizontal = style.horizontalAlign;
  if (style.verticalAlign) alignment.vertical = style.verticalAlign;
  if (style.wrapText != null) alignment.wrapText = style.wrapText;
  if (style.textRotation != null) alignment.textRotation = style.textRotation;
  if (Object.keys(alignment).length > 0) excelCell.alignment = alignment;

  if (style.border) {
    excelCell.border = {
      top: toExcelBorder(style.border.top),
      right: toExcelBorder(style.border.right),
      bottom: toExcelBorder(style.border.bottom),
      left: toExcelBorder(style.border.left),
    };
  }
}

function toExcelBorder(border: SpreadsheetBorder | undefined): Partial<ExcelJS.Border> | undefined {
  if (!border) return undefined;
  return { style: border.style, color: border.color ? { argb: toArgb(border.color) } : undefined };
}

function toArgb(hex: string): string {
  const clean = hex.replace('#', '');
  return clean.length === 8 ? clean.toUpperCase() : `FF${clean.toUpperCase()}`;
}

// Excel column width is "characters of the default font that fit", not
// pixels. This is Microsoft's own documented approximation for Calibri 11
// (max digit width 7px): https://learn.microsoft.com/en-us/office/troubleshoot/excel/determine-column-widths
const EXCEL_MAX_DIGIT_WIDTH = 7;

function pxToExcelColumnWidth(px: number): number {
  return Math.round((((px - 5) / EXCEL_MAX_DIGIT_WIDTH) * 100 + 0.5)) / 100;
}

// Excel row height is in points (1/72in); our grid stores CSS pixels (96dpi).
function pxToPoints(px: number): number {
  return px * 0.75;
}
