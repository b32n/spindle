import ExcelJS from 'exceljs';
import type {
  SpreadsheetBorder,
  SpreadsheetCell,
  SpreadsheetCellStyle,
  SpreadsheetCellValue,
  SpreadsheetImportResult,
  SpreadsheetMergedRange,
  SpreadsheetSheet,
} from './types';

export async function importXlsx(bytes: Uint8Array): Promise<SpreadsheetImportResult> {
  const workbook = new ExcelJS.Workbook();
  // @ts-expect-error -- Buffer<ArrayBuffer> vs exceljs's plain Buffer type (see writer.test.ts)
  await workbook.xlsx.load(Buffer.from(bytes));
  const warnings: string[] = [];
  const sheets = workbook.worksheets.map((ws) => readSheet(ws, warnings));
  const activeTab = workbook.views?.[0]?.activeTab;
  return { sheets, activeSheetIndex: typeof activeTab === 'number' ? activeTab : undefined, warnings };
}

function readSheet(worksheet: ExcelJS.Worksheet, warnings: string[]): SpreadsheetSheet {
  const cells = new Map<string, SpreadsheetCell>();
  let maxRow = 0;
  let maxCol = 0;

  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
      // Only the merge anchor carries content; the rest of the region reads
      // through to it in exceljs. mergedRanges (below) re-establishes the merge.
      if (cell.isMerged && cell.master !== cell) return;

      const r = rowNumber - 1;
      const c = colNumber - 1;
      maxRow = Math.max(maxRow, r);
      maxCol = Math.max(maxCol, c);

      const { value, formula, hyperlink, warning } = readValue(cell);
      if (warning) warnings.push(`${worksheet.name}!${cell.address}: ${warning}`);

      const spreadsheetCell: SpreadsheetCell = { value };
      if (formula) spreadsheetCell.formula = formula;
      if (hyperlink) spreadsheetCell.hyperlink = hyperlink;
      const style = readStyle(cell);
      if (style) spreadsheetCell.style = style;
      if (cell.numFmt && cell.numFmt !== 'General') spreadsheetCell.numberFormat = { code: cell.numFmt };
      cells.set(`${r}:${c}`, spreadsheetCell);
    });
  });

  const columnWidths = new Map<number, number>();
  const hiddenCols = new Set<number>();
  for (let col = 0; col <= maxCol; col++) {
    const column = worksheet.getColumn(col + 1);
    if (column.width != null) columnWidths.set(col, excelColumnWidthToPx(column.width));
    if (column.hidden) hiddenCols.add(col);
  }

  const rowHeights = new Map<number, number>();
  const hiddenRows = new Set<number>();
  for (let row = 0; row <= maxRow; row++) {
    const r = worksheet.getRow(row + 1);
    if (r.height != null) rowHeights.set(row, pointsToPx(r.height));
    if (r.hidden) hiddenRows.add(row);
  }

  const frozenView = worksheet.views?.find(
    (v): v is ExcelJS.WorksheetViewFrozen & ExcelJS.WorksheetViewCommon => v.state === 'frozen'
  );

  const mergedRanges = (worksheet.model.merges ?? []).map(decodeRange);

  return {
    name: worksheet.name,
    cells,
    columnWidths,
    rowHeights,
    hiddenCols,
    hiddenRows,
    frozenRows: frozenView?.ySplit || undefined,
    frozenCols: frozenView?.xSplit || undefined,
    mergedRanges: mergedRanges.length > 0 ? mergedRanges : undefined,
  };
}

function readValue(cell: ExcelJS.Cell): {
  value: SpreadsheetCellValue;
  formula?: string;
  hyperlink?: string;
  warning?: string;
} {
  const raw = cell.value;
  if (raw == null) return { value: null };
  if (raw instanceof Date) return { value: dateToExcelSerial(raw) };
  if (typeof raw === 'object') {
    if ('formula' in raw) {
      return { value: normalizeResult(raw.result), formula: raw.formula };
    }
    if ('sharedFormula' in raw) {
      if (!raw.formula) {
        return {
          value: normalizeResult(raw.result),
          warning: 'shared formula without a resolvable expression; imported as a plain value',
        };
      }
      return { value: normalizeResult(raw.result), formula: raw.formula };
    }
    if ('hyperlink' in raw) {
      return { value: raw.text ?? null, hyperlink: raw.hyperlink };
    }
    if ('richText' in raw) {
      return { value: raw.richText.map((r) => r.text).join('') };
    }
    if ('error' in raw) {
      return { value: raw.error };
    }
    return { value: null, warning: 'unrecognized cell value shape' };
  }
  return { value: raw };
}

function normalizeResult(
  result: number | string | boolean | Date | ExcelJS.CellErrorValue | undefined
): SpreadsheetCellValue {
  if (result == null) return null;
  if (result instanceof Date) return dateToExcelSerial(result);
  if (typeof result === 'object' && 'error' in result) return result.error;
  return result;
}

// Standard xlsx date serial (days since 1899-12-30, including the Excel 1900
// leap-year quirk) — a property of the file format itself, not any editor's
// own model, so it belongs here rather than in a caller's adapter.
const EXCEL_EPOCH_UTC = Date.UTC(1899, 11, 30);
const DAY_MS = 24 * 60 * 60 * 1000;

function dateToExcelSerial(date: Date): number {
  return (date.getTime() - EXCEL_EPOCH_UTC) / DAY_MS;
}

const VALID_HORIZONTAL_ALIGN = new Set(['left', 'center', 'right']);
const VALID_VERTICAL_ALIGN = new Set(['top', 'middle', 'bottom']);
const VALID_BORDER_STYLES = new Set(['thin', 'medium', 'thick', 'dashed', 'dotted', 'double']);

function readStyle(cell: ExcelJS.Cell): SpreadsheetCellStyle | undefined {
  const style: SpreadsheetCellStyle = {};

  const font = cell.font;
  if (font) {
    if (font.bold) style.bold = true;
    if (font.italic) style.italic = true;
    if (font.underline) style.underline = true;
    if (font.strike) style.strikethrough = true;
    if (font.name) style.fontFamily = font.name;
    if (font.size) style.fontSize = font.size;
    if (font.color?.argb) style.fontColor = fromArgb(font.color.argb);
  }

  const fill = cell.fill;
  if (fill?.type === 'pattern' && fill.pattern === 'solid' && fill.fgColor?.argb) {
    style.backgroundColor = fromArgb(fill.fgColor.argb);
  }

  const alignment = cell.alignment;
  if (alignment) {
    if (alignment.horizontal && VALID_HORIZONTAL_ALIGN.has(alignment.horizontal)) {
      style.horizontalAlign = alignment.horizontal as SpreadsheetCellStyle['horizontalAlign'];
    }
    if (alignment.vertical && VALID_VERTICAL_ALIGN.has(alignment.vertical)) {
      style.verticalAlign = alignment.vertical as SpreadsheetCellStyle['verticalAlign'];
    }
    if (alignment.wrapText) style.wrapText = true;
    if (typeof alignment.textRotation === 'number') style.textRotation = alignment.textRotation;
  }

  const border = cell.border;
  if (border && (border.top || border.right || border.bottom || border.left)) {
    style.border = {
      top: readBorder(border.top),
      right: readBorder(border.right),
      bottom: readBorder(border.bottom),
      left: readBorder(border.left),
    };
  }

  return Object.keys(style).length > 0 ? style : undefined;
}

function readBorder(border: Partial<ExcelJS.Border> | undefined): SpreadsheetBorder | undefined {
  if (!border?.style) return undefined;
  const style = VALID_BORDER_STYLES.has(border.style) ? (border.style as SpreadsheetBorder['style']) : 'thin';
  return { style, color: border.color?.argb ? fromArgb(border.color.argb) : undefined };
}

function fromArgb(argb: string): string {
  // Drop the leading alpha byte — we only ever write/read opaque colors.
  return `#${argb.slice(-6).toLowerCase()}`;
}

function decodeAddress(address: string): { row: number; col: number } {
  const match = address.match(/^([A-Z]+)(\d+)$/);
  if (!match) return { row: 0, col: 0 };
  const [, colLetters, rowDigits] = match;
  let col = 0;
  for (const ch of colLetters) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { row: Number(rowDigits) - 1, col: col - 1 };
}

function decodeRange(range: string): SpreadsheetMergedRange {
  const [start, end] = range.split(':');
  const s = decodeAddress(start);
  const e = decodeAddress(end ?? start);
  return { startRow: s.row, startCol: s.col, endRow: e.row, endCol: e.col };
}

// Inverse of writer.ts's pxToExcelColumnWidth (Microsoft's documented
// Calibri-11 approximation); both directions are already best-effort.
const EXCEL_MAX_DIGIT_WIDTH = 7;

function excelColumnWidthToPx(excelWidth: number): number {
  return Math.round(excelWidth * EXCEL_MAX_DIGIT_WIDTH) + 5;
}

function pointsToPx(points: number): number {
  return points / 0.75;
}
