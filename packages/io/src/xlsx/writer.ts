import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { toFileFormula } from './future-functions';
import type {
  SpreadsheetBorder,
  SpreadsheetCell,
  SpreadsheetCellStyle,
  SpreadsheetExportInput,
  SpreadsheetSheet,
} from '@b32nio/spindle-shared/export/xlsx';

export async function exportXlsx(input: SpreadsheetExportInput): Promise<Uint8Array> {
  const workbook = new ExcelJS.Workbook();
  // the cells of each sheet (by part name) that are dynamic-array formulas: exceljs cannot write their cell metadata, so it is added to the package
  const dynamic = new Map<string, Set<string>>();
  input.sheets.forEach((sheet) => {
    const worksheet = workbook.addWorksheet(sheet.name);
    const addresses = writeSheet(worksheet, sheet);
    if (addresses.size > 0) dynamic.set(`xl/worksheets/sheet${worksheet.id}.xml`, addresses);
  });

  // The sheet that was active in the editor is the one Excel opens on.
  const last = input.sheets.length - 1;
  const active = Math.min(Math.max(input.activeSheetIndex ?? 0, 0), Math.max(last, 0));
  if (last >= 0) {
    workbook.views = [{ x: 0, y: 0, width: 28800, height: 17000, firstSheet: 0, activeTab: active, visibility: 'visible' }];
  }
  const buffer = await workbook.xlsx.writeBuffer();
  const bytes = new Uint8Array(buffer);
  return dynamic.size > 0 ? addDynamicArrayMetadata(bytes, dynamic) : bytes;
}

/** The metadata part Excel writes for dynamic-array formulas (`XLDAPR`): every cell marked `cm="1"` refers to its one entry. */
const DYNAMIC_ARRAY_METADATA =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">' +
  '<metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes>' +
  '<futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata>' +
  '<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>';

/** Add the dynamic-array metadata part and mark the given cells (addresses per sheet part) with it. */
async function addDynamicArrayMetadata(bytes: Uint8Array, dynamic: Map<string, Set<string>>): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(bytes);
  for (const [part, addresses] of dynamic) {
    // (exceljs wrote the part: the name comes from the worksheet it was made for)
    const xml = await zip.file(part)!.async('string');
    zip.file(
      part,
      xml.replace(/<c r="([A-Z]{1,3}[0-9]+)"/g, (whole, address: string) => (addresses.has(address) ? `<c r="${address}" cm="1"` : whole))
    );
  }
  zip.file('xl/metadata.xml', DYNAMIC_ARRAY_METADATA);
  const types = await zip.file('[Content_Types].xml')!.async('string');
  zip.file(
    '[Content_Types].xml',
    types.replace('</Types>', '<Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/></Types>')
  );
  const rels = await zip.file('xl/_rels/workbook.xml.rels')!.async('string');
  zip.file(
    'xl/_rels/workbook.xml.rels',
    rels.replace('</Relationships>', '<Relationship Id="rIdMetadata1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata" Target="metadata.xml"/></Relationships>')
  );
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

/** Write a sheet; returns the addresses of its dynamic-array formula cells. */
function writeSheet(worksheet: ExcelJS.Worksheet, sheet: SpreadsheetSheet): Set<string> {
  const dynamicCells = new Set<string>();
  for (const [key, cell] of sheet.cells) {
    const [row, col] = key.split(':').map(Number);
    // exceljs is 1-based; our grid is 0-based.
    const excelCell = worksheet.getCell(row + 1, col + 1);
    writeCell(excelCell, cell);
    if (cell.formula && cell.dynamicArray && cell.arrayRef !== undefined && ARRAY_REF.test(cell.arrayRef)) dynamicCells.add(excelCell.address);
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
  return dynamicCells;
}

/** Engine-internal marker for a circular reference; not a value Excel knows, so it is never written. */
const INTERNAL_CIRCULAR = '#CIRCULAR!';

/** An A1 cell or range: all the writer accepts as the extent of an array formula. */
const ARRAY_REF = /^[A-Z]{1,3}[0-9]{1,7}(?::[A-Z]{1,3}[0-9]{1,7})?$/;

function writeCell(excelCell: ExcelJS.Cell, cell: SpreadsheetCell): void {
  // Excel evaluates a circular reference to 0.
  if (cell.value === INTERNAL_CIRCULAR) cell = { ...cell, value: 0 };
  if (cell.formula) {
    // exceljs caches a formula's last-known result alongside its expression;
    // `undefined` result is fine (recalculated on open) but we have it, so
    // pass it through.
    // a formula the engine could not evaluate is written as the file had it; any other gets the prefixes the file format needs
    const formula = cell.cachedOnly ? cell.formula : toFileFormula(cell.formula);
    excelCell.value =
      cell.arrayRef === undefined || !ARRAY_REF.test(cell.arrayRef)
        ? { formula, result: cell.value ?? undefined }
        : ({ formula, result: cell.value ?? undefined, shareType: 'array', ref: cell.arrayRef } as unknown as ExcelJS.CellValue);
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
