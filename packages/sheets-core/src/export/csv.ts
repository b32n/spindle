import { exportCsv, importCsv, detectCsvLoss, type CsvCellValue, type ImportLossReport } from '@b32nio/spindle-shared/export/csv';
import { csvTooLargeError } from './import-error';
import type { Sheet } from '../types';
import type { WorkbookImpl } from '../workbook';

export function exportToCSV(workbook: WorkbookImpl, sheetId?: string): string {
  const sheet = workbook.getSheet(sheetId);

  // Find the maximum row and column with data
  let maxRow = 0;
  let maxCol = 0;
  for (const [row, col] of sheet.entries()) {
    maxRow = Math.max(maxRow, row);
    maxCol = Math.max(maxCol, col);
  }
  const spill = workbook.getSpillExtent(sheetId);
  if (spill) {
    maxRow = Math.max(maxRow, spill.row);
    maxCol = Math.max(maxCol, spill.col);
  }

  const rows: CsvCellValue[][] = [];
  for (let row = 0; row <= maxRow; row++) {
    const csvRow: CsvCellValue[] = [];
    for (let col = 0; col <= maxCol; col++) {
      const cell = sheet.getCell(row, col);
      if (cell?.formula || workbook.isSpilledCell(sheetId, row, col)) {
        // Like Excel/Sheets/LibreOffice: a formula cell is saved as its computed value.
        const v = workbook.getCellCalculatedValue(sheetId, row, col);
        csvRow.push(typeof v === 'number' || typeof v === 'boolean' ? v : v != null ? String(v) : null);
      } else if (typeof cell?.value === 'number' || typeof cell?.value === 'boolean') {
        csvRow.push(cell.value);
      } else if (cell?.value != null) {
        csvRow.push(cell.value.toString());
      } else {
        csvRow.push(null);
      }
    }
    rows.push(csvRow);
  }

  return exportCsv({ rows });
}

/** Largest CSV text `importFromCSV` takes by default: the importers' package limit (100 MB, `DEFAULT_IMPORT_LIMITS`). */
export const DEFAULT_CSV_MAX_CHARS = 100 * 1024 * 1024;

export interface ImportCsvOptions {
  /** Longest text accepted, in characters (which never exceed the file's bytes); default `DEFAULT_CSV_MAX_CHARS`. */
  maxChars?: number;
}

/** A byte order mark at the start of the text is not data: `\uFEFFName` would otherwise name the first column wrongly. */
const stripBom = (csv: string): string => (csv.charCodeAt(0) === 0xfeff ? csv.slice(1) : csv);

/** What importing this CSV text would not bring across (currently: delimiters other than comma). */
export function inspectCsv(csv: string): ImportLossReport {
  return detectCsvLoss(stripBom(csv));
}

/**
 * Put CSV text into `sheet`. A leading byte order mark is dropped. Throws an error named `ImportError` with `code`
 * `size-limit` (the shape of the XLSX importer's `ImportError`, recognised by `isImportError`) for text over `options.maxChars`.
 */
export function importFromCSV(csv: string, sheet: Sheet, options: ImportCsvOptions = {}): void {
  const maxChars = options.maxChars ?? DEFAULT_CSV_MAX_CHARS;
  if (csv.length > maxChars) {
    throw csvTooLargeError(maxChars);
  }
  const { rows } = importCsv(stripBom(csv));
  // A sheet starts with 1000 rows x 100 columns; grow it so a larger file stays reachable
  // in the grid (the XLSX importer does the same).
  let widest = 0;
  for (const values of rows) widest = Math.max(widest, values.length);
  sheet.rowCount = Math.max(sheet.rowCount, rows.length);
  sheet.colCount = Math.max(sheet.colCount, widest);
  rows.forEach((values, row) => {
    values.forEach((value, col) => {
      // Set any non-empty field (including whitespace-only, which is a real
      // value between commas: `a, ,c`). Empty fields are left blank.
      if (value !== '') {
        sheet.setCellValue(row, col, value);
      }
    });
  });
}
