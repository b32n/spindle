import { exportCsv, importCsv, type CsvCellValue } from '@b32nio/spindle-shared/export/csv';
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

  const rows: CsvCellValue[][] = [];
  for (let row = 0; row <= maxRow; row++) {
    const csvRow: CsvCellValue[] = [];
    for (let col = 0; col <= maxCol; col++) {
      const cell = sheet.getCell(row, col);
      if (cell?.formula) {
        csvRow.push(cell.formula);
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

export function importFromCSV(csv: string, sheet: Sheet): void {
  const { rows } = importCsv(csv);
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
