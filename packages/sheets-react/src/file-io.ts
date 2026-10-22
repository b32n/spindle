import {
  WorkbookImpl,
  exportToCSV,
  exportToXlsx,
  importFromCSV,
  importFromXlsx,
  inspectCsv,
  inspectXlsx,
} from '@b32nio/spindle-sheets-core';
import { fileExtension, sanitizeFileName, stripExtension } from '@b32nio/spindle-shared/react';
import type { ImportLossReport } from '@b32nio/spindle-shared';

/** Formats the toolbar's Download menu can produce. */
export type SheetsExportFormat = 'xlsx' | 'csv';

export interface SheetsExportFormatInfo {
  id: SheetsExportFormat;
  label: string;
  extension: string;
  mimeType: string;
}

export const SHEETS_EXPORT_FORMATS: readonly SheetsExportFormatInfo[] = [
  {
    id: 'xlsx',
    label: 'Microsoft Excel (.xlsx)',
    extension: 'xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  },
  { id: 'csv', label: 'Comma-separated values (.csv)', extension: 'csv', mimeType: 'text/csv;charset=utf-8' },
];

/** `accept` value for the Open file input. */
export const SHEETS_IMPORT_ACCEPT = '.xlsx,.csv';

export interface SheetsFileImportResult {
  warnings: string[];
  /** What the file contained that could not be imported (empty `items` = nothing lost). */
  loss: ImportLossReport;
}

export interface SheetsFileExport {
  data: Uint8Array | string;
  fileName: string;
  mimeType: string;
}

/**
 * Build the downloadable file for `format`. XLSX covers every sheet; CSV is
 * inherently single-sheet and exports the active one.
 */
export async function exportWorkbookFile(
  workbook: WorkbookImpl,
  format: SheetsExportFormat,
  baseName: string = workbook.name
): Promise<SheetsFileExport> {
  const info = SHEETS_EXPORT_FORMATS.find((f) => f.id === format);
  if (!info) throw new Error(`Unsupported export format: ${String(format)}`);
  const fileName = `${sanitizeFileName(baseName, 'Workbook')}.${info.extension}`;
  const data = format === 'xlsx' ? await exportToXlsx(workbook) : exportToCSV(workbook);
  return { data, fileName, mimeType: info.mimeType };
}

/**
 * Replace `workbook`'s contents with the file's contents. Format is chosen by
 * extension (.xlsx or .csv); anything else is rejected. The file is only ever
 * read locally.
 */
export async function importWorkbookFile(workbook: WorkbookImpl, file: File): Promise<SheetsFileImportResult> {
  const ext = fileExtension(file.name);
  if (ext === 'xlsx') {
    return importFromXlsx(new Uint8Array(await file.arrayBuffer()), workbook);
  }
  if (ext === 'csv') {
    // "Open" semantics: build a fresh single-sheet workbook from the CSV and
    // swap it in, rather than merging into whatever was open before.
    const scratch = new WorkbookImpl(workbook.id, stripExtension(file.name) || workbook.name);
    const text = await file.text();
    const loss = inspectCsv(text);
    importFromCSV(text, scratch.getSheet());
    workbook.setData(scratch.getData());
    return { warnings: [], loss };
  }
  throw new Error(`Unsupported file type: ".${ext || '?'}". Open an .xlsx or .csv file.`);
}

/**
 * Headless pre-check for "Open": what importing `file` would lose, without touching the
 * workbook. Hosts that render their own confirmation call this, show `describeImportLoss(report)`
 * when `hasImportLoss(report)`, and only then call `importWorkbookFile`.
 */
export async function inspectWorkbookFile(file: File): Promise<ImportLossReport> {
  const ext = fileExtension(file.name);
  if (ext === 'xlsx') return inspectXlsx(new Uint8Array(await file.arrayBuffer()));
  if (ext === 'csv') return inspectCsv(await file.text());
  throw new Error(`Unsupported file type: ".${ext || '?'}". Open an .xlsx or .csv file.`);
}
