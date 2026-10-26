// @b32nio/spindle-io — XLSX and DOCX readers/writers (and the CSV helpers,
// re-exported for one-stop imports). Import `@b32nio/spindle-io/xlsx` or
// `@b32nio/spindle-io/docx` to bundle only one format's dependencies.

export * from './xlsx';
export * from './docx';
export { exportCsv, importCsv, detectCsvLoss } from '@b32nio/spindle-shared/export/csv';
export type { CsvCellValue, CsvExportInput, CsvImportResult } from '@b32nio/spindle-shared/export/csv';
export {
  createLossReport,
  hasImportLoss,
  describeImportLoss,
  type ImportLossFormat,
  type ImportLossItem,
  type ImportLossReport,
} from '@b32nio/spindle-shared';
