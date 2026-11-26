import { exportToXlsx, importFromXlsx, inspectXlsx } from '../export/xlsx';
import { WorkbookImpl } from '../workbook';
import { FormulaParser } from '../formula-parser/parser';
import {
  createApi,
  type SheetsCoreModule,
  type XlsxEditorApi,
} from '../../../../scripts/xlsx-fidelity/api';

/** The harness's editor API bound to sheets-core source (the CLI binds the built package instead). */
export const sourceApi: XlsxEditorApi = createApi({
  WorkbookImpl,
  importFromXlsx,
  exportToXlsx,
  inspectXlsx,
  FormulaParser,
} as unknown as SheetsCoreModule);
