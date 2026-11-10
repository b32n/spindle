/**
 * @deprecated The XLSX implementation moved to `@b32nio/spindle-io`. These
 * functions still work (they load it on demand, so it must be installed) but
 * new code should import `@b32nio/spindle-io/xlsx`. The types stay here.
 */
import type { ImportLossReport } from '../loss';
import { loadIoXlsx } from '../io-loader';
import type { SpreadsheetExportInput, SpreadsheetImportResult } from './types';

export * from './types';
export * from '../loss';

/** @deprecated Use `exportXlsx` from `@b32nio/spindle-io/xlsx`. */
export async function exportXlsx(input: SpreadsheetExportInput): Promise<Uint8Array> {
  return (await loadIoXlsx()).exportXlsx(input);
}

/** @deprecated Use `importXlsx` from `@b32nio/spindle-io/xlsx`. */
export async function importXlsx(bytes: Uint8Array): Promise<SpreadsheetImportResult> {
  return (await loadIoXlsx()).importXlsx(bytes);
}

/** @deprecated Use `detectXlsxLoss` from `@b32nio/spindle-io/xlsx`. */
export async function detectXlsxLoss(bytes: Uint8Array): Promise<ImportLossReport> {
  return (await loadIoXlsx()).detectXlsxLoss(bytes);
}
