// Type stub for the optional `@b32nio/spindle-io/xlsx` peer (see ../io-loader.ts and
// tsconfig `paths`). spindle-io depends on this package's types, so resolving the real
// package here would make the two build in a cycle.
import type { ImportLossReport } from '../loss';
import type { SpreadsheetExportInput, SpreadsheetImportResult } from '../xlsx/types';

export function exportXlsx(input: SpreadsheetExportInput): Promise<Uint8Array>;
export function importXlsx(bytes: Uint8Array): Promise<SpreadsheetImportResult>;
export function detectXlsxLoss(bytes: Uint8Array): Promise<ImportLossReport>;
