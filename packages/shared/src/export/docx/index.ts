/**
 * @deprecated The DOCX implementation moved to `@b32nio/spindle-io`. These
 * functions still work (they load it on demand, so it must be installed) but
 * new code should import `@b32nio/spindle-io/docx`. The types stay here.
 */
import type { ImportLossReport } from '../loss';
import { loadIoDocx } from '../io-loader';
import type { DocumentExportInput, DocumentExportResult, DocumentImportResult } from './types';

export * from './types';
export * from '../loss';

/** @deprecated Use `exportDocx` from `@b32nio/spindle-io/docx`. */
export async function exportDocx(input: DocumentExportInput): Promise<DocumentExportResult> {
  return (await loadIoDocx()).exportDocx(input);
}

/** @deprecated Use `importDocx` from `@b32nio/spindle-io/docx`. */
export async function importDocx(bytes: Uint8Array): Promise<DocumentImportResult> {
  return (await loadIoDocx()).importDocx(bytes);
}

/** @deprecated Use `detectDocxLoss` from `@b32nio/spindle-io/docx`. */
export async function detectDocxLoss(bytes: Uint8Array): Promise<ImportLossReport> {
  return (await loadIoDocx()).detectDocxLoss(bytes);
}
