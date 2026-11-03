// Type stub for the optional `@b32nio/spindle-io/docx` peer (see ../io-loader.ts and
// tsconfig `paths`). spindle-io depends on this package's types, so resolving the real
// package here would make the two build in a cycle.
import type { ImportLossReport } from '../loss';
import type { DocumentExportInput, DocumentExportResult, DocumentImportResult } from '../docx/types';

export function exportDocx(input: DocumentExportInput): Promise<DocumentExportResult>;
export function importDocx(bytes: Uint8Array): Promise<DocumentImportResult>;
export function detectDocxLoss(bytes: Uint8Array): Promise<ImportLossReport>;
