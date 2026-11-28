// Type stub for the optional `@b32nio/spindle-io/pptx` peer (see ../io-loader.ts and
// tsconfig `paths`). spindle-io depends on this package's types, so resolving the real
// package here would make the two build in a cycle.
import type { ImportLossReport } from '../loss';
import type { DetectPptxLossOptions, PresentationImportResult } from '../pptx/types';

export interface PptxExportResult {
  bytes: Uint8Array;
  warnings: string[];
}

export function importPptx(bytes: Uint8Array, options?: Record<string, unknown>): Promise<PresentationImportResult>;
export function exportPptx(presentation: PresentationImportResult, options?: Record<string, unknown>): Promise<PptxExportResult>;
export function detectPptxLoss(bytes: Uint8Array, options?: DetectPptxLossOptions): Promise<ImportLossReport>;
