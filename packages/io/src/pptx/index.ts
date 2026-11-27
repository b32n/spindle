export { detectPptxLoss } from './loss';
export { importPptx } from './reader';
export { exportPptx } from './writer';
export type { PptxExportMetadata, PptxExportOptions, PptxExportResult } from './writer';
export { lintPptxPackage } from './writer-lint';
export type { ImportPptxOptions } from './reader';
export type { DetectPptxLossOptions } from '@b32nio/spindle-shared/export/pptx';
export { ImportError } from '../common/import-error';
export type { ImportErrorCode } from '../common/import-error';
