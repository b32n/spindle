export { exportDocx } from './writer';
export { importDocx } from './reader';
export type { ImportDocxOptions } from './reader';
export { ImportError, type ImportErrorCode } from '../common/import-error';
export { DEFAULT_IMPORT_LIMITS, type ImportLimits } from '../common/safe-zip';
export { detectDocxLoss } from './loss';
export type { DetectDocxLossOptions } from './loss';
export type {
  DocRunBase,
  DocTextRun,
  DocLinkRun,
  DocComment,
  DocCommentThread,
  DocNamedStyle,
  DocImageRun,
  DocPageNumberRun,
  DocTotalPagesRun,
  DocBookmarkRun,
  DocContentControl,
  DocContentControlItem,
  DocContentControlRun,
  DocContentControlType,
  DocNoteRefRun,
  DocFieldRun,
  DocInline,
  DocParagraphStyle,
  DocTableCell,
  DocTableRow,
  DocBlock,
  DocHeaderFooterParagraph,
  DocHeaderFooter,
  DocSection,
  DocumentExportInput,
  DocumentExportResult,
  DocumentImportResult,
} from '@b32nio/spindle-shared/export/docx';
