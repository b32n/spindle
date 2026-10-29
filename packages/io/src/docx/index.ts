export { exportDocx } from './writer';
export { importDocx } from './reader';
export { detectDocxLoss } from './loss';
export type {
  DocRunBase,
  DocTextRun,
  DocLinkRun,
  DocImageRun,
  DocPageNumberRun,
  DocTotalPagesRun,
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
