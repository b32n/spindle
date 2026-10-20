import { exportToDocx, importFromDocx, type DocumentImpl } from '@b32nio/spindle-docs-core';
import { fileExtension, sanitizeFileName, stripExtension } from '@b32nio/spindle-shared/react';

/** Formats the toolbar's Download menu can produce. */
export type DocsExportFormat = 'docx';

export interface DocsExportFormatInfo {
  id: DocsExportFormat;
  label: string;
  extension: string;
  mimeType: string;
}

export const DOCS_EXPORT_FORMATS: readonly DocsExportFormatInfo[] = [
  {
    id: 'docx',
    label: 'Microsoft Word (.docx)',
    extension: 'docx',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  },
];

/** `accept` value for the Open file input. */
export const DOCS_IMPORT_ACCEPT = '.docx';

export interface DocsFileExport {
  data: Uint8Array;
  fileName: string;
  mimeType: string;
  /** Features the exporter could not carry over. */
  warnings: string[];
}

export async function exportDocumentFile(
  document: DocumentImpl,
  format: DocsExportFormat,
  baseName: string = document.getTitle()
): Promise<DocsFileExport> {
  const info = DOCS_EXPORT_FORMATS.find((f) => f.id === format);
  if (!info) throw new Error(`Unsupported export format: ${String(format)}`);
  const { bytes, warnings } = await exportToDocx(document);
  return {
    data: bytes,
    fileName: `${sanitizeFileName(baseName, 'Document')}.${info.extension}`,
    mimeType: info.mimeType,
    warnings,
  };
}

/**
 * Replace `document`'s contents with the .docx file's contents. The file is
 * only ever read locally.
 */
export async function importDocumentFile(document: DocumentImpl, file: File): Promise<{ warnings: string[] }> {
  const ext = fileExtension(file.name);
  if (ext !== 'docx') {
    throw new Error(`Unsupported file type: ".${ext || '?'}". Open a .docx file.`);
  }
  const result = await importFromDocx(new Uint8Array(await file.arrayBuffer()), document);
  // A .docx without a title property keeps the old title; use the file name instead
  // when the importer found nothing better.
  if (!document.getTitle() || document.getTitle() === 'Untitled Document') {
    document.setTitle(stripExtension(file.name));
  }
  return result;
}
