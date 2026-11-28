import { exportToPptx, importFromPptx, inspectPptx, type DeckImpl } from '@b32nio/spindle-slides-core';
import type { ImportLossReport } from '@b32nio/spindle-shared';
import { downloadFile, fileExtension, sanitizeFileName } from '@b32nio/spindle-shared/react';

/** Formats the toolbar's Download control can produce. */
export type SlidesExportFormat = 'pptx';

export interface SlidesExportFormatInfo {
  id: SlidesExportFormat;
  label: string;
  extension: string;
  mimeType: string;
}

export const SLIDES_EXPORT_FORMATS: readonly SlidesExportFormatInfo[] = [
  {
    id: 'pptx',
    label: 'Microsoft PowerPoint (.pptx)',
    extension: 'pptx',
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  },
];

/** `accept` value for the Open file input. */
export const SLIDES_IMPORT_ACCEPT = '.pptx';

export interface SlidesFileExport {
  data: Uint8Array;
  fileName: string;
  mimeType: string;
  /** What the writer left out or changed. */
  warnings: string[];
}

export async function exportDeckFile(
  deck: DeckImpl,
  format: SlidesExportFormat,
  baseName: string = deck.getTitle()
): Promise<SlidesFileExport> {
  const info = SLIDES_EXPORT_FORMATS.find((f) => f.id === format);
  if (!info) throw new Error(`Unsupported export format: ${String(format)}`);
  const { bytes, warnings } = await exportToPptx(deck);
  return {
    data: bytes,
    fileName: `${sanitizeFileName(baseName, 'Presentation')}.${info.extension}`,
    mimeType: info.mimeType,
    warnings,
  };
}

function assertPptx(file: File): void {
  const ext = fileExtension(file.name);
  if (ext !== 'pptx') throw new Error(`Unsupported file type: ".${ext || '?'}". Open a .pptx file.`);
}

/**
 * Replace `deck`'s contents with the .pptx file's contents. The file is only ever read locally. An unsafe or unreadable file throws
 * (an `ImportError` with a `code`) and leaves the deck as it was.
 */
export async function importDeckFile(
  deck: DeckImpl,
  file: File
): Promise<{ warnings: string[]; loss: ImportLossReport }> {
  assertPptx(file);
  return importFromPptx(new Uint8Array(await file.arrayBuffer()), deck);
}

/**
 * Headless pre-check for "Open": what importing `file` would lose, and what it keeps without being able to edit (charts,
 * transitions...), without touching the deck. Hosts that render their own confirmation call this, show `describeImportLoss(report)`
 * when `hasImportLoss(report)`, then call `importDeckFile`.
 */
export async function inspectDeckFile(file: File): Promise<ImportLossReport> {
  assertPptx(file);
  return inspectPptx(new Uint8Array(await file.arrayBuffer()));
}

/** What the editor's Open flow reports to the host. */
export interface DeckOpenHandlers {
  /** The confirmation policy (`useImportConfirm`): resolves `false` when the user cancels. */
  confirm: (file: File, report: ImportLossReport) => Promise<boolean>;
  onFileOpened?: (info: { fileName: string; warnings: string[]; loss: ImportLossReport }) => void;
  onFileError?: (error: Error, action: 'open' | 'download') => void;
}

const asError = (err: unknown): Error => (err instanceof Error ? err : new Error(String(err)));

/** The toolbar's Open: pre-check (the file is parsed twice, here and in the import; revisit for the K17 100/200-slide benchmark), ask before replacing the deck when the file loses something, import, report. Never throws. */
export async function openDeckFile(deck: DeckImpl, file: File, handlers: DeckOpenHandlers): Promise<void> {
  try {
    if (!(await handlers.confirm(file, await inspectDeckFile(file)))) return;
    const { warnings, loss } = await importDeckFile(deck, file);
    handlers.onFileOpened?.({ fileName: file.name, warnings, loss });
  } catch (err) {
    handlers.onFileError?.(asError(err), 'open');
  }
}

/** The toolbar's Download: build the file, save it, tell the host what the writer left out. Never throws. */
export async function downloadDeckFile(
  deck: DeckImpl,
  format: SlidesExportFormat,
  handlers: { fileName?: string; onFileError?: DeckOpenHandlers['onFileError']; onExportWarning?: (warnings: string[]) => void }
): Promise<void> {
  try {
    const out = await exportDeckFile(deck, format, handlers.fileName);
    downloadFile(out.data, out.fileName, out.mimeType);
    if (out.warnings.length > 0) handlers.onExportWarning?.(out.warnings);
  } catch (err) {
    handlers.onFileError?.(asError(err), 'download');
  }
}
