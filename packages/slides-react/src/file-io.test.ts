import type { DeckImpl } from '@b32nio/spindle-slides-core';

// The core's PPTX functions load `@b32nio/spindle-io/pptx` with a dynamic import, which this package's CommonJS jest cannot do; the
// real reader and writer are exercised by the slides-core suite and by the E2E specs (e2e/slides/file-io.spec.ts). These tests pin
// what this layer adds: the file type check, the names, the confirmation, the reporting and the error handling.
const mockImport = jest.fn();
const mockInspect = jest.fn();
const mockExport = jest.fn();
jest.mock('@b32nio/spindle-slides-core', () => ({
  importFromPptx: (...a: unknown[]) => mockImport(...a),
  inspectPptx: (...a: unknown[]) => mockInspect(...a),
  exportToPptx: (...a: unknown[]) => mockExport(...a),
}));
const mockDownload = jest.fn();
jest.mock('@b32nio/spindle-shared/react', () => ({
  ...jest.requireActual('@b32nio/spindle-shared/react'),
  downloadFile: (...a: unknown[]) => mockDownload(...a),
}));

import {
  SLIDES_EXPORT_FORMATS,
  SLIDES_IMPORT_ACCEPT,
  downloadDeckFile,
  exportDeckFile,
  importDeckFile,
  inspectDeckFile,
  openDeckFile,
} from './file-io';

const deck = (title = 'Mine') => ({ getTitle: () => title }) as unknown as DeckImpl;
const pptx = (name = 'deck.pptx', bytes = [1, 2, 3]) => new File([new Uint8Array(bytes)], name);
const EMPTY = { items: [] };
const KEPT = { items: [{ id: 'pptx.transitions', label: 'transitions', count: 2, state: 'kept', effect: 'kept' }] };

beforeEach(() => {
  jest.resetAllMocks();
  mockImport.mockResolvedValue({ warnings: [], loss: EMPTY });
  mockInspect.mockResolvedValue(EMPTY);
  mockExport.mockResolvedValue({ bytes: new Uint8Array([9, 9]), warnings: [] });
});

describe('slides file I/O', () => {
  it('offers PPTX downloads and accepts .pptx', () => {
    expect(SLIDES_EXPORT_FORMATS.map((f) => f.id)).toEqual(['pptx']);
    expect(SLIDES_IMPORT_ACCEPT).toBe('.pptx');
  });

  it('imports the file bytes into the deck and returns the core result', async () => {
    const d = deck();
    const result = { warnings: ['w'], loss: KEPT };
    mockImport.mockResolvedValue(result);
    expect(await importDeckFile(d, pptx())).toBe(result);
    expect(mockImport).toHaveBeenCalledWith(new Uint8Array([1, 2, 3]), d);
  });

  it('inspects the file bytes', async () => {
    mockInspect.mockResolvedValue(KEPT);
    expect(await inspectDeckFile(pptx('A.PPTX'))).toBe(KEPT);
    expect(mockInspect).toHaveBeenCalledWith(new Uint8Array([1, 2, 3]));
  });

  it('refuses a file that is not a .pptx before reading it', async () => {
    await expect(importDeckFile(deck(), pptx('notes.txt'))).rejects.toThrow(/Unsupported file type: "\.txt"\. Open a \.pptx file/);
    await expect(inspectDeckFile(pptx('README'))).rejects.toThrow(/Unsupported file type: "\.\?"/);
    expect(mockImport).not.toHaveBeenCalled();
    expect(mockInspect).not.toHaveBeenCalled();
  });

  it('names the download after the title, sanitized, or an explicit base name', async () => {
    const out = await exportDeckFile(deck('Plan: 2026/Q1'), 'pptx');
    expect(out).toEqual({
      data: new Uint8Array([9, 9]),
      fileName: 'Plan_ 2026_Q1.pptx',
      mimeType: SLIDES_EXPORT_FORMATS[0].mimeType,
      warnings: [],
    });
    expect((await exportDeckFile(deck(), 'pptx', 'custom')).fileName).toBe('custom.pptx');
    expect((await exportDeckFile(deck(''), 'pptx')).fileName).toBe('Presentation.pptx');
  });

  it('rejects an unknown export format', async () => {
    await expect(exportDeckFile(deck(), 'odp' as never)).rejects.toThrow(/Unsupported export format: odp/);
    expect(mockExport).not.toHaveBeenCalled();
  });
});

describe('openDeckFile', () => {
  it('asks with the report, imports and reports the file', async () => {
    const confirm = jest.fn().mockResolvedValue(true);
    const onFileOpened = jest.fn();
    mockInspect.mockResolvedValue(KEPT);
    mockImport.mockResolvedValue({ warnings: ['w'], loss: KEPT });
    await openDeckFile(deck(), pptx('a.pptx'), { confirm, onFileOpened });
    expect(confirm).toHaveBeenCalledWith(expect.any(File), KEPT);
    expect(onFileOpened).toHaveBeenCalledWith({ fileName: 'a.pptx', warnings: ['w'], loss: KEPT });
  });

  it('leaves the deck alone when the confirmation is declined', async () => {
    const onFileOpened = jest.fn();
    await openDeckFile(deck(), pptx(), { confirm: async () => false, onFileOpened });
    expect(mockImport).not.toHaveBeenCalled();
    expect(onFileOpened).not.toHaveBeenCalled();
  });

  it('reports a failure to onFileError (wrapping a non-Error), without throwing', async () => {
    const onFileError = jest.fn();
    await openDeckFile(deck(), pptx('a.txt'), { confirm: async () => true, onFileError });
    expect(onFileError).toHaveBeenCalledWith(expect.any(Error), 'open');
    const odd = jest.fn();
    await openDeckFile(deck(), pptx(), { confirm: () => Promise.reject('boom'), onFileError: odd });
    expect(odd).toHaveBeenCalledWith(new Error('boom'), 'open');
    await expect(openDeckFile(deck(), pptx('a.txt'), { confirm: async () => true })).resolves.toBeUndefined();
  });
});

describe('downloadDeckFile', () => {
  it('saves the file under the base name', async () => {
    await downloadDeckFile(deck(), 'pptx', { fileName: 'out' });
    expect(mockDownload).toHaveBeenCalledWith(new Uint8Array([9, 9]), 'out.pptx', SLIDES_EXPORT_FORMATS[0].mimeType);
  });

  it('reports what the writer left out, only when something was', async () => {
    const onExportWarning = jest.fn();
    await downloadDeckFile(deck(), 'pptx', { onExportWarning });
    expect(onExportWarning).not.toHaveBeenCalled();
    mockExport.mockResolvedValue({ bytes: new Uint8Array(), warnings: ['one shape was not written'] });
    await downloadDeckFile(deck(), 'pptx', { onExportWarning });
    expect(onExportWarning).toHaveBeenCalledWith(['one shape was not written']);
    await expect(downloadDeckFile(deck(), 'pptx', {})).resolves.toBeUndefined();
  });

  it('reports a failure to onFileError, without throwing', async () => {
    const onFileError = jest.fn();
    await downloadDeckFile(deck(), 'odp' as never, { onFileError });
    expect(onFileError).toHaveBeenCalledWith(expect.any(Error), 'download');
    await expect(downloadDeckFile(deck(), 'odp' as never, {})).resolves.toBeUndefined();
  });
});
