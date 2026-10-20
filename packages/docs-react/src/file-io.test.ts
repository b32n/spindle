import { DocumentImpl } from '@b32nio/spindle-docs-core';
import { DOCS_EXPORT_FORMATS, exportDocumentFile, importDocumentFile } from './file-io';

const PAGE_CONFIG = {
  size: { w: 816, h: 1056 },
  margins: { top: 96, right: 96, bottom: 96, left: 96 },
  orientation: 'portrait' as const,
};

function makeDoc(text: string, title = 'Quarterly Plan'): DocumentImpl {
  const doc = new DocumentImpl('doc', title);
  doc.setData({
    id: 'doc',
    title,
    defaultPageConfig: PAGE_CONFIG,
    textStylePool: {},
    paragraphStylePool: {},
    sections: [
      {
        id: 'sec1',
        pageConfig: PAGE_CONFIG,
        blocks: [{ id: 'b1', type: 'paragraph', content: [{ type: 'text', text }] }],
      },
    ],
  });
  return doc;
}

describe('docs file I/O', () => {
  it('offers DOCX downloads', () => {
    expect(DOCS_EXPORT_FORMATS.map((f) => f.id)).toEqual(['docx']);
  });

  it('names the download after the title, sanitized', async () => {
    const out = await exportDocumentFile(makeDoc('x', 'Plan: 2026/Q1'), 'docx');
    expect(out.fileName).toBe('Plan_ 2026_Q1.docx');
    expect(out.data).toBeInstanceOf(Uint8Array);
  });

  it('honours an explicit base name', async () => {
    const out = await exportDocumentFile(makeDoc('x'), 'docx', 'custom');
    expect(out.fileName).toBe('custom.docx');
  });

  it('round-trips through DOCX', async () => {
    const out = await exportDocumentFile(makeDoc('Hello from Spindle'), 'docx');
    const target = makeDoc('old content');
    const file = new File([new Uint8Array(out.data)], out.fileName);
    await importDocumentFile(target, file);
    expect(JSON.stringify(target.getSections())).toContain('Hello from Spindle');
    expect(JSON.stringify(target.getSections())).not.toContain('old content');
  });

  it('rejects non-docx files without touching the document', async () => {
    const target = makeDoc('keep me');
    await expect(importDocumentFile(target, new File(['hi'], 'notes.txt'))).rejects.toThrow(/Unsupported file type/);
    expect(JSON.stringify(target.getSections())).toContain('keep me');
  });
});
