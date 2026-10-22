import { DocumentImpl } from '@b32nio/spindle-docs-core';
import JSZip from 'jszip';
import { DOCS_EXPORT_FORMATS, exportDocumentFile, importDocumentFile, inspectDocumentFile } from './file-io';

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

describe('docs import loss report', () => {
  async function docxWithTrackedChange(): Promise<File> {
    const out = await exportDocumentFile(makeDoc('Hello'), 'docx');
    const zip = await JSZip.loadAsync(out.data as Uint8Array);
    const xml = await zip.file('word/document.xml')!.async('string');
    zip.file('word/document.xml', xml.replace('</w:body>', '<w:p><w:ins w:id="9"><w:r><w:t>x</w:t></w:r></w:ins></w:p></w:body>'));
    return new File([new Uint8Array(await zip.generateAsync({ type: 'uint8array' }))], 'tracked.docx');
  }

  it('inspectDocumentFile reports losses without touching the document', async () => {
    const doc = makeDoc('Keep');
    const report = await inspectDocumentFile(await docxWithTrackedChange());
    expect(report.items.map((i) => i.id)).toEqual(['docx.trackedInsertions']);
    expect(JSON.stringify(doc.getData())).toContain('Keep');
  });

  it('importDocumentFile returns the report; a clean file reports none', async () => {
    const result = await importDocumentFile(makeDoc('x'), await docxWithTrackedChange());
    expect(result.loss.items.map((i) => i.id)).toEqual(['docx.trackedInsertions']);
    const clean = await exportDocumentFile(makeDoc('Hello'), 'docx');
    expect((await inspectDocumentFile(new File([new Uint8Array(clean.data as Uint8Array)], 'ok.docx'))).items).toEqual([]);
  });

  it('rejects non-docx files', async () => {
    await expect(inspectDocumentFile(new File(['x'], 'a.txt'))).rejects.toThrow(/Unsupported/);
  });
});
