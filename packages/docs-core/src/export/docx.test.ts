import JSZip from 'jszip';
import { exportToDocx } from './docx';
import { DocumentImpl } from '../document';
import type { DocumentData } from '../types';

const PAGE_CONFIG = {
  size: { w: 816, h: 1056 },
  margins: { top: 96, right: 96, bottom: 96, left: 96 },
  orientation: 'portrait' as const,
};

function docWith(data: Partial<DocumentData>): DocumentImpl {
  const doc = new DocumentImpl('doc', 'My Document');
  doc.setData({
    id: 'doc',
    title: 'My Document',
    defaultPageConfig: PAGE_CONFIG,
    textStylePool: {},
    paragraphStylePool: {},
    sections: [{ id: 'sec1', pageConfig: PAGE_CONFIG, blocks: [] }],
    ...data,
  });
  return doc;
}

async function exportAndUnzip(doc: DocumentImpl) {
  const result = await exportToDocx(doc);
  const zip = await JSZip.loadAsync(Buffer.from(result.bytes));
  const documentXml = await zip.file('word/document.xml')!.async('string');
  return { result, zip, documentXml };
}

describe('exportToDocx', () => {
  it('produces a well-formed OOXML zip containing the body text', async () => {
    const doc = docWith({
      sections: [
        {
          id: 'sec1',
          pageConfig: PAGE_CONFIG,
          blocks: [{ id: 'b1', type: 'paragraph', content: [{ type: 'text', text: 'Hello world' }] }],
        },
      ],
    });
    const { zip, documentXml } = await exportAndUnzip(doc);
    expect(zip.file('word/document.xml')).not.toBeNull();
    expect(documentXml).toContain('Hello world');
  });

  it('resolves a run styleId to concrete formatting from the text style pool', async () => {
    const doc = docWith({
      textStylePool: { style_1: { bold: true, color: '#ff0000' } },
      sections: [
        {
          id: 'sec1',
          pageConfig: PAGE_CONFIG,
          blocks: [{ id: 'b1', type: 'paragraph', content: [{ type: 'text', text: 'Styled', styleId: 'style_1' }] }],
        },
      ],
    });
    const { documentXml } = await exportAndUnzip(doc);
    expect(documentXml).toContain('<w:b/>');
    expect(documentXml).toMatch(/<w:color w:val="FF0000"/);
  });

  it('prefers inline run properties over the style pool when there is no styleId', async () => {
    const doc = docWith({
      sections: [
        {
          id: 'sec1',
          pageConfig: PAGE_CONFIG,
          blocks: [{ id: 'b1', type: 'paragraph', content: [{ type: 'text', text: 'Inline', italic: true }] }],
        },
      ],
    });
    const { documentXml } = await exportAndUnzip(doc);
    expect(documentXml).toContain('<w:i/>');
  });

  it('resolves paragraph alignment/spacing from the paragraph style pool, inline overriding it', async () => {
    const doc = docWith({
      paragraphStylePool: { pstyle_1: { alignment: 'center', spaceAfter: 16 } },
      sections: [
        {
          id: 'sec1',
          pageConfig: PAGE_CONFIG,
          blocks: [
            { id: 'b1', type: 'paragraph', styleId: 'pstyle_1', content: [{ type: 'text', text: 'Pooled' }] },
            { id: 'b2', type: 'paragraph', styleId: 'pstyle_1', alignment: 'right', content: [{ type: 'text', text: 'Overridden' }] },
          ],
        },
      ],
    });
    const { documentXml } = await exportAndUnzip(doc);
    expect(documentXml).toMatch(/<w:jc w:val="center"\/>[\s\S]*Pooled/);
    expect(documentXml).toMatch(/<w:jc w:val="right"\/>[\s\S]*Overridden/);
  });

  it('writes each heading level', async () => {
    const doc = docWith({
      sections: [
        {
          id: 'sec1',
          pageConfig: PAGE_CONFIG,
          blocks: [{ id: 'b1', type: 'heading', level: 3, content: [{ type: 'text', text: 'A Heading' }] }],
        },
      ],
    });
    const { documentXml } = await exportAndUnzip(doc);
    expect(documentXml).toContain('Heading3');
    expect(documentXml).toContain('A Heading');
  });

  it('writes a table with colspan and rowspan', async () => {
    const doc = docWith({
      sections: [
        {
          id: 'sec1',
          pageConfig: PAGE_CONFIG,
          blocks: [
            {
              id: 'b1',
              type: 'table',
              rows: [
                { id: 'r1', cells: [{ id: 'c1', content: [{ type: 'text', text: 'Wide' }], colspan: 2 }] },
                {
                  id: 'r2',
                  cells: [
                    { id: 'c2', content: [{ type: 'text', text: 'Left' }] },
                    { id: 'c3', content: [{ type: 'text', text: 'Right' }] },
                  ],
                },
              ],
            },
          ],
        },
      ],
    });
    const { documentXml } = await exportAndUnzip(doc);
    expect(documentXml).toContain('<w:tbl>');
    expect(documentXml).toMatch(/<w:gridSpan w:val="2"/);
    expect(documentXml).toContain('Wide');
    expect(documentXml).toContain('Right');
  });

  it('writes a bulleted list item and a numbered list item', async () => {
    const doc = docWith({
      sections: [
        {
          id: 'sec1',
          pageConfig: PAGE_CONFIG,
          blocks: [
            { id: 'b1', type: 'list-item', listType: 'bullet', level: 0, content: [{ type: 'text', text: 'Bullet' }] },
            { id: 'b2', type: 'list-item', listType: 'numbered', level: 0, content: [{ type: 'text', text: 'Numbered' }] },
          ],
        },
      ],
    });
    const { documentXml, zip } = await exportAndUnzip(doc);
    const numIds = [...documentXml.matchAll(/<w:numId w:val="(\d+)"\/>/g)].map((m) => m[1]);
    expect(numIds).toHaveLength(2);
    expect(numIds[0]).not.toBe(numIds[1]);
    const numberingXml = await zip.file('word/numbering.xml')!.async('string');
    expect(numberingXml).toMatch(/<w:numFmt w:val="bullet"\/>/);
    expect(numberingXml).toMatch(/<w:numFmt w:val="decimal"\/>/);
  });

  it('writes a horizontal rule and an explicit page break', async () => {
    const doc = docWith({
      sections: [
        {
          id: 'sec1',
          pageConfig: PAGE_CONFIG,
          blocks: [
            { id: 'b1', type: 'paragraph', content: [{ type: 'text', text: 'Before' }] },
            { id: 'b2', type: 'horizontal-rule' },
            { id: 'b3', type: 'page-break' },
            { id: 'b4', type: 'paragraph', content: [{ type: 'text', text: 'After' }] },
          ],
        },
      ],
    });
    const { documentXml } = await exportAndUnzip(doc);
    expect(documentXml).toContain('<w:bottom');
    expect(documentXml).toMatch(/<w:br w:type="page"/);
  });

  it('writes a page-number field in the footer and substitutes the doc title into a title field', async () => {
    const doc = docWith({
      sections: [
        {
          id: 'sec1',
          pageConfig: PAGE_CONFIG,
          blocks: [{ id: 'b1', type: 'paragraph', content: [{ type: 'text', text: 'Body' }] }],
          header: { blocks: [{ type: 'paragraph', content: [{ type: 'dynamicField', fieldType: 'title' }] }] },
          footer: {
            blocks: [
              { type: 'paragraph', content: [{ type: 'dynamicField', fieldType: 'pageNumber' }] },
            ],
          },
        },
      ],
    });
    const { zip } = await exportAndUnzip(doc);
    const headerFile = Object.keys(zip.files).find((f) => /^word\/header\d*\.xml$/.test(f));
    const footerFile = Object.keys(zip.files).find((f) => /^word\/footer\d*\.xml$/.test(f));
    const headerXml = await zip.file(headerFile!)!.async('string');
    const footerXml = await zip.file(footerFile!)!.async('string');
    expect(headerXml).toContain('My Document');
    expect(footerXml).toContain('PAGE');
  });

  it('sets page orientation and swapped width/height for a landscape section', async () => {
    const doc = docWith({
      sections: [
        {
          id: 'sec1',
          pageConfig: { ...PAGE_CONFIG, orientation: 'landscape' },
          blocks: [{ id: 'b1', type: 'paragraph', content: [] }],
        },
      ],
    });
    const { documentXml } = await exportAndUnzip(doc);
    expect(documentXml).toMatch(/<w:pgSz[^>]*w:orient="landscape"/);
  });

  it('reports no warnings for a document with no unembeddable images', async () => {
    const doc = docWith({
      sections: [
        {
          id: 'sec1',
          pageConfig: PAGE_CONFIG,
          blocks: [{ id: 'b1', type: 'paragraph', content: [{ type: 'text', text: 'Plain' }] }],
        },
      ],
    });
    const { result } = await exportAndUnzip(doc);
    expect(result.warnings).toEqual([]);
  });
});
