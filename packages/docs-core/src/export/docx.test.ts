import JSZip from 'jszip';
import { exportToDocx, importFromDocx, inspectDocx } from './docx';
import { DocumentImpl } from '../document';
import type { DocumentData, InlineContent, ParagraphBlock } from '../types';

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

describe('importFromDocx', () => {
  it('rejects a zip that is not a .docx and leaves the open document untouched', async () => {
    const doc = docWith({
      sections: [{ id: 'sec1', pageConfig: PAGE_CONFIG, blocks: [{ id: 'old', type: 'paragraph', content: [{ type: 'text', text: 'Original' }] }] }],
    });
    const zip = new JSZip();
    zip.file('[Content_Types].xml', '<Types/>');
    zip.file('hello.txt', 'not a docx');
    await expect(importFromDocx(await zip.generateAsync({ type: 'uint8array' }), doc)).rejects.toThrow(/word\/document\.xml is missing/);
    expect(doc.getData().sections[0].blocks[0]).toMatchObject({ id: 'old' });
  });

  it('refuses a file that is not a package with an ImportError and leaves the open document untouched', async () => {
    const doc = docWith({
      sections: [{ id: 'sec1', pageConfig: PAGE_CONFIG, blocks: [{ id: 'old', type: 'paragraph', content: [{ type: 'text', text: 'Original' }] }] }],
    });
    await expect(importFromDocx(new TextEncoder().encode('plain text'), doc)).rejects.toMatchObject({ name: 'ImportError', code: 'not-a-package' });
    expect(doc.getData().sections[0].blocks[0]).toMatchObject({ id: 'old' });
  });

  it('replaces the document contents and preserves id/createdAt', async () => {
    const doc = docWith({
      id: 'doc',
      createdAt: '2020-01-01T00:00:00.000Z',
      sections: [
        {
          id: 'sec1',
          pageConfig: PAGE_CONFIG,
          blocks: [{ id: 'old', type: 'paragraph', content: [{ type: 'text', text: 'Original' }] }],
        },
      ],
    });
    const { result: exported } = await exportAndUnzip(
      docWith({ sections: [{ id: 'sec1', pageConfig: PAGE_CONFIG, blocks: [{ id: 'b1', type: 'paragraph', content: [{ type: 'text', text: 'New content' }] }] }] })
    );

    const { warnings } = await importFromDocx(exported.bytes, doc);
    expect(warnings).toEqual([]);

    const data = doc.getData();
    expect(data.id).toBe('doc');
    expect(data.createdAt).toBe('2020-01-01T00:00:00.000Z');
    expect(data.sections).toHaveLength(1);
    expect(data.sections[0].blocks).toHaveLength(1);
    const block = data.sections[0].blocks[0];
    expect(block.type).toBe('paragraph');
    if (block.type === 'paragraph') {
      expect(block.content).toEqual([{ type: 'text', text: 'New content' }]);
    }
  });

  it('round-trips run formatting via inline text style fields', async () => {
    const source = docWith({
      sections: [
        {
          id: 'sec1',
          pageConfig: PAGE_CONFIG,
          blocks: [
            {
              id: 'b1',
              type: 'paragraph',
              content: [{ type: 'text', text: 'Styled', bold: true, italic: true, color: '#ff0000' }],
            },
          ],
        },
      ],
    });
    const { result: exported } = await exportAndUnzip(source);

    const target = docWith({});
    await importFromDocx(exported.bytes, target);
    const block = target.getData().sections[0].blocks[0];
    expect(block.type).toBe('paragraph');
    if (block.type === 'paragraph') {
      expect(block.content).toEqual([{ type: 'text', text: 'Styled', bold: true, italic: true, color: '#ff0000' }]);
    }
  });

  it('round-trips a heading level and a bullet/numbered list', async () => {
    const source = docWith({
      sections: [
        {
          id: 'sec1',
          pageConfig: PAGE_CONFIG,
          blocks: [
            { id: 'b1', type: 'heading', level: 2, content: [{ type: 'text', text: 'Section' }] },
            { id: 'b2', type: 'list-item', listType: 'bullet', level: 0, content: [{ type: 'text', text: 'Bullet' }] },
            { id: 'b3', type: 'list-item', listType: 'numbered', level: 0, content: [{ type: 'text', text: 'Numbered' }] },
          ],
        },
      ],
    });
    const { result: exported } = await exportAndUnzip(source);

    const target = docWith({});
    await importFromDocx(exported.bytes, target);
    const blocks = target.getData().sections[0].blocks;
    expect(blocks[0]).toMatchObject({ type: 'heading', level: 2, content: [{ type: 'text', text: 'Section' }] });
    expect(blocks[1]).toMatchObject({ type: 'list-item', listType: 'bullet', level: 0, content: [{ type: 'text', text: 'Bullet' }] });
    expect(blocks[2]).toMatchObject({ type: 'list-item', listType: 'numbered', level: 0, content: [{ type: 'text', text: 'Numbered' }] });
  });

  it('round-trips a table with colspan', async () => {
    const source = docWith({
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
    const { result: exported } = await exportAndUnzip(source);

    const target = docWith({});
    await importFromDocx(exported.bytes, target);
    const block = target.getData().sections[0].blocks[0];
    expect(block.type).toBe('table');
    if (block.type === 'table') {
      expect(block.rows[0].cells[0]).toMatchObject({ content: [{ type: 'text', text: 'Wide' }], colspan: 2 });
      expect(block.rows[1].cells.map((c) => c.content)).toEqual([
        [{ type: 'text', text: 'Left' }],
        [{ type: 'text', text: 'Right' }],
      ]);
    }
  });

  it('round-trips table borders, shading, header row, widths, alignment and a table style through DOCX', async () => {
    const source = docWith({
      namedStyles: {
        Banded: { name: 'Banded', kind: 'table', table: { borders: { top: { style: 'single' } }, conditions: { firstRow: { backgroundColor: '#336699', text: { bold: true } } } } },
      },
      sections: [
        {
          id: 'sec1',
          pageConfig: PAGE_CONFIG,
          blocks: [
            {
              id: 'b1',
              type: 'table',
              colWidths: [96, 192],
              borders: { top: { style: 'double', widthPx: 2, color: '#ff0000' } },
              look: { firstRow: true },
              namedStyleId: 'Banded',
              rows: [
                {
                  id: 'r1',
                  repeatHeader: true,
                  cantSplit: true,
                  cells: [
                    { id: 'c1', content: [{ type: 'text', text: 'H' }], backgroundColor: '#ddeeff', verticalAlign: 'center', colwidth: [96], borders: { bottom: { style: 'dashed' } } },
                    { id: 'c2', content: [{ type: 'text', text: 'I' }], colwidth: [192] },
                  ],
                },
              ],
            },
          ],
        },
      ],
    });
    const { result: exported } = await exportAndUnzip(source);
    const target = docWith({});
    await importFromDocx(exported.bytes, target);
    const data = target.getData();
    const block = data.sections[0].blocks[0];
    if (block.type !== 'table') throw new Error('expected a table');
    expect(block).toMatchObject({ colWidths: [96, 192], namedStyleId: 'Banded', look: { firstRow: true }, borders: { top: { style: 'double', widthPx: 2, color: '#ff0000' } } });
    expect(block.rows[0]).toMatchObject({ repeatHeader: true, cantSplit: true });
    expect(block.rows[0].cells[0]).toMatchObject({ backgroundColor: '#ddeeff', verticalAlign: 'center', colwidth: [96], borders: { bottom: { style: 'dashed' } } });
    expect(data.namedStyles?.Banded).toMatchObject({ kind: 'table', table: { conditions: { firstRow: { backgroundColor: '#336699', text: { bold: true } } } } });
  });

  it('round-trips cells of several paragraphs, nested tables and the text direction through DOCX', async () => {
    const text = (id: string, t: string) => ({ id, type: 'paragraph' as const, content: [{ type: 'text' as const, text: t }] });
    const inner = { id: 'in', type: 'table' as const, rows: [{ id: 'ir', cells: [{ id: 'ic', content: [{ type: 'text' as const, text: 'inner' }] }] }] };
    const source = docWith({
      sections: [
        {
          id: 'sec1',
          pageConfig: PAGE_CONFIG,
          blocks: [
            {
              id: 'b1',
              type: 'table',
              rows: [{ id: 'r1', cells: [{ id: 'c1', content: [{ type: 'text', text: 'one' }], blocks: [text('p1', 'one'), text('p2', 'two'), inner, text('p3', '')], textDirection: 'tbRl' }] }],
            },
          ],
        },
      ],
    });
    const { result: exported } = await exportAndUnzip(source);
    const target = docWith({});
    await importFromDocx(exported.bytes, target);
    const block = target.getData().sections[0].blocks[0];
    if (block.type !== 'table') throw new Error('expected a table');
    const cell = block.rows[0].cells[0];
    expect(cell.textDirection).toBe('tbRl');
    expect(cell.blocks?.map((b) => b.type)).toEqual(['paragraph', 'paragraph', 'table', 'paragraph']);
    const nested = cell.blocks?.[2];
    if (nested?.type !== 'table') throw new Error('expected a nested table');
    expect(nested.rows[0].cells[0].content).toEqual([expect.objectContaining({ text: 'inner' })]);
  });

  it('round-trips header and footer variants and the section properties through DOCX', async () => {
    const hfp = (t: string) => ({ type: 'paragraph' as const, content: [{ type: 'text' as const, text: t }] });
    const source = docWith({
      sections: [
        {
          id: 'sec1',
          pageConfig: { ...PAGE_CONFIG, columns: { count: 2, spacing: 24, separator: true } },
          blocks: [{ id: 'b1', type: 'paragraph', content: [{ type: 'text', text: 'body' }] }],
          header: { blocks: [hfp('odd')], differentFirstPage: true, firstPageBlocks: [hfp('first')], differentOddEven: true, evenPageBlocks: [hfp('even')] },
          footer: { blocks: [hfp('foot')] },
          pageNumbering: { start: 7, format: 'upperRoman' },
          verticalAlign: 'center',
        },
      ],
    });
    const { result: exported } = await exportAndUnzip(source);
    const target = docWith({});
    await importFromDocx(exported.bytes, target);
    const section = target.getData().sections[0];
    const texts = (blocks: Array<{ content: Array<{ type: string }> }> | undefined) => blocks?.map((b) => b.content.map((c) => (c as { text?: string }).text).join(''));
    expect(texts(section.header?.blocks)).toEqual(['odd']);
    expect(texts(section.header?.firstPageBlocks)).toEqual(['first']);
    expect(texts(section.header?.evenPageBlocks)).toEqual(['even']);
    expect(section.header).toMatchObject({ differentFirstPage: true, differentOddEven: true });
    expect(texts(section.footer?.firstPageBlocks)).toEqual(['foot']);
    expect(section.pageConfig.columns).toEqual({ count: 2, spacing: 24, separator: true });
    expect(section.pageNumbering).toEqual({ start: 7, format: 'upperRoman' });
    expect(section.verticalAlign).toBe('center');
  });

  it('round-trips how each section starts (continuous, even and odd page); a new page is not written', async () => {
    const section = (id: string, breakType?: 'nextPage' | 'continuous' | 'evenPage' | 'oddPage') => ({
      id,
      pageConfig: PAGE_CONFIG,
      blocks: [{ id: `b_${id}`, type: 'paragraph' as const, content: [{ type: 'text' as const, text: id }] }],
      ...(breakType && { breakType }),
    });
    const source = docWith({ sections: [section('a', 'continuous'), section('b', 'continuous'), section('c', 'evenPage'), section('d', 'oddPage'), section('e', 'nextPage'), section('f')] });
    const { result: exported } = await exportAndUnzip(source);
    const target = docWith({});
    await importFromDocx(exported.bytes, target);
    // The first section has no previous one, but its value is kept as written.
    expect(target.getData().sections.map((s) => s.breakType)).toEqual(['continuous', 'continuous', 'evenPage', 'oddPage', undefined, undefined]);
  });

  it('round-trips a hyperlink via the text style pool when it carries formatting', async () => {
    const source = docWith({
      textStylePool: { style_1: { bold: true } },
      sections: [
        {
          id: 'sec1',
          pageConfig: PAGE_CONFIG,
          blocks: [
            { id: 'b1', type: 'paragraph', content: [{ type: 'link', text: 'Spindle', href: 'https://example.com', styleId: 'style_1' }] },
          ],
        },
      ],
    });
    const { result: exported } = await exportAndUnzip(source);

    const target = docWith({});
    await importFromDocx(exported.bytes, target);
    const block = target.getData().sections[0].blocks[0];
    expect(block.type).toBe('paragraph');
    if (block.type === 'paragraph') {
      const link = block.content[0];
      expect(link.type).toBe('link');
      if (link.type === 'link') {
        expect(link.text).toBe('Spindle');
        expect(link.href).toBe('https://example.com');
        expect(link.styleId).toBeDefined();
        expect(target.getData().textStylePool[link.styleId!]).toMatchObject({ bold: true });
      }
    }
  });

  it('round-trips a page-number dynamic field in the footer', async () => {
    const source = docWith({
      sections: [
        {
          id: 'sec1',
          pageConfig: PAGE_CONFIG,
          blocks: [{ id: 'b1', type: 'paragraph', content: [{ type: 'text', text: 'Body' }] }],
          footer: { blocks: [{ type: 'paragraph', content: [{ type: 'dynamicField', fieldType: 'pageNumber' }] }] },
        },
      ],
    });
    const { result: exported } = await exportAndUnzip(source);

    const target = docWith({});
    await importFromDocx(exported.bytes, target);
    const footer = target.getData().sections[0].footer;
    expect(footer?.blocks[0].content).toEqual([{ type: 'dynamicField', fieldType: 'pageNumber' }]);
  });

  it('round-trips bookmarks and internal links through the model', async () => {
    const content: InlineContent[] = [
      { type: 'bookmark', name: '_Toc1' },
      { type: 'text', text: 'Intro ' },
      { type: 'link', text: 'back', href: '#_Toc1' },
    ];
    const source = docWith({
      sections: [{ id: 'sec1', pageConfig: PAGE_CONFIG, blocks: [{ id: 'b1', type: 'paragraph', content }] }],
    });
    const { result: exported } = await exportAndUnzip(source);
    const target = docWith({});
    await importFromDocx(exported.bytes, target);
    expect(target.getData().sections[0].blocks[0]).toMatchObject({ content });
  });

  it('round-trips generic fields with their instruction and cached result through the model', async () => {
    const content: InlineContent[] = [
      { type: 'text', text: 'See ' },
      { type: 'field', instruction: 'REF _Ref1 \\h', kind: 'REF', result: 'Figure 1' },
      { type: 'field', instruction: 'SEQ Table', kind: 'SEQ', result: [{ type: 'text', text: '2', bold: true }], dirty: true },
    ];
    const source = docWith({
      sections: [{ id: 'sec1', pageConfig: PAGE_CONFIG, blocks: [{ id: 'b1', type: 'paragraph', content }] }],
    });
    const { result: exported } = await exportAndUnzip(source);
    const target = docWith({});
    await importFromDocx(exported.bytes, target);
    expect(target.getData().sections[0].blocks[0]).toMatchObject({ content });
  });

  it('round-trips a block-level image', async () => {
    const onePxPng =
      'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=';
    const source = docWith({
      sections: [
        {
          id: 'sec1',
          pageConfig: PAGE_CONFIG,
          blocks: [{ id: 'b1', type: 'image', src: onePxPng, width: 10, height: 10, alt: 'dot' }],
        },
      ],
    });
    const { result: exported } = await exportAndUnzip(source);

    const target = docWith({});
    const { warnings } = await importFromDocx(exported.bytes, target);
    expect(warnings).toEqual([]);
    const block = target.getData().sections[0].blocks[0];
    expect(block).toMatchObject({ type: 'image', src: onePxPng, width: 10, height: 10, alt: 'dot' });
  });
});

describe('inspectDocx detail', () => {
  async function bytesWithDropCap() {
    const exported = await exportToDocx(docWith({}));
    const zip = await JSZip.loadAsync(exported.bytes);
    const xml = await zip.file('word/document.xml')!.async('string');
    zip.file('word/document.xml', xml.replace('</w:body>', '<w:p><w:pPr><w:framePr w:dropCap="drop" w:lines="3"/></w:pPr><w:r><w:t>a</w:t></w:r></w:p></w:body>'));
    return zip.generateAsync({ type: 'uint8array' });
  }

  it('reports only confirmation-level losses by default', async () => {
    expect((await inspectDocx(await bytesWithDropCap())).items).toEqual([]);
  });

  it("reports formatting and structure losses with detail 'full'", async () => {
    const report = await inspectDocx(await bytesWithDropCap(), { detail: 'full' });
    expect(report.items.map((i) => i.id)).toEqual(['docx.dropCaps']);
  });
});

describe('tabs, line breaks and wrapped text through the editor API', () => {
  it('keeps tabs, line breaks, smart tags, tracked insertions and a content control on import, and writes tabs and breaks back', async () => {
    const source = docWith({
      sections: [{ id: 'sec1', pageConfig: PAGE_CONFIG, blocks: [{ id: 'b1', type: 'paragraph', content: [{ type: 'text', text: 'x' }] }] }],
    });
    const exported = await exportToDocx(source);
    const zip = await JSZip.loadAsync(exported.bytes);
    const xml = await zip.file('word/document.xml')!.async('string');
    zip.file(
      'word/document.xml',
      xml.replace(
        /<w:p>[\s\S]*?<\/w:p>/,
        '<w:p><w:r><w:t>Chapter</w:t><w:tab/><w:t>1</w:t><w:br/><w:t>next</w:t></w:r>' +
          '<w:smartTag w:element="place"><w:r><w:t> Paris</w:t></w:r></w:smartTag>' +
          '<w:ins w:id="1" w:author="A"><w:r><w:t> new</w:t></w:r></w:ins></w:p>' +
          '<w:sdt><w:sdtContent><w:p><w:r><w:t>control</w:t></w:r></w:p></w:sdtContent></w:sdt>'
      )
    );
    const bytes = await zip.generateAsync({ type: 'uint8array' });

    const target = docWith({});
    await importFromDocx(bytes, target);
    const blocks = target.getData().sections[0].blocks;
    expect(blocks).toHaveLength(2);
    const first = blocks[0];
    expect(first.type === 'paragraph' && first.content.map((c) => ('text' in c ? c.text : '')).join('')).toBe('Chapter\t1\nnext Paris new');
    const second = blocks[1];
    expect(second.type === 'content-control' && second.blocks).toMatchObject([{ type: 'paragraph', content: [{ type: 'text', text: 'control' }] }]);

    const { documentXml } = await exportAndUnzip(target);
    expect(documentXml).toContain('<w:tab/>');
    expect(documentXml).toContain('<w:br/>');
  });
});

describe('comments through the editor API', () => {
  const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"';

  async function withComments(): Promise<Uint8Array> {
    const zip = await JSZip.loadAsync((await exportToDocx(docWith({}))).bytes);
    zip.file(
      'word/document.xml',
      `<?xml version="1.0"?><w:document ${W}><w:body>` +
        '<w:p><w:r><w:t xml:space="preserve">See </w:t></w:r><w:commentRangeStart w:id="1"/><w:r><w:t>this part</w:t></w:r><w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r></w:p>' +
        '<w:sectPr/></w:body></w:document>'
    );
    zip.file(
      'word/comments.xml',
      `<w:comments ${W}>` +
        '<w:comment w:id="1" w:author="Alice" w:date="2026-01-08T14:00:00Z" w:initials="AA"><w:p w14:paraId="A1"><w:r><w:t>Root</w:t></w:r></w:p></w:comment>' +
        '<w:comment w:id="2" w:author="Bob" w:date="2026-01-09T14:00:00Z"><w:p w14:paraId="B2"><w:r><w:t>Reply</w:t></w:r></w:p></w:comment>' +
        '<w:comment w:id="3" w:author="Cy"><w:p><w:r><w:t>Floating</w:t></w:r></w:p></w:comment></w:comments>'
    );
    zip.file('word/commentsExtended.xml', `<w15:commentsEx ${W}><w15:commentEx w15:paraId="A1" w15:done="1"/><w15:commentEx w15:paraId="B2" w15:paraIdParent="A1" w15:done="0"/></w15:commentsEx>`);
    return zip.generateAsync({ type: 'uint8array' });
  }

  it('fills the comment store and tags the anchored run with the thread id', async () => {
    const doc = docWith({});
    const { warnings } = await importFromDocx(await withComments(), doc);
    expect(warnings).toEqual([]);

    const data = doc.getData();
    const block = data.sections[0].blocks[0];
    expect(block.type === 'paragraph' && block.content.map((c) => (c.type === 'text' ? [c.text, c.commentThreadId] : []))).toEqual([
      ['See ', undefined],
      ['this part', 'docx-thread-1'],
    ]);

    const threads = doc.getComments().getThreads();
    expect(threads.map((t) => t.id)).toEqual(['docx-thread-1', 'docx-thread-3']);
    expect(threads[0]).toMatchObject({
      anchor: { quote: 'this part' },
      status: 'resolved',
      createdAt: '2026-01-08T14:00:00Z',
      comments: [
        { id: 'docx-comment-1', authorId: 'Alice', authorName: 'Alice', authorInitials: 'AA', body: 'Root', createdAt: '2026-01-08T14:00:00Z' },
        { id: 'docx-comment-2', authorName: 'Bob', body: 'Reply', createdAt: '2026-01-09T14:00:00Z' },
      ],
    });
    expect(threads[1]).toMatchObject({ status: 'open', anchor: { quote: '' }, createdAt: '1970-01-01T00:00:00.000Z' });
    expect(data.threads).toHaveLength(2);
  });

  it('leaves no threads on a document without comments, replacing those of the previous content', async () => {
    const doc = docWith({});
    await importFromDocx(await withComments(), doc);
    expect(doc.getComments().getThreads()).toHaveLength(2);
    await importFromDocx((await exportToDocx(docWith({}))).bytes, doc);
    expect(doc.getComments().getThreads()).toEqual([]);
  });

  it('writes threads on export and re-imports them intact (author, initials, dates, replies, resolved, anchors)', async () => {
    const source = docWith({});
    await importFromDocx(await withComments(), source);
    // A thread made in the editor: anchored run plus a reply, never dated through a file.
    const store = source.getComments();
    const added = store.addThread({ quote: 'See ' }, 'From the editor', { id: 'u1', name: 'Eve' });
    store.addReply(added.id, 'A reply', { id: 'u2', name: 'Fay' });
    const data = source.getData();
    const block = data.sections[0].blocks[0];
    if (block.type !== 'paragraph') throw new Error('expected a paragraph');
    block.content[0] = { ...block.content[0], commentThreadId: added.id } as typeof block.content[0];
    source.setData(data);

    const { zip, documentXml } = await exportAndUnzip(source);
    expect(documentXml.match(/<w:commentRangeStart /g)).toHaveLength(2 + 2 + 1); // 2 for the file's thread, 2 for the new one, 1 orphan
    const comments = await zip.file('word/comments.xml')!.async('string');
    expect(comments).toContain('w:author="Alice"');
    expect(comments).toContain('w:author="Eve"');

    const target = docWith({});
    await importFromDocx((await exportToDocx(source)).bytes, target);
    const threads = target.getComments().getThreads();
    expect(threads).toHaveLength(3);
    const strip = (ts: typeof threads) =>
      ts.map((t) => ({ quote: t.anchor.quote, status: t.status, comments: t.comments.map((c) => [c.authorName, c.authorInitials, c.body]) }));
    expect(strip(threads)).toEqual(strip(source.getComments().getThreads()));
    const resolved = threads.find((t) => t.status === 'resolved');
    expect(resolved?.comments.map((c) => c.body)).toEqual(['Root', 'Reply']);
  });

  it('does not write the unknown-date placeholder as a date', async () => {
    const source = docWith({});
    await importFromDocx(await withComments(), source);
    const { zip } = await exportAndUnzip(source);
    const comments = await zip.file('word/comments.xml')!.async('string');
    expect(comments).toMatch(/<w:comment w:id="\d+" w:author="Cy">/);
  });
});

describe('paragraph formatting through the editor API', () => {
  const block = (extra: Partial<ParagraphBlock>): ParagraphBlock => ({ id: 'b1', type: 'paragraph', content: [{ type: 'text', text: 'x' }], ...extra });
  const roundTrip = async (blocks: DocumentData['sections'][number]['blocks'], pool: DocumentData['paragraphStylePool'] = {}) => {
    const source = docWith({ paragraphStylePool: pool, sections: [{ id: 'sec1', pageConfig: PAGE_CONFIG, blocks }] });
    const target = docWith({});
    await importFromDocx((await exportToDocx(source)).bytes, target);
    return target.getData().sections[0].blocks;
  };

  it('keeps the spacing and indent of a list item through export and import (inline, not through the pool)', async () => {
    const out = await roundTrip([
      { id: 'b1', type: 'list-item', listType: 'bullet', level: 0, content: [{ type: 'text', text: 'item' }], spaceAfter: 16, leftIndent: 48 },
    ]);
    expect(out[0]).toMatchObject({ type: 'list-item', spaceAfter: 16, leftIndent: 48 });
  });

  it('round-trips line spacing, indents, hanging indent and pagination flags on paragraphs and headings', async () => {
    const fmt = { lineSpacing: { type: 'multiple' as const, value: 1.15 }, rightIndent: 16, firstLineIndent: -24, leftIndent: 48, keepWithNext: true, keepTogether: true, pageBreakBefore: true, spaceAfter: 0 };
    const out = await roundTrip([block(fmt), { id: 'b2', type: 'heading', level: 2, content: [{ type: 'text', text: 'h' }], ...fmt }]);
    expect(out[0]).toMatchObject(fmt);
    expect(out[1]).toMatchObject({ type: 'heading', ...fmt });
  });

  it('writes atLeast and exact line heights in points and reads them back in pixels', async () => {
    const out = await roundTrip([
      block({ lineSpacing: { type: 'atLeast', value: 20 } }),
      block({ lineSpacing: { type: 'exactly', value: 24 } }),
      block({ lineSpacing: { type: 'double' } }),
      block({ lineSpacing: { type: 'multiple' } }), // no value: single
    ]);
    expect(out.map((b) => (b.type === 'paragraph' ? b.lineSpacing : undefined))).toEqual([
      { type: 'atLeast', value: 20 },
      { type: 'exactly', value: 24 },
      { type: 'double' },
      { type: 'single' },
    ]);
  });

  it('treats the indent of a paragraph as levels of 24 px on top of its left indent, and takes the pool style as a base', async () => {
    const source = docWith({
      paragraphStylePool: { ps1: { leftIndent: 12, rightIndent: 6, keepWithNext: true, lineSpacing: { type: 'onePointFive' } } },
      sections: [{ id: 'sec1', pageConfig: PAGE_CONFIG, blocks: [block({ indent: 2, styleId: 'ps1' }), block({ indent: 1, leftIndent: 10 })] }],
    });
    const { documentXml } = await exportAndUnzip(source);
    // (12 + 2 x 24) px = 45 pt = 900 twips, (10 + 24) px = 25.5 pt = 510 twips
    expect(documentXml).toContain('w:left="900"');
    expect(documentXml).toContain('w:right="90"');
    expect(documentXml).toContain('<w:keepNext/>');
    expect(documentXml).toContain('w:left="510"');
  });

  it('round-trips caps, hidden text, per-script fonts and numbered list starts', async () => {
    const out = await roundTrip([
      block({ content: [{ type: 'text', text: 'a', allCaps: true, hidden: true, fontFamily: 'Georgia', fontFamilyEastAsia: 'MS Mincho' }] }),
      { id: 'l1', type: 'list-item', listType: 'numbered', level: 0, start: 5, content: [{ type: 'text', text: 'five' }] },
      { id: 'l2', type: 'list-item', listType: 'numbered', level: 0, content: [{ type: 'text', text: 'six' }] },
      { id: 'l3', type: 'list-item', listType: 'numbered', level: 0, start: 1, content: [{ type: 'text', text: 'one again' }] },
    ]);
    expect(out[0].type === 'paragraph' && out[0].content[0]).toMatchObject({ allCaps: true, hidden: true, fontFamily: 'Georgia', fontFamilyEastAsia: 'MS Mincho' });
    expect(out.slice(1).map((b) => (b.type === 'list-item' ? b.start : 'x'))).toEqual([5, undefined, 1]);
  });
});

describe('named styles through DOCX', () => {
  const styles = {
    Callout: {
      name: 'Callout',
      kind: 'paragraph' as const,
      basedOn: 'Normal',
      next: 'Callout',
      text: { bold: true, color: '#c00000', fontSize: 12 },
      // px in the model: 48 px = 36 pt, hanging 24 px = 18 pt
      paragraph: { leftIndent: 48, firstLineIndent: -24, spaceBefore: 8, keepWithNext: true, lineSpacing: { type: 'multiple' as const, value: 1.25 } },
    },
    AccentWord: { name: 'Accent Word', kind: 'character' as const, text: { smallCaps: true, underline: true } },
  };
  const blocks: ParagraphBlock[] = [
    {
      id: 'b1',
      type: 'paragraph',
      namedStyleId: 'Callout',
      leftIndent: 48,
      content: [{ type: 'text', text: 'plain ' }, { type: 'text', text: 'accent', namedStyleId: 'AccentWord', smallCaps: true, underline: true }],
    },
  ];

  it('writes the styles and the references, and reads them back with the same ids, names and px values', async () => {
    const doc = docWith({ namedStyles: styles, sections: [{ id: 'sec1', pageConfig: PAGE_CONFIG, blocks }] });
    const { result, zip, documentXml } = await exportAndUnzip(doc);
    const stylesXml = await zip.file('word/styles.xml')!.async('string');
    expect(stylesXml).toContain('w:styleId="Callout"');
    expect(stylesXml).toContain('<w:name w:val="Accent Word"/>');
    expect(stylesXml).toContain('<w:ind w:left="720" w:hanging="360"/>');
    expect(documentXml).toContain('<w:pStyle w:val="Callout"/>');
    expect(documentXml).toContain('<w:rStyle w:val="AccentWord"/>');

    const target = new DocumentImpl('t', 'T');
    await importFromDocx(result.bytes, target);
    const data = target.getData();
    expect(data.namedStyles?.Callout).toMatchObject({
      name: 'Callout',
      kind: 'paragraph',
      basedOn: 'Normal',
      text: { bold: true, color: '#c00000', fontSize: 12 },
      paragraph: { leftIndent: 48, firstLineIndent: -24, spaceBefore: 8, keepWithNext: true, lineSpacing: { type: 'multiple', value: 1.25 } },
    });
    expect(data.namedStyles?.AccentWord).toMatchObject({ name: 'Accent Word', kind: 'character', text: { smallCaps: true, underline: true } });
    const block = data.sections[0].blocks[0] as ParagraphBlock;
    expect(block.namedStyleId).toBe('Callout');
    expect(block.content[1]).toMatchObject({ text: 'accent', namedStyleId: 'AccentWord' });
    expect(block.content[0]).not.toHaveProperty('namedStyleId');
  });

  it('a style removed from the document while still referenced is left out of the file', async () => {
    const doc = docWith({ namedStyles: styles, sections: [{ id: 'sec1', pageConfig: PAGE_CONFIG, blocks }] });
    doc.setNamedStyles(undefined);
    const { documentXml } = await exportAndUnzip(doc);
    expect(documentXml).not.toContain('w:val="Callout"');
    expect(documentXml).not.toContain('w:val="AccentWord"');
  });
});
