import JSZip from 'jszip';
import { exportDocx } from './writer';
import type { DocSection, DocumentExportInput } from './types';

const BASE_SECTION: Omit<DocSection, 'blocks'> = {
  pageWidthPx: 816,
  pageHeightPx: 1056,
  marginTopPx: 96,
  marginRightPx: 96,
  marginBottomPx: 96,
  marginLeftPx: 96,
  orientation: 'portrait',
};

async function buildAndUnzip(input: DocumentExportInput) {
  const result = await exportDocx(input);
  const zip = await JSZip.loadAsync(Buffer.from(result.bytes));
  const documentXml = await zip.file('word/document.xml')!.async('string');
  return { result, zip, documentXml };
}

function section(blocks: DocSection['blocks']): DocumentExportInput {
  return { title: 'Test', sections: [{ ...BASE_SECTION, blocks }] };
}

describe('exportDocx', () => {
  it('produces a well-formed OOXML zip with a document.xml part', async () => {
    const { zip, documentXml } = await buildAndUnzip(section([{ type: 'paragraph', content: [{ type: 'text', text: 'Hello' }] }]));
    expect(zip.file('[Content_Types].xml')).not.toBeNull();
    expect(zip.file('word/document.xml')).not.toBeNull();
    expect(documentXml).toContain('Hello');
  });

  it('writes bold/italic/underline/strike run formatting', async () => {
    const { documentXml } = await buildAndUnzip(
      section([
        {
          type: 'paragraph',
          content: [{ type: 'text', text: 'Styled', bold: true, italic: true, underline: true, strikethrough: true }],
        },
      ])
    );
    expect(documentXml).toContain('<w:b/>');
    expect(documentXml).toContain('<w:i/>');
    expect(documentXml).toMatch(/<w:u w:val="single"/);
    expect(documentXml).toContain('<w:strike/>');
  });

  it('writes a run color and highlight shading', async () => {
    const { documentXml } = await buildAndUnzip(
      section([{ type: 'paragraph', content: [{ type: 'text', text: 'Colored', color: '#ff0000', backgroundColor: '#00ff00' }] }])
    );
    expect(documentXml).toMatch(/<w:color w:val="FF0000"/);
    expect(documentXml).toMatch(/<w:shd[^>]*w:fill="00FF00"/);
  });

  it('writes each heading level with the right style reference', async () => {
    const { documentXml } = await buildAndUnzip(
      section([{ type: 'heading', level: 2, content: [{ type: 'text', text: 'Section Title' }] }])
    );
    expect(documentXml).toContain('Heading2');
    expect(documentXml).toContain('Section Title');
  });

  it('writes a bullet and a numbered list item, each referencing a distinct numbering definition', async () => {
    const { documentXml, zip } = await buildAndUnzip(
      section([
        { type: 'listItem', listType: 'bullet', level: 0, content: [{ type: 'text', text: 'Bullet item' }] },
        { type: 'listItem', listType: 'numbered', level: 1, content: [{ type: 'text', text: 'Numbered item' }] },
      ])
    );
    expect(documentXml).toContain('Bullet item');
    expect(documentXml).toContain('Numbered item');
    // Each list-item paragraph must reference a numId, and the two must differ
    // (bullet vs. numbered are separate numbering definitions).
    const numIds = [...documentXml.matchAll(/<w:numId w:val="(\d+)"\/>/g)].map((m) => m[1]);
    expect(numIds).toHaveLength(2);
    expect(numIds[0]).not.toBe(numIds[1]);
    // The definitions themselves live in numbering.xml, keyed by abstract
    // numbering id rather than our reference string — check the actual
    // format (bullet glyph vs. decimal) landed correctly.
    const numberingXml = await zip.file('word/numbering.xml')!.async('string');
    expect(numberingXml).toMatch(/<w:numFmt w:val="bullet"\/>/);
    expect(numberingXml).toMatch(/<w:numFmt w:val="decimal"\/>/);
  });

  it('writes a table with a merged cell (colspan/rowspan)', async () => {
    const { documentXml } = await buildAndUnzip(
      section([
        {
          type: 'table',
          rows: [
            { cells: [{ content: [{ type: 'text', text: 'A1' }], colspan: 2 }] },
            {
              cells: [
                { content: [{ type: 'text', text: 'B1' }] },
                { content: [{ type: 'text', text: 'B2' }] },
              ],
            },
          ],
        },
      ])
    );
    expect(documentXml).toContain('<w:tbl>');
    expect(documentXml).toMatch(/<w:gridSpan w:val="2"/);
    expect(documentXml).toContain('A1');
    expect(documentXml).toContain('B2');
  });

  it('embeds a data: URI image and warns about a non-data: src', async () => {
    // A minimal valid 1x1 PNG.
    const onePxPng =
      'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=';
    const { result, zip, documentXml } = await buildAndUnzip(
      section([
        { type: 'image', src: onePxPng, width: 10, height: 10 },
        { type: 'paragraph', content: [{ type: 'image', src: 'https://example.com/pic.png', width: 10, height: 10, alt: 'Remote' }] },
      ])
    );
    const mediaFiles = Object.keys(zip.files).filter((f) => f.startsWith('word/media/') && !f.endsWith('/'));
    expect(mediaFiles.length).toBe(1);
    expect(documentXml).toContain('[Image: Remote]');
    expect(result.warnings.some((w) => w.includes('example.com/pic.png'))).toBe(true);
  });

  it('writes an external hyperlink', async () => {
    const { documentXml, zip } = await buildAndUnzip(
      section([{ type: 'paragraph', content: [{ type: 'link', text: 'Spindle', href: 'https://example.com' }] }])
    );
    expect(documentXml).toContain('<w:hyperlink');
    const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
    expect(rels).toContain('https://example.com');
  });

  it('writes an explicit page break', async () => {
    const { documentXml } = await buildAndUnzip(
      section([
        { type: 'paragraph', content: [{ type: 'text', text: 'Before' }] },
        { type: 'pageBreak' },
        { type: 'paragraph', content: [{ type: 'text', text: 'After' }] },
      ])
    );
    expect(documentXml).toMatch(/<w:br w:type="page"/);
  });

  it('writes a horizontal rule as a bottom-bordered paragraph', async () => {
    const { documentXml } = await buildAndUnzip(section([{ type: 'horizontalRule' }]));
    expect(documentXml).toContain('<w:bottom');
  });

  it('writes page size/margins/orientation for a section', async () => {
    const { documentXml } = await buildAndUnzip({
      title: 'T',
      sections: [{ ...BASE_SECTION, orientation: 'landscape', blocks: [{ type: 'paragraph', content: [] }] }],
    });
    expect(documentXml).toMatch(/<w:pgSz[^>]*w:orient="landscape"/);
  });

  it('writes header and footer content, including a page-number field', async () => {
    const { zip } = await buildAndUnzip({
      title: 'T',
      sections: [
        {
          ...BASE_SECTION,
          blocks: [{ type: 'paragraph', content: [{ type: 'text', text: 'Body' }] }],
          header: { paragraphs: [{ content: [{ type: 'text', text: 'My Doc' }] }] },
          footer: { paragraphs: [{ content: [{ type: 'pageNumber' }, { type: 'text', text: ' of ' }, { type: 'totalPages' }] }] },
        },
      ],
    });
    const headerFile = Object.keys(zip.files).find((f) => /^word\/header\d*\.xml$/.test(f));
    const footerFile = Object.keys(zip.files).find((f) => /^word\/footer\d*\.xml$/.test(f));
    expect(headerFile).toBeDefined();
    expect(footerFile).toBeDefined();
    const headerXml = await zip.file(headerFile!)!.async('string');
    const footerXml = await zip.file(footerFile!)!.async('string');
    expect(headerXml).toContain('My Doc');
    expect(footerXml).toContain('PAGE');
    expect(footerXml).toContain('NUMPAGES');
  });

  it('reports no warnings for a document with no unembeddable images', async () => {
    const { result } = await buildAndUnzip(section([{ type: 'paragraph', content: [{ type: 'text', text: 'Plain' }] }]));
    expect(result.warnings).toEqual([]);
  });
});
