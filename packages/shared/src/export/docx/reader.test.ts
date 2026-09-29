import JSZip from 'jszip';
import { exportDocx } from './writer';
import { importDocx } from './reader';
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

const ONE_PX_PNG =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=';

async function roundTrip(input: DocumentExportInput) {
  const exported = await exportDocx(input);
  return importDocx(exported.bytes);
}

function section(blocks: DocSection['blocks'], extra: Partial<DocSection> = {}): DocumentExportInput {
  return { title: 'Test', sections: [{ ...BASE_SECTION, ...extra, blocks }] };
}

describe('importDocx', () => {
  it('rejects a non-zip byte stream', async () => {
    await expect(importDocx(new Uint8Array([0, 1, 2, 3]))).rejects.toThrow();
  });

  it('surfaces a warning instead of throwing for a zip with no word/document.xml', async () => {
    const zip = new JSZip();
    zip.file('hello.txt', 'not a docx');
    const bytes = await zip.generateAsync({ type: 'uint8array' });
    const result = await importDocx(bytes);
    expect(result.sections).toEqual([]);
    expect(result.warnings.length).toBeGreaterThan(0);
  });

  it('round-trips the title', async () => {
    const result = await roundTrip(section([{ type: 'paragraph', content: [{ type: 'text', text: 'Hi' }] }]));
    expect(result.title).toBe('Test');
  });

  it('round-trips plain text and run formatting', async () => {
    const result = await roundTrip(
      section([
        {
          type: 'paragraph',
          content: [
            { type: 'text', text: 'Bold', bold: true },
            { type: 'text', text: 'Italic', italic: true },
            { type: 'text', text: 'Underline', underline: true },
            { type: 'text', text: 'Strike', strikethrough: true },
            { type: 'text', text: 'Super', superscript: true },
            { type: 'text', text: 'Sub', subscript: true },
            { type: 'text', text: 'Colored', color: '#ff0000' },
            { type: 'text', text: 'Highlighted', backgroundColor: '#00ff00' },
            { type: 'text', text: 'Sized', fontSize: 18 },
            { type: 'text', text: 'Fonted', fontFamily: 'Georgia' },
          ],
        },
      ])
    );
    expect(result.sections[0].blocks).toEqual([
      {
        type: 'paragraph',
        content: [
          { type: 'text', text: 'Bold', bold: true },
          { type: 'text', text: 'Italic', italic: true },
          { type: 'text', text: 'Underline', underline: true },
          { type: 'text', text: 'Strike', strikethrough: true },
          { type: 'text', text: 'Super', superscript: true },
          { type: 'text', text: 'Sub', subscript: true },
          { type: 'text', text: 'Colored', color: '#ff0000' },
          { type: 'text', text: 'Highlighted', backgroundColor: '#00ff00' },
          { type: 'text', text: 'Sized', fontSize: 18 },
          { type: 'text', text: 'Fonted', fontFamily: 'Georgia' },
        ],
      },
    ]);
    expect(result.warnings).toEqual([]);
  });

  it('round-trips each heading level via the style-name map, not a hard-coded styleId', async () => {
    const blocks: DocSection['blocks'] = [1, 2, 3, 4, 5, 6].map((level) => ({
      type: 'heading',
      level: level as 1 | 2 | 3 | 4 | 5 | 6,
      content: [{ type: 'text', text: `H${level}` }],
    }));
    const result = await roundTrip(section(blocks));
    expect(result.sections[0].blocks).toEqual(
      [1, 2, 3, 4, 5, 6].map((level) => ({ type: 'heading', level, content: [{ type: 'text', text: `H${level}` }] }))
    );
  });

  it('round-trips bullet and numbered list items, including nesting level', async () => {
    const result = await roundTrip(
      section([
        { type: 'listItem', listType: 'bullet', level: 0, content: [{ type: 'text', text: 'Bullet' }] },
        { type: 'listItem', listType: 'bullet', level: 1, content: [{ type: 'text', text: 'Nested bullet' }] },
        { type: 'listItem', listType: 'numbered', level: 0, content: [{ type: 'text', text: 'Numbered' }] },
      ])
    );
    expect(result.sections[0].blocks).toEqual([
      { type: 'listItem', listType: 'bullet', level: 0, content: [{ type: 'text', text: 'Bullet' }] },
      { type: 'listItem', listType: 'bullet', level: 1, content: [{ type: 'text', text: 'Nested bullet' }] },
      { type: 'listItem', listType: 'numbered', level: 0, content: [{ type: 'text', text: 'Numbered' }] },
    ]);
  });

  it('round-trips a table with a colspan and a vertical rowspan merge', async () => {
    const result = await roundTrip(
      section([
        {
          type: 'table',
          rows: [
            { cells: [{ content: [{ type: 'text', text: 'A1' }], colspan: 2 }] },
            {
              cells: [
                { content: [{ type: 'text', text: 'B1' }], rowspan: 2 },
                { content: [{ type: 'text', text: 'B2' }] },
              ],
            },
            { cells: [{ content: [{ type: 'text', text: 'C2' }] }] },
          ],
        },
      ])
    );
    expect(result.sections[0].blocks).toEqual([
      {
        type: 'table',
        rows: [
          { cells: [{ content: [{ type: 'text', text: 'A1' }], colspan: 2 }] },
          {
            cells: [
              { content: [{ type: 'text', text: 'B1' }], rowspan: 2 },
              { content: [{ type: 'text', text: 'B2' }] },
            ],
          },
          { cells: [{ content: [{ type: 'text', text: 'C2' }] }] },
        ],
      },
    ]);
  });

  it('round-trips an external hyperlink', async () => {
    const result = await roundTrip(
      section([{ type: 'paragraph', content: [{ type: 'text', text: 'Before ' }, { type: 'link', text: 'Spindle', href: 'https://example.com' }] }])
    );
    expect(result.sections[0].blocks).toEqual([
      {
        type: 'paragraph',
        content: [
          { type: 'text', text: 'Before ' },
          { type: 'link', text: 'Spindle', href: 'https://example.com' },
        ],
      },
    ]);
  });

  it('round-trips a block-level image and an inline image mixed with text', async () => {
    const result = await roundTrip(
      section([
        { type: 'image', src: ONE_PX_PNG, width: 10, height: 10, alt: 'dot' },
        { type: 'paragraph', content: [{ type: 'text', text: 'Caption: ' }, { type: 'image', src: ONE_PX_PNG, width: 5, height: 5 }] },
      ])
    );
    expect(result.sections[0].blocks).toEqual([
      { type: 'image', src: ONE_PX_PNG, width: 10, height: 10, alt: 'dot' },
      { type: 'paragraph', content: [{ type: 'text', text: 'Caption: ' }, { type: 'image', src: ONE_PX_PNG, width: 5, height: 5 }] },
    ]);
    expect(result.warnings).toEqual([]);
  });

  it('round-trips a page break and a horizontal rule', async () => {
    const result = await roundTrip(
      section([
        { type: 'paragraph', content: [{ type: 'text', text: 'Before' }] },
        { type: 'pageBreak' },
        { type: 'horizontalRule' },
        { type: 'paragraph', content: [{ type: 'text', text: 'After' }] },
      ])
    );
    expect(result.sections[0].blocks).toEqual([
      { type: 'paragraph', content: [{ type: 'text', text: 'Before' }] },
      { type: 'pageBreak' },
      { type: 'horizontalRule' },
      { type: 'paragraph', content: [{ type: 'text', text: 'After' }] },
    ]);
  });

  it('round-trips paragraph alignment, spacing, and indent', async () => {
    const result = await roundTrip(
      section([
        {
          type: 'paragraph',
          content: [{ type: 'text', text: 'Styled' }],
          alignment: 'center',
          spaceBeforePt: 12,
          spaceAfterPt: 6,
          leftIndentPt: 24,
          firstLineIndentPt: 18,
        },
      ])
    );
    expect(result.sections[0].blocks).toEqual([
      {
        type: 'paragraph',
        content: [{ type: 'text', text: 'Styled' }],
        alignment: 'center',
        spaceBeforePt: 12,
        spaceAfterPt: 6,
        leftIndentPt: 24,
        firstLineIndentPt: 18,
      },
    ]);
  });

  it('round-trips page size and margins', async () => {
    const result = await roundTrip(
      section([{ type: 'paragraph', content: [] }], {
        pageWidthPx: 900,
        pageHeightPx: 1200,
        marginTopPx: 48,
        marginRightPx: 72,
        marginBottomPx: 48,
        marginLeftPx: 72,
      })
    );
    const s = result.sections[0];
    expect(s.pageWidthPx).toBe(900);
    expect(s.pageHeightPx).toBe(1200);
    expect(s.marginTopPx).toBe(48);
    expect(s.marginRightPx).toBe(72);
  });

  it('round-trips landscape orientation', async () => {
    // The docx library normalizes page-size numbers for landscape internally,
    // so this only asserts the orientation flag itself round-trips — matching
    // the scope of the writer's own "landscape" test.
    const result = await roundTrip(section([{ type: 'paragraph', content: [] }], { orientation: 'landscape' }));
    expect(result.sections[0].orientation).toBe('landscape');
  });

  it('round-trips multiple sections with independent orientation', async () => {
    const exported = await exportDocx({
      title: 'T',
      sections: [
        { ...BASE_SECTION, blocks: [{ type: 'paragraph', content: [{ type: 'text', text: 'One' }] }] },
        { ...BASE_SECTION, orientation: 'landscape', blocks: [{ type: 'paragraph', content: [{ type: 'text', text: 'Two' }] }] },
      ],
    });
    const result = await importDocx(exported.bytes);
    expect(result.sections).toHaveLength(2);
    expect(result.sections[0].orientation).toBe('portrait');
    expect(result.sections[1].orientation).toBe('landscape');
    expect(result.sections[0].blocks).toEqual([{ type: 'paragraph', content: [{ type: 'text', text: 'One' }] }]);
    expect(result.sections[1].blocks).toEqual([{ type: 'paragraph', content: [{ type: 'text', text: 'Two' }] }]);
  });

  it('round-trips header and footer content, including page-number/total-pages fields', async () => {
    const result = await roundTrip(
      section([{ type: 'paragraph', content: [{ type: 'text', text: 'Body' }] }], {
        header: { paragraphs: [{ content: [{ type: 'text', text: 'My Doc' }] }] },
        footer: { paragraphs: [{ content: [{ type: 'pageNumber' }, { type: 'text', text: ' of ' }, { type: 'totalPages' }] }] },
      })
    );
    expect(result.sections[0].header).toEqual({ paragraphs: [{ content: [{ type: 'text', text: 'My Doc' }] }] });
    expect(result.sections[0].footer).toEqual({
      paragraphs: [{ content: [{ type: 'pageNumber' }, { type: 'text', text: ' of ' }, { type: 'totalPages' }] }],
    });
  });

  it('reports no warnings for a document it fully understands', async () => {
    const result = await roundTrip(section([{ type: 'paragraph', content: [{ type: 'text', text: 'Plain' }] }]));
    expect(result.warnings).toEqual([]);
  });
});
