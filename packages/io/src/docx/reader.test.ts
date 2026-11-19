import JSZip from 'jszip';
import { exportDocx } from './writer';
import { importDocx } from './reader';
import type { DocSection, DocumentExportInput } from '@b32nio/spindle-shared/export/docx';

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

  it('reads an image wrapped in mc:AlternateContent (LibreOffice output)', async () => {
    const exported = await exportDocx(
      section([{ type: 'paragraph', content: [{ type: 'text', text: 'Caption: ' }, { type: 'image', src: ONE_PX_PNG, width: 5, height: 5 }] }])
    );
    const zip = await JSZip.loadAsync(exported.bytes);
    let xml = await zip.file('word/document.xml')!.async('string');
    xml = xml
      .replace(/<w:document /, '<w:document xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ')
      .replace(/(<w:drawing>[\s\S]*?<\/w:drawing>)/, '<mc:AlternateContent><mc:Choice Requires="wps">$1</mc:Choice><mc:Fallback><w:pict/></mc:Fallback></mc:AlternateContent>');
    expect(xml).toContain('<mc:AlternateContent>');
    zip.file('word/document.xml', xml);
    const result = await importDocx(await zip.generateAsync({ type: 'uint8array' }));
    expect(result.sections[0].blocks).toEqual([
      { type: 'paragraph', content: [{ type: 'text', text: 'Caption: ' }, { type: 'image', src: ONE_PX_PNG, width: 5, height: 5 }] },
    ]);
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

  describe('page breaks stored inside content paragraphs (#76)', () => {
    const BR = '<w:r><w:br w:type="page"/></w:r>';
    const run = (t: string) => `<w:r><w:t>${t}</w:t></w:r>`;

    async function importBody(paragraphsXml: string) {
      const exported = await exportDocx(section([{ type: 'paragraph', content: [{ type: 'text', text: 'x' }] }]));
      const zip = await JSZip.loadAsync(exported.bytes);
      let xml = await zip.file('word/document.xml')!.async('string');
      xml = xml.replace(/<w:p>[\s\S]*?<\/w:p>/, paragraphsXml);
      zip.file('word/document.xml', xml);
      return importDocx(await zip.generateAsync({ type: 'uint8array' }));
    }

    it('trailing break run (LibreOffice) becomes a pageBreak block after the paragraph', async () => {
      const result = await importBody(`<w:p>${run('One')}${BR}</w:p><w:p>${run('Two')}</w:p>`);
      expect(result.sections[0].blocks).toEqual([
        { type: 'paragraph', content: [{ type: 'text', text: 'One' }] },
        { type: 'pageBreak' },
        { type: 'paragraph', content: [{ type: 'text', text: 'Two' }] },
      ]);
    });

    it('a break after a block-level image keeps the image and adds the pageBreak', async () => {
      const exported = await exportDocx(section([{ type: 'image', src: ONE_PX_PNG, width: 5, height: 5 }]));
      const zip = await JSZip.loadAsync(exported.bytes);
      let xml = await zip.file('word/document.xml')!.async('string');
      xml = xml.replace(/<\/w:drawing><\/w:r>/, `</w:drawing></w:r>${BR}`);
      expect(xml).toContain('w:type="page"');
      zip.file('word/document.xml', xml);
      const result = await importDocx(await zip.generateAsync({ type: 'uint8array' }));
      expect(result.sections[0].blocks.map((b) => b.type)).toEqual(['image', 'pageBreak']);
    });

    it('a leading break run (page-break-before) puts the pageBreak first', async () => {
      const result = await importBody(`<w:p>${BR}${run('Two')}</w:p>`);
      expect(result.sections[0].blocks).toEqual([
        { type: 'pageBreak' },
        { type: 'paragraph', content: [{ type: 'text', text: 'Two' }] },
      ]);
    });

    it('an otherwise empty break paragraph is still a single pageBreak', async () => {
      const result = await importBody(`<w:p>${BR}</w:p>`);
      expect(result.sections[0].blocks).toEqual([{ type: 'pageBreak' }]);
    });
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

  it('round-trips landscape orientation and the on-page size', async () => {
    const result = await roundTrip(
      section([{ type: 'paragraph', content: [] }], { pageWidthPx: 1056, pageHeightPx: 816, orientation: 'landscape' })
    );
    expect(result.sections[0].orientation).toBe('landscape');
    expect(result.sections[0].pageWidthPx).toBe(1056);
    expect(result.sections[0].pageHeightPx).toBe(816);
  });

  describe('drawing kinds and warnings', () => {
    const NS = 'xmlns:wp="x" xmlns:a="x" xmlns:r="x" xmlns:pic="x" xmlns:wps="x" xmlns:o="x"';
    const drawing = (uri: string, inner: string, kind: 'inline' | 'anchor' = 'inline') =>
      `<w:p><w:r><w:drawing ${NS}><wp:${kind}><wp:extent cx="952500" cy="476250"/><wp:docPr id="1" name="Obj"/>` +
      `<a:graphic><a:graphicData uri="${uri}">${inner}</a:graphicData></a:graphic></wp:${kind}></w:drawing></w:r></w:p>`;
    const OOXML = 'http://schemas.openxmlformats.org/drawingml/2006';
    const picture = (blipAttrs: string, kind: 'inline' | 'anchor' = 'inline') =>
      drawing(`${OOXML}/picture`, `<pic:pic><pic:blipFill><a:blip ${blipAttrs}/></pic:blipFill></pic:pic>`, kind);

    async function importBody(paragraphsXml: string, extraRel?: string) {
      const exported = await exportDocx(section([{ type: 'image', src: ONE_PX_PNG, width: 5, height: 5 }]));
      const zip = await JSZip.loadAsync(exported.bytes);
      let xml = await zip.file('word/document.xml')!.async('string');
      xml = xml.replace(/<w:p>[\s\S]*?<\/w:p>/, paragraphsXml);
      zip.file('word/document.xml', xml);
      if (extraRel) {
        const relsPath = 'word/_rels/document.xml.rels';
        const rels = await zip.file(relsPath)!.async('string');
        zip.file(relsPath, rels.replace('</Relationships>', `${extraRel}</Relationships>`));
      }
      return importDocx(await zip.generateAsync({ type: 'uint8array' }));
    }

    it('imports a resolvable inline picture with no warning', async () => {
      const exported = await exportDocx(section([{ type: 'image', src: ONE_PX_PNG, width: 5, height: 5 }]));
      const result = await importDocx(exported.bytes);
      expect(result.sections[0].blocks[0]).toMatchObject({ type: 'image' });
      expect(result.warnings).toEqual([]);
    });

    it('keeps "could not be resolved" for a picture whose blip target is missing', async () => {
      const result = await importBody(picture('r:embed="rId999"'));
      expect(result.warnings).toEqual(['An embedded image could not be resolved and was skipped.']);
    });

    it('keeps "could not be resolved" for a picture with no blip reference at all', async () => {
      const result = await importBody(picture(''));
      expect(result.warnings).toEqual(['An embedded image could not be resolved and was skipped.']);
    });

    it('gives a linked (r:link) picture its own warning and never resolves it', async () => {
      const rel = '<Relationship Id="rId900" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/linked.png" TargetMode="External"/>';
      const result = await importBody(picture('r:link="rId900"'), rel);
      expect(result.warnings).toEqual(['A linked picture was skipped — images linked from outside the document are never fetched.']);
      expect(result.sections[0].blocks.some((b) => b.type === 'image')).toBe(false);
    });

    it('prefers r:embed over r:link when both are present', async () => {
      const exported = await exportDocx(section([{ type: 'image', src: ONE_PX_PNG, width: 5, height: 5 }]));
      const zip = await JSZip.loadAsync(exported.bytes);
      const xml = await zip.file('word/document.xml')!.async('string');
      const embed = xml.match(/r:embed="([^"]+)"/)![1];
      const result = await importBody(picture(`r:embed="${embed}" r:link="rId900"`));
      expect(result.warnings).toEqual([]);
      expect(result.sections[0].blocks[0]).toMatchObject({ type: 'image' });
    });

    it('keeps the floating-image warning for an anchored picture', async () => {
      const result = await importBody(picture('r:embed="rId1"', 'anchor'));
      expect(result.warnings).toEqual(['A floating (anchored) image was skipped — only inline images are supported.']);
    });

    it.each([
      ['a chart', `${OOXML}/chart`, '<c:chart xmlns:c="x" r:id="rId5"/>', 'A chart was not imported — charts are not supported yet.'],
      ['a chartex chart', 'http://schemas.microsoft.com/office/drawing/2014/chartex', '<cx:chart xmlns:cx="x"/>', 'A chart was not imported — charts are not supported yet.'],
      ['a SmartArt diagram', `${OOXML}/diagram`, '<dgm:relIds xmlns:dgm="x"/>', 'A SmartArt diagram was not imported — diagrams are not supported yet.'],
      ['a shape', 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape', '<wps:wsp/>', 'A shape or text box was not imported — shapes and text boxes are not supported yet.'],
      ['a shape group', 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup', '<wpg:wgp xmlns:wpg="x"/>', 'A shape or text box was not imported — shapes and text boxes are not supported yet.'],
      ['an OLE preview', 'urn:example:ole', '<o:OLEObject/>', 'An embedded object preview was not imported — embedded objects are not supported yet.'],
      ['an unknown drawing', 'urn:example:other', '<x:thing xmlns:x="x"/>', 'A drawing object of an unsupported kind was not imported.'],
    ])('reports %s as what it is, not as a missing image', async (_name, uri, inner, expected) => {
      const result = await importBody(drawing(uri, inner));
      expect(result.warnings).toEqual([expected]);
    });

    it('reports an anchored non-picture by its kind', async () => {
      const result = await importBody(drawing(`${OOXML}/chart`, '<c:chart xmlns:c="x"/>', 'anchor'));
      expect(result.warnings).toEqual(['A chart was not imported — charts are not supported yet.']);
    });

    it('classifies a drawing inside mc:AlternateContent', async () => {
      const inner = drawing('http://schemas.microsoft.com/office/word/2010/wordprocessingShape', '<wps:wsp/>', 'anchor')
        .replace('<w:p><w:r>', '<w:p><w:r><mc:AlternateContent xmlns:mc="x"><mc:Choice>')
        .replace('</w:r></w:p>', '</mc:Choice></mc:AlternateContent></w:r></w:p>');
      const result = await importBody(inner);
      expect(result.warnings).toEqual(['A shape or text box was not imported — shapes and text boxes are not supported yet.']);
    });
  });

  describe('page size from pgSz', () => {
    async function importPgSz(attrs: string) {
      const exported = await exportDocx(section([{ type: 'paragraph', content: [] }]));
      const zip = await JSZip.loadAsync(exported.bytes);
      let xml = await zip.file('word/document.xml')!.async('string');
      xml = xml.replace(/<w:pgSz[^>]*\/>/, `<w:pgSz ${attrs}/>`);
      zip.file('word/document.xml', xml);
      return (await importDocx(await zip.generateAsync({ type: 'uint8array' }))).sections[0];
    }

    it('reads landscape written with w < h (natural size plus orient)', async () => {
      const s = await importPgSz('w:w="12240" w:h="15840" w:orient="landscape"');
      expect(s).toMatchObject({ orientation: 'landscape', pageWidthPx: 1056, pageHeightPx: 816 });
    });

    it('reads landscape written with w > h', async () => {
      const s = await importPgSz('w:w="15840" w:h="12240" w:orient="landscape"');
      expect(s).toMatchObject({ orientation: 'landscape', pageWidthPx: 1056, pageHeightPx: 816 });
    });

    it('treats w > h without an orient attribute as landscape', async () => {
      const s = await importPgSz('w:w="15840" w:h="12240"');
      expect(s).toMatchObject({ orientation: 'landscape', pageWidthPx: 1056, pageHeightPx: 816 });
    });

    it('reads portrait unchanged', async () => {
      const s = await importPgSz('w:w="12240" w:h="15840" w:orient="portrait"');
      expect(s).toMatchObject({ orientation: 'portrait', pageWidthPx: 816, pageHeightPx: 1056 });
    });

    it('keeps a custom landscape size', async () => {
      const s = await importPgSz('w:w="7200" w:h="14400" w:orient="landscape"');
      expect(s).toMatchObject({ orientation: 'landscape', pageWidthPx: 960, pageHeightPx: 480 });
    });
  });

  it('round-trips multiple sections with independent orientation', async () => {
    const exported = await exportDocx({
      title: 'T',
      sections: [
        { ...BASE_SECTION, blocks: [{ type: 'paragraph', content: [{ type: 'text', text: 'One' }] }] },
        { ...BASE_SECTION, pageWidthPx: 1056, pageHeightPx: 816, orientation: 'landscape', blocks: [{ type: 'paragraph', content: [{ type: 'text', text: 'Two' }] }] },
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
