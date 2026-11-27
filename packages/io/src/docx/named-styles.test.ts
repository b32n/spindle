import { XMLValidator } from 'fast-xml-parser';
import JSZip from 'jszip';
import type { DocBlock, DocSection, DocumentExportInput } from '@b32nio/spindle-shared/export/docx';
import { exportDocx } from './writer';
import { importDocx } from './reader';

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

async function docx(body: string, styles: string): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types/>');
  zip.file('word/document.xml', `<?xml version="1.0"?><w:document ${NS}><w:body>${body}<w:sectPr/></w:body></w:document>`);
  zip.file('word/styles.xml', `<?xml version="1.0"?><w:styles ${NS}>${styles}</w:styles>`);
  return zip.generateAsync({ type: 'uint8array' });
}

const p = (style: string | undefined, text: string, rStyle?: string) =>
  `<w:p>${style ? `<w:pPr><w:pStyle w:val="${style}"/></w:pPr>` : ''}<w:r>${rStyle ? `<w:rPr><w:rStyle w:val="${rStyle}"/></w:rPr>` : ''}<w:t>${text}</w:t></w:r></w:p>`;

const STYLES =
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Callout"><w:name w:val="Callout"/><w:basedOn w:val="Normal"/><w:next w:val="Body"/><w:link w:val="CalloutChar"/>' +
  '<w:pPr><w:keepNext/><w:spacing w:before="120" w:after="240" w:line="360" w:lineRule="auto"/><w:ind w:left="720" w:hanging="360"/><w:jc w:val="both"/></w:pPr>' +
  '<w:rPr><w:b/><w:i w:val="0"/><w:color w:val="C00000"/><w:sz w:val="24"/><w:rFonts w:ascii="Georgia"/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Body"><w:name w:val="Body Text"/><w:basedOn w:val="Normal"/></w:style>' +
  '<w:style w:type="character" w:styleId="CalloutChar"><w:name w:val="Callout Char"/><w:link w:val="Callout"/><w:rPr><w:b/></w:rPr></w:style>' +
  '<w:style w:type="character" w:styleId="AccentWord"><w:name w:val="Accent Word"/><w:basedOn w:val="DefaultParagraphFont"/><w:rPr><w:smallCaps/><w:u w:val="single"/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Unused"><w:name w:val="Unused"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:sz w:val="40"/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="TocHead"><w:name w:val="TOC Heading"/><w:basedOn w:val="Heading1"/><w:pPr><w:outlineLvl w:val="9"/></w:pPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Quote2"><w:name w:val="Quote &amp; &quot;Aside&quot;"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="2"/></w:pPr></w:style>';

const readBlocks = async (body: string) => {
  const result = await importDocx(await docx(body, STYLES));
  return { result, blocks: result.sections[0].blocks };
};

describe('named styles: import', () => {
  it('keeps the style id on paragraphs and runs, but not for the default style or an undefined one', async () => {
    const { blocks } = await readBlocks(p('Callout', 'a') + p('Normal', 'b') + p('Missing', 'c') + p(undefined, 'd', 'AccentWord') + p(undefined, 'e', 'NoSuchChar'));
    expect(blocks.map((b) => (b as { namedStyleId?: string }).namedStyleId)).toEqual(['Callout', undefined, undefined, undefined, undefined]);
    const run = (b: DocBlock) => ('content' in b ? b.content[0] : undefined);
    expect(run(blocks[3])).toMatchObject({ namedStyleId: 'AccentWord', smallCaps: true, underline: true });
    expect(run(blocks[4])).not.toHaveProperty('namedStyleId');
    // The effective formatting is still applied.
    expect(blocks[0]).toMatchObject({ alignment: 'justify', spaceBeforePt: 6, leftIndentPt: 36, firstLineIndentPt: -18, keepWithNext: true });
    expect(run(blocks[0])).toMatchObject({ bold: true, color: '#c00000', fontSize: 12, fontFamily: 'Georgia' });
  });

  it('keeps headings in their style as well', async () => {
    const { blocks } = await readBlocks(p('Heading1', 'h'));
    expect(blocks[0]).toMatchObject({ type: 'heading', level: 1, namedStyleId: 'Heading1' });
  });

  it('returns the used styles with names, relations and their own properties, plus the styles they refer to', async () => {
    const { result } = await readBlocks(p('Callout', 'a') + p(undefined, 'd', 'AccentWord'));
    expect(Object.keys(result.namedStyles ?? {}).sort()).toEqual(['AccentWord', 'Body', 'Callout', 'CalloutChar', 'DefaultParagraphFont', 'Normal']);
    expect(result.namedStyles?.Callout).toEqual({
      name: 'Callout',
      kind: 'paragraph',
      basedOn: 'Normal',
      next: 'Body',
      link: 'CalloutChar',
      text: { bold: true, italic: false, color: '#c00000', fontSize: 12, fontFamily: 'Georgia' },
      paragraph: {
        alignment: 'justify',
        spaceBeforePt: 6,
        spaceAfterPt: 12,
        leftIndentPt: 36,
        firstLineIndentPt: -18,
        lineSpacing: 'onePointFive',
        keepWithNext: true,
      },
    });
    expect(result.namedStyles?.AccentWord).toMatchObject({ name: 'Accent Word', kind: 'character', basedOn: 'DefaultParagraphFont', text: { smallCaps: true, underline: true } });
    expect(result.namedStyles?.Normal).toMatchObject({ kind: 'paragraph', isDefault: true });
    expect(result.namedStyles).not.toHaveProperty('Unused');
  });

  it('records an outline level for a style that has one', async () => {
    const { result } = await readBlocks(p('Quote2', 'q'));
    expect(result.namedStyles?.Quote2).toMatchObject({ outlineLevel: 3 });
    expect(result.sections[0].blocks[0]).toMatchObject({ type: 'heading', level: 3, namedStyleId: 'Quote2' });
  });

  it('keeps a body-text outline level (it cancels the heading level of the style it is based on) and re-exports it', async () => {
    const { result, blocks } = await readBlocks(p('TocHead', 't'));
    expect(blocks[0]).toMatchObject({ type: 'paragraph', namedStyleId: 'TocHead' });
    expect(result.namedStyles?.TocHead).toMatchObject({ basedOn: 'Heading1', outlineLevel: 0 });
    const { bytes } = await exportDocx({ title: 't', sections: result.sections, namedStyles: result.namedStyles });
    const again = await importDocx(bytes);
    expect(again.sections[0].blocks[0]).toMatchObject({ type: 'paragraph', namedStyleId: 'TocHead' });
    expect(again.namedStyles?.TocHead).toEqual(result.namedStyles?.TocHead);
  });

  it('does not keep the implicit List Paragraph style on list items (the writer adds it again)', async () => {
    const numbering =
      '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>';
    const { result, blocks } = await readBlocks(
      `<w:p><w:pPr><w:pStyle w:val="ListParagraph"/>${numbering}</w:pPr><w:r><w:t>item</w:t></w:r></w:p>` + p('ListParagraph', 'plain')
    );
    expect(blocks[0]).toMatchObject({ type: 'listItem' });
    expect(blocks[0]).not.toHaveProperty('namedStyleId');
    // A paragraph that is not a list item keeps it.
    expect(blocks[1]).toMatchObject({ type: 'paragraph', namedStyleId: 'ListParagraph' });
    expect(result.namedStyles).toHaveProperty('ListParagraph');
  });

  it('does not keep the implicit Hyperlink style on a link, but keeps another character style there', async () => {
    const rels = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId5" Type="x" Target="https://example.com/" TargetMode="External"/></Relationships>';
    const zip = await JSZip.loadAsync(
      await docx(
        '<w:p><w:hyperlink r:id="rId5"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>a</w:t></w:r><w:r><w:rPr><w:rStyle w:val="AccentWord"/></w:rPr><w:t>b</w:t></w:r></w:hyperlink></w:p>',
        STYLES + '<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="0563C1"/></w:rPr></w:style>'
      )
    );
    zip.file('word/_rels/document.xml.rels', rels);
    const result = await importDocx(await zip.generateAsync({ type: 'uint8array' }));
    const content = (result.sections[0].blocks[0] as { content: object[] }).content;
    expect(content[0]).toMatchObject({ type: 'link', text: 'a', color: '#0563c1' });
    expect(content[0]).not.toHaveProperty('namedStyleId');
    expect(content[1]).toMatchObject({ type: 'link', text: 'b', namedStyleId: 'AccentWord' });
  });

  it('has no namedStyles when the document uses none', async () => {
    const { result } = await readBlocks(p(undefined, 'plain') + p('Normal', 'normal'));
    expect(result.namedStyles).toBeUndefined();
  });
});

const para = (text: string, namedStyleId?: string, extra: Partial<DocBlock> = {}): DocBlock =>
  ({ type: 'paragraph', content: [{ type: 'text', text }], ...(namedStyleId && { namedStyleId }), ...extra }) as DocBlock;

const section = (blocks: DocBlock[]): DocSection => ({
  pageWidthPx: 816, pageHeightPx: 1056, marginTopPx: 96, marginRightPx: 96, marginBottomPx: 96, marginLeftPx: 96, orientation: 'portrait', blocks,
});

const stylesXml = async (input: DocumentExportInput) => {
  const { bytes } = await exportDocx(input);
  const zip = await JSZip.loadAsync(bytes);
  return { bytes, xml: await zip.file('word/styles.xml')!.async('string'), doc: await zip.file('word/document.xml')!.async('string') };
};

describe('named styles: export', () => {
  const callout = {
    name: 'Callout "Box"',
    kind: 'paragraph' as const,
    basedOn: 'Normal',
    next: 'Callout',
    text: { bold: true, color: '#c00000', fontSize: 12, fontFamily: 'Georgia' },
    paragraph: { alignment: 'justify' as const, spaceBeforePt: 6, leftIndentPt: 36, firstLineIndentPt: -18, keepWithNext: true, lineSpacingCustom: { rule: 'multiple' as const, value: 1.5 } },
  };

  it('writes the styles into styles.xml and references them with pStyle and rStyle', async () => {
    const { xml, doc } = await stylesXml({
      title: 't',
      namedStyles: {
        Callout: callout,
        AccentWord: { name: 'Accent Word', kind: 'character', text: { smallCaps: true } },
      },
      sections: [
        section([
          para('boxed', 'Callout'),
          { type: 'paragraph', content: [{ type: 'text', text: 'accent', namedStyleId: 'AccentWord', smallCaps: true }] },
        ]),
      ],
    });
    expect(xml).toContain('w:styleId="Callout"');
    expect(xml).toContain('<w:name w:val="Callout &quot;Box&quot;"/>');
    expect(xml).toContain('<w:basedOn w:val="Normal"/>');
    expect(xml).toContain('<w:next w:val="Callout"/>');
    expect(xml).toMatch(/w:styleId="AccentWord"[\s\S]*?<w:name w:val="Accent Word"\/>/);
    expect(xml).toContain('<w:keepNext/>');
    expect(xml).toContain('<w:ind w:left="720" w:hanging="360"/>');
    expect(xml).toContain('<w:spacing w:before="120" w:line="360" w:lineRule="auto"/>');
    expect(doc).toContain('<w:pStyle w:val="Callout"/>');
    expect(doc).toContain('<w:rStyle w:val="AccentWord"/>');
    expect(XMLValidator.validate(xml)).toBe(true);
  });

  it('replaces a built-in style of the same id instead of duplicating it', async () => {
    const { xml } = await stylesXml({
      title: 't',
      namedStyles: {
        Heading1: { name: 'heading 1', kind: 'paragraph', basedOn: 'Normal', outlineLevel: 1, text: { fontSize: 20 } },
        Normal: { name: 'Normal', kind: 'paragraph', isDefault: true },
        Hyperlink: { name: 'Hyperlink', kind: 'character', text: { color: '#112233', underline: true } },
      },
      sections: [section([{ type: 'heading', level: 1, namedStyleId: 'Heading1', content: [{ type: 'text', text: 'T' }] }])],
    });
    const ids = [...xml.matchAll(/w:styleId="([^"]+)"/g)].map((m) => m[1]);
    expect(new Set(ids).size).toBe(ids.length);
    expect(xml).toContain('<w:name w:val="heading 1"/>');
    expect(xml).toContain('<w:outlineLvl w:val="0"/>');
    expect(xml).toMatch(/w:default="1" w:styleId="Normal"/);
    expect(xml).toContain('<w:color w:val="112233"/>');
    expect(xml).toContain('<w:sz w:val="40"/>');
  });

  it('writes a heading with a named style in that style, and one without in the built-in heading style', async () => {
    const { doc } = await stylesXml({
      title: 't',
      namedStyles: { Chapter: { name: 'Chapter', kind: 'paragraph', outlineLevel: 1 } },
      sections: [
        section([
          { type: 'heading', level: 1, namedStyleId: 'Chapter', content: [{ type: 'text', text: 'A' }] },
          { type: 'heading', level: 2, content: [{ type: 'text', text: 'B' }] },
        ]),
      ],
    });
    expect(doc).toContain('<w:pStyle w:val="Chapter"/>');
    expect(doc).toContain('<w:pStyle w:val="Heading2"/>');
  });

  it('leaves out references to styles the document does not define (a removed style), or of the wrong kind', async () => {
    const { xml, doc } = await stylesXml({
      title: 't',
      namedStyles: { Callout: callout },
      sections: [
        section([
          para('gone', 'Removed'),
          { type: 'paragraph', content: [{ type: 'text', text: 'x', namedStyleId: 'Callout' }, { type: 'text', text: 'y', namedStyleId: 'Removed' }] },
        ]),
      ],
    });
    expect(doc).not.toContain('w:val="Removed"');
    expect(doc).not.toContain('<w:rStyle w:val="Callout"/>'); // a paragraph style cannot style a run
    expect(xml).toContain('w:styleId="Callout"');
  });

  it('drops basedOn/next/link targets that are not defined', async () => {
    const { xml } = await stylesXml({
      title: 't',
      namedStyles: { A: { name: 'A', kind: 'paragraph', basedOn: 'A', next: 'Nope', link: 'Nope2' } },
      sections: [section([para('x', 'A')])],
    });
    const style = xml.match(/<w:style [^>]*w:styleId="A">[\s\S]*?<\/w:style>/)![0];
    expect(style).not.toMatch(/basedOn|next|link/);
  });

  it('writes the Hyperlink style on a link unless the link has a named character style of its own', async () => {
    const { doc } = await stylesXml({
      title: 't',
      namedStyles: { Fancy: { name: 'Fancy', kind: 'character' } },
      sections: [
        section([
          { type: 'paragraph', content: [{ type: 'link', text: 'a', href: 'https://a.example/' }, { type: 'link', text: 'b', href: 'https://b.example/', namedStyleId: 'Fancy' }] },
        ]),
      ],
    });
    expect(doc).toContain('<w:rStyle w:val="Hyperlink"/>');
    expect(doc).toContain('<w:rStyle w:val="Fancy"/>');
  });

  it('exports an unchanged document when there are no named styles', async () => {
    const { xml } = await stylesXml({ title: 't', sections: [section([para('x')])] });
    expect(xml).not.toContain('w:styleId="Callout"');
    expect(xml).toContain('w:styleId="Heading1"');
  });
});

describe('named styles: round trip', () => {
  it('re-imports the same ids, names and properties it exported', async () => {
    const first = await importDocx(
      await docx(p('Callout', 'a') + p(undefined, 'b', 'AccentWord') + p('Heading1', 'h') + p('Quote2', 'q'), STYLES)
    );
    const { bytes } = await exportDocx({ title: 't', sections: first.sections, namedStyles: first.namedStyles });
    const second = await importDocx(bytes);
    const ids = (r: typeof first) =>
      r.sections[0].blocks.map((b) => [('namedStyleId' in b && b.namedStyleId) || null, 'content' in b ? b.content.map((c) => ('namedStyleId' in c ? c.namedStyleId : null)) : null]);
    expect(ids(second)).toEqual(ids(first));
    for (const id of ['Callout', 'AccentWord', 'Heading1', 'Quote2']) {
      expect(second.namedStyles?.[id]).toEqual(first.namedStyles?.[id]);
    }
    expect(second.sections[0].blocks.map((b) => b.type)).toEqual(['paragraph', 'paragraph', 'heading', 'heading']);
  });
});
