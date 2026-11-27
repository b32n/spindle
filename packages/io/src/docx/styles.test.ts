import JSZip from 'jszip';
import { importDocx } from './reader';
import { StyleSheet } from './styles';
import type { DocBlock, DocInline } from '@b32nio/spindle-shared/export/docx';

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

async function docx(body: string, styles?: string, extra: Record<string, string> = {}): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types/>');
  zip.file('word/document.xml', `<?xml version="1.0"?><w:document ${NS}><w:body>${body}<w:sectPr/></w:body></w:document>`);
  if (styles !== undefined) zip.file('word/styles.xml', `<?xml version="1.0"?><w:styles ${NS}>${styles}</w:styles>`);
  for (const [n, x] of Object.entries(extra)) zip.file(n, x);
  return zip.generateAsync({ type: 'uint8array' });
}

const pStyle = (id: string, name: string, basedOn: string | undefined, props: { pPr?: string; rPr?: string; link?: string; def?: boolean } = {}) =>
  `<w:style w:type="paragraph"${props.def ? ' w:default="1"' : ''} w:styleId="${id}"><w:name w:val="${name}"/>${basedOn ? `<w:basedOn w:val="${basedOn}"/>` : ''}${props.link ? `<w:link w:val="${props.link}"/>` : ''}${props.pPr ? `<w:pPr>${props.pPr}</w:pPr>` : ''}${props.rPr ? `<w:rPr>${props.rPr}</w:rPr>` : ''}</w:style>`;
const cStyle = (id: string, name: string, basedOn: string | undefined, rPr: string, link?: string) =>
  `<w:style w:type="character" w:styleId="${id}"><w:name w:val="${name}"/>${basedOn ? `<w:basedOn w:val="${basedOn}"/>` : ''}${link ? `<w:link w:val="${link}"/>` : ''}<w:rPr>${rPr}</w:rPr></w:style>`;

const run = (text: string, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const p = (pPr: string, ...runs: string[]) => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${runs.join('')}</w:p>`;

const firstBlock = async (body: string, styles?: string): Promise<DocBlock> => (await importDocx(await docx(body, styles))).sections[0].blocks[0];
const runsOf = (b: DocBlock): DocInline[] => ('content' in b ? b.content : []);
const runProps = async (body: string, styles?: string) => {
  const props: Record<string, unknown> = { ...runsOf(await firstBlock(body, styles))[0] };
  delete props.type;
  delete props.text;
  delete props.namedStyleId; // asserted separately (named styles are kept, see "named styles" below)
  return props;
};

describe('style resolution: runs', () => {
  it('applies docDefaults to every run', async () => {
    const styles = '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Cambria"/><w:sz w:val="23"/><w:color w:val="112233"/></w:rPr></w:rPrDefault></w:docDefaults>';
    expect(await runProps(p('', run('x')), styles)).toEqual({ fontFamily: 'Cambria', fontSize: 11.5, color: '#112233' });
  });

  it('layers docDefaults, the default paragraph style, a named style and direct properties', async () => {
    const styles =
      '<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/><w:color w:val="111111"/></w:rPr></w:rPrDefault></w:docDefaults>' +
      pStyle('Normal', 'Normal', undefined, { def: true, rPr: '<w:rFonts w:ascii="Georgia"/>' }) +
      pStyle('Fancy', 'Fancy', 'Normal', { rPr: '<w:sz w:val="30"/><w:i/>' });
    expect(await runProps(p('', run('plain')), styles)).toEqual({ fontFamily: 'Georgia', fontSize: 11, color: '#111111' });
    expect(await runProps(p('<w:pStyle w:val="Fancy"/>', run('x')), styles)).toEqual({ fontFamily: 'Georgia', fontSize: 15, color: '#111111', italic: true });
    expect(await runProps(p('<w:pStyle w:val="Fancy"/>', run('x', '<w:sz w:val="40"/><w:i w:val="0"/><w:color w:val="auto"/>')), styles)).toEqual({
      fontFamily: 'Georgia',
      fontSize: 20,
    });
  });

  it('follows a basedOn chain several levels deep, nearest style winning', async () => {
    const styles =
      pStyle('A', 'A', undefined, { rPr: '<w:b/><w:sz w:val="20"/><w:color w:val="aa0000"/>' }) +
      pStyle('B', 'B', 'A', { rPr: '<w:sz w:val="24"/>' }) +
      pStyle('C', 'C', 'B', { rPr: '<w:color w:val="00bb00"/>' }) +
      pStyle('D', 'D', 'C', {});
    expect(await runProps(p('<w:pStyle w:val="D"/>', run('x')), styles)).toEqual({ bold: true, fontSize: 12, color: '#00bb00' });
  });

  it('survives a basedOn loop, a missing base and an unknown style id', async () => {
    const styles = pStyle('A', 'A', 'B', { rPr: '<w:b/>' }) + pStyle('B', 'B', 'A', { rPr: '<w:i/>' }) + pStyle('C', 'C', 'Gone', { rPr: '<w:u w:val="single"/>' });
    expect(await runProps(p('<w:pStyle w:val="A"/>', run('x')), styles)).toEqual({ bold: true, italic: true });
    expect(await runProps(p('<w:pStyle w:val="C"/>', run('x')), styles)).toEqual({ underline: true });
    expect(await runProps(p('<w:pStyle w:val="Nope"/>', run('x')), styles)).toEqual({});
  });

  it('applies a character style over the paragraph style, and direct formatting over both', async () => {
    const styles =
      pStyle('Body', 'Body', undefined, { rPr: '<w:sz w:val="22"/><w:color w:val="000000"/>' }) +
      cStyle('Loud', 'Loud', undefined, '<w:color w:val="ff0000"/><w:sz w:val="28"/><w:u w:val="single"/>');
    const rs = '<w:rStyle w:val="Loud"/>';
    expect(await runProps(p('<w:pStyle w:val="Body"/>', run('x', rs)), styles)).toEqual({ color: '#ff0000', fontSize: 14, underline: true });
    expect(await runProps(p('<w:pStyle w:val="Body"/>', run('x', rs + '<w:color w:val="0000ff"/><w:u w:val="none"/>')), styles)).toEqual({ color: '#0000ff', fontSize: 14 });
  });

  it('follows a basedOn chain of character styles', async () => {
    const styles = cStyle('Base', 'Base', undefined, '<w:i/><w:color w:val="123456"/>') + cStyle('Derived', 'Derived', 'Base', '<w:b/>');
    expect(await runProps(p('', run('x', '<w:rStyle w:val="Derived"/>')), styles)).toEqual({ italic: true, bold: true, color: '#123456' });
  });

  describe('toggle properties', () => {
    const styles =
      pStyle('Bold', 'Bold', undefined, { rPr: '<w:b/><w:i/>' }) +
      pStyle('BoldChild', 'BoldChild', 'Bold', { rPr: '<w:b/>' }) + // restating bold in a derived style does not toggle it off
      pStyle('Plain', 'Plain', 'Bold', { rPr: '<w:b w:val="0"/>' }) + // a derived style may switch it off
      cStyle('Emph', 'Emph', undefined, '<w:b/><w:i/><w:smallCaps/>') +
      cStyle('BoldChar', 'Bold Char', undefined, '<w:b/>', 'Bold');

    it('XORs the paragraph and character style states', async () => {
      const r = await runProps(p('<w:pStyle w:val="Bold"/>', run('x', '<w:rStyle w:val="Emph"/>')), styles);
      expect(r).toEqual({ smallCaps: true }); // bold and italic are on in both -> off; small caps only in the character style -> on
    });

    it('turns a property on when only the character style has it', async () => {
      expect(await runProps(p('', run('x', '<w:rStyle w:val="Emph"/>')), styles)).toEqual({ bold: true, italic: true, smallCaps: true });
    });

    it('lets direct formatting set the value outright, on or off', async () => {
      expect(await runProps(p('<w:pStyle w:val="Bold"/>', run('x', '<w:rStyle w:val="Emph"/><w:b/><w:i w:val="false"/>')), styles)).toEqual({ bold: true, smallCaps: true });
    });

    it('overrides inside one basedOn chain instead of toggling', async () => {
      expect(await runProps(p('<w:pStyle w:val="BoldChild"/>', run('x')), styles)).toEqual({ bold: true, italic: true });
      expect(await runProps(p('<w:pStyle w:val="Plain"/>', run('x')), styles)).toEqual({ italic: true });
    });

    it('does not toggle through the character style linked to the paragraph style', async () => {
      expect(await runProps(p('<w:pStyle w:val="Bold"/>', run('x', '<w:rStyle w:val="BoldChar"/>')), styles)).toEqual({ bold: true, italic: true });
    });

    it('reads w:val off/0/false as off and the bare element as on', async () => {
      for (const off of ['0', 'false', 'off']) {
        expect(await runProps(p('', run('x', `<w:b w:val="${off}"/>`)))).toEqual({});
      }
      expect(await runProps(p('', run('x', '<w:b w:val="1"/>')))).toEqual({ bold: true });
    });
  });

  it('resolves vertical alignment, strike, highlight and shading through styles', async () => {
    const styles = cStyle('Mark', 'Mark', undefined, '<w:vertAlign w:val="superscript"/><w:strike/><w:highlight w:val="yellow"/>') + cStyle('Shaded', 'Shaded', undefined, '<w:shd w:val="clear" w:fill="ddeeff"/>');
    expect(await runProps(p('', run('x', '<w:rStyle w:val="Mark"/>')), styles)).toEqual({ superscript: true, strikethrough: true, backgroundColor: '#ffff00' });
    expect(await runProps(p('', run('x', '<w:rStyle w:val="Shaded"/><w:vertAlign w:val="baseline"/>')), styles)).toEqual({ backgroundColor: '#ddeeff' });
    expect(await runProps(p('', run('x', '<w:rStyle w:val="Mark"/><w:highlight w:val="none"/>')), styles)).toEqual({ superscript: true, strikethrough: true });
  });

  it('resolves the Hyperlink character style on link runs', async () => {
    const styles = cStyle('Hyperlink', 'Hyperlink', undefined, '<w:color w:val="0563C1"/><w:u w:val="single"/>');
    const rels = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId5" Type="x" Target="https://example.com/" TargetMode="External"/></Relationships>';
    const result = await importDocx(
      await docx(p('', '<w:hyperlink r:id="rId5">' + run('site', '<w:rStyle w:val="Hyperlink"/>') + '</w:hyperlink>'), styles, { 'word/_rels/document.xml.rels': rels })
    );
    expect(runsOf(result.sections[0].blocks[0])[0]).toEqual({ type: 'link', text: 'site', href: 'https://example.com/', color: '#0563c1', underline: true });
  });

  it('keeps run formatting unstyled when there is no styles part or it is damaged', async () => {
    expect(await runProps(p('', run('x', '<w:b/>')))).toEqual({ bold: true });
    const damaged = await importDocx(await docx(p('', run('x', '<w:i/>')), undefined, { 'word/styles.xml': '<w:styles><w:style' }));
    expect(runsOf(damaged.sections[0].blocks[0])[0]).toMatchObject({ italic: true });
    expect(damaged.warnings).toEqual([]);
  });

  it('applies the paragraph style of runs in table cells and header/footer paragraphs', async () => {
    const styles = pStyle('Cell', 'Cell', undefined, { rPr: '<w:i/>' });
    const table = `<w:tbl><w:tr><w:tc>${p('<w:pStyle w:val="Cell"/>', run('in cell'))}</w:tc></w:tr></w:tbl>`;
    const block = await firstBlock(table, styles);
    expect(block.type === 'table' && block.rows[0].cells[0].content[0]).toMatchObject({ text: 'in cell', italic: true });
  });
});

describe('style resolution: paragraphs', () => {
  it('layers docDefaults, the style chain and direct properties', async () => {
    const styles =
      '<w:docDefaults><w:pPrDefault><w:pPr><w:spacing w:after="200"/></w:pPr></w:pPrDefault></w:docDefaults>' +
      pStyle('Normal', 'Normal', undefined, { def: true, pPr: '<w:jc w:val="both"/>' }) +
      pStyle('Indented', 'Indented', 'Normal', { pPr: '<w:ind w:left="720" w:firstLine="360"/><w:spacing w:before="120"/>' });
    expect(await firstBlock(p('', run('x')), styles)).toMatchObject({ type: 'paragraph', alignment: 'justify', spaceAfterPt: 10 });
    expect(await firstBlock(p('<w:pStyle w:val="Indented"/>', run('x')), styles)).toMatchObject({
      alignment: 'justify',
      spaceBeforePt: 6,
      spaceAfterPt: 10,
      leftIndentPt: 36,
      firstLineIndentPt: 18,
    });
    expect(await firstBlock(p('<w:pStyle w:val="Indented"/><w:jc w:val="center"/><w:ind w:left="0"/><w:spacing w:after="0"/>', run('x')), styles)).toMatchObject({
      alignment: 'center',
      leftIndentPt: 0,
      spaceAfterPt: 0,
      firstLineIndentPt: 18,
    });
  });

  it('merges spacing attributes individually', async () => {
    const styles = pStyle('S', 'S', undefined, { pPr: '<w:spacing w:before="240" w:after="240"/>' });
    expect(await firstBlock(p('<w:pStyle w:val="S"/><w:spacing w:after="0"/>', run('x')), styles)).toMatchObject({ spaceBeforePt: 12, spaceAfterPt: 0 });
  });

  describe('headings', () => {
    it('detects a Heading N style by name whatever its id', async () => {
      const styles = pStyle('Kop1', 'heading 1', undefined, {}) + pStyle('H3', 'Heading 3', undefined, {});
      expect(await firstBlock(p('<w:pStyle w:val="Kop1"/>', run('x')), styles)).toMatchObject({ type: 'heading', level: 1 });
      expect(await firstBlock(p('<w:pStyle w:val="H3"/>', run('x')), styles)).toMatchObject({ type: 'heading', level: 3 });
    });

    it('detects a heading by outline level, clamping to 6, and not for body text (9)', async () => {
      const styles = pStyle('Chapter', 'Chapter', undefined, { pPr: '<w:outlineLvl w:val="1"/>' }) + pStyle('Deep', 'Deep', undefined, { pPr: '<w:outlineLvl w:val="7"/>' }) + pStyle('Body', 'Body', undefined, { pPr: '<w:outlineLvl w:val="9"/>' });
      expect(await firstBlock(p('<w:pStyle w:val="Chapter"/>', run('x')), styles)).toMatchObject({ type: 'heading', level: 2 });
      expect(await firstBlock(p('<w:pStyle w:val="Deep"/>', run('x')), styles)).toMatchObject({ type: 'heading', level: 6 });
      expect((await firstBlock(p('<w:pStyle w:val="Body"/>', run('x')), styles)).type).toBe('paragraph');
    });

    it('lets a direct outline level on the paragraph decide', async () => {
      const styles = pStyle('Plain', 'Plain', undefined, {});
      expect(await firstBlock(p('<w:outlineLvl w:val="2"/>', run('x')), styles)).toMatchObject({ type: 'heading', level: 3 });
      const h = pStyle('Heading1', 'heading 1', undefined, { pPr: '<w:outlineLvl w:val="0"/>' });
      expect((await firstBlock(p('<w:pStyle w:val="Heading1"/><w:outlineLvl w:val="9"/>', run('x')), h)).type).toBe('paragraph');
    });

    it('inherits heading-ness from a Heading N ancestor unless a style overrides the outline level', async () => {
      const styles =
        pStyle('Heading2', 'heading 2', undefined, {}) +
        pStyle('Callout', 'Callout', 'Heading2', {}) +
        pStyle('TocHeading', 'TOC Heading', 'Heading2', { pPr: '<w:outlineLvl w:val="9"/>' });
      expect(await firstBlock(p('<w:pStyle w:val="Callout"/>', run('x')), styles)).toMatchObject({ type: 'heading', level: 2 });
      expect((await firstBlock(p('<w:pStyle w:val="TocHeading"/>', run('x')), styles)).type).toBe('paragraph');
    });

    it('resolves heading run formatting from the style', async () => {
      const styles = pStyle('Heading1', 'heading 1', undefined, { rPr: '<w:b/><w:sz w:val="32"/><w:color w:val="365F91"/>' });
      const block = await firstBlock(p('<w:pStyle w:val="Heading1"/>', run('Title')), styles);
      expect(block).toMatchObject({ type: 'heading', content: [{ text: 'Title', bold: true, fontSize: 16, color: '#365f91' }] });
    });
  });

  describe('lists', () => {
    const numbering = `<w:numbering ${NS}><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/></w:lvl><w:lvl w:ilvl="1"><w:numFmt w:val="decimal"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;
    const withNumbering = async (body: string, styles: string) => (await importDocx(await docx(body, styles, { 'word/numbering.xml': numbering }))).sections[0].blocks;

    it('makes a paragraph whose style carries numbering a list item', async () => {
      const styles = pStyle('ListBullet', 'List Bullet', undefined, { pPr: '<w:numPr><w:numId w:val="1"/></w:numPr>' });
      expect((await withNumbering(p('<w:pStyle w:val="ListBullet"/>', run('x')), styles))[0]).toMatchObject({ type: 'listItem', listType: 'bullet', level: 0 });
    });

    it('takes the level from the style and lets a direct numPr override it', async () => {
      const styles = pStyle('L', 'L', undefined, { pPr: '<w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr>' });
      const blocks = await withNumbering(p('<w:pStyle w:val="L"/>', run('a')) + p('<w:pStyle w:val="L"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>', run('b')), styles);
      expect(blocks[0]).toMatchObject({ type: 'listItem', listType: 'numbered', level: 1 });
      expect(blocks[1]).toMatchObject({ type: 'listItem', listType: 'bullet', level: 0 });
    });

    it('treats numId 0 as no list, in a style or directly', async () => {
      const styles = pStyle('L', 'L', undefined, { pPr: '<w:numPr><w:numId w:val="1"/></w:numPr>' });
      const blocks = await withNumbering(p('<w:pStyle w:val="L"/><w:numPr><w:numId w:val="0"/></w:numPr>', run('a')) + p('<w:numPr><w:numId w:val="0"/></w:numPr>', run('b')), styles);
      expect(blocks.map((b) => b.type)).toEqual(['paragraph', 'paragraph']);
    });

    it('keeps a heading that is numbered through its style a heading', async () => {
      const styles = pStyle('Heading1', 'heading 1', undefined, { pPr: '<w:numPr><w:numId w:val="1"/></w:numPr>' });
      expect((await withNumbering(p('<w:pStyle w:val="Heading1"/>', run('x')), styles))[0]).toMatchObject({ type: 'heading', level: 1 });
    });
  });

  it('applies the default paragraph style to paragraphs without a pStyle, and to headers', async () => {
    const styles = pStyle('Normal', 'Normal', undefined, { def: true, pPr: '<w:jc w:val="center"/>' }) + pStyle('Footer', 'footer', 'Normal', { pPr: '<w:jc w:val="right"/>' });
    const rels = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9" Type="x" Target="footer1.xml"/></Relationships>';
    const footer = `<w:ftr ${NS}>${p('<w:pStyle w:val="Footer"/>', run('foot'))}${p('', run('plain'))}</w:ftr>`;
    const bytes = await docx(p('', run('body')), styles, { 'word/_rels/document.xml.rels': rels, 'word/footer1.xml': footer });
    const zip = await JSZip.loadAsync(bytes);
    const doc = await zip.file('word/document.xml')!.async('string');
    zip.file('word/document.xml', doc.replace('<w:sectPr/>', '<w:sectPr><w:footerReference w:type="default" r:id="rId9"/></w:sectPr>'));
    const result = await importDocx(await zip.generateAsync({ type: 'uint8array' }));
    expect(result.sections[0].blocks[0]).toMatchObject({ alignment: 'center' });
    expect(result.sections[0].footer?.paragraphs.map((x) => x.alignment)).toEqual(['right', 'center']);
  });
});

describe('StyleSheet', () => {
  it('exposes style names and tolerates an absent part', () => {
    const sheet = new StyleSheet(`<w:styles ${NS}>${pStyle('X', 'My Style', undefined, {})}</w:styles>`);
    expect(sheet.name('X')).toBe('My Style');
    expect(sheet.name('Y')).toBeUndefined();
    expect(new StyleSheet(undefined).paragraph(undefined, {})).toEqual({ headingLevel: undefined });
  });
});
