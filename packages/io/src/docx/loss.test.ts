import JSZip from 'jszip';
import { detectDocxLoss } from './loss';
import { exportDocx } from './writer';

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

/** A minimal docx: just the parts the detector reads. */
async function docx(body: string, parts: Record<string, string> = {}): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('word/document.xml', `<?xml version="1.0"?><w:document ${NS}><w:body>${body}</w:body></w:document>`);
  for (const [name, xml] of Object.entries(parts)) zip.file(name, xml);
  return zip.generateAsync({ type: 'uint8array' });
}

async function item(bytes: Uint8Array, id: string) {
  return (await detectDocxLoss(bytes, { detail: 'full' })).items.find((i) => i.id === id);
}

const p = (text: string, pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}<w:r><w:t>${text}</w:t></w:r></w:p>`;
const styles = (...defs: [id: string, name: string][]) =>
  `<w:styles ${NS}>${defs.map(([id, name]) => `<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${name}"/></w:style>`).join('')}</w:styles>`;

describe('detectDocxLoss: XML-based detectors', () => {
  it('reports nothing for plain paragraphs, headings and Normal', async () => {
    const bytes = await docx(p('a') + p('b', '<w:pStyle w:val="Heading1"/>') + p('c', '<w:pStyle w:val="Normal"/>'), {
      'word/styles.xml': styles(['Heading1', 'heading 1'], ['Normal', 'Normal']),
    });
    expect((await detectDocxLoss(bytes, { detail: 'full' })).items).toEqual([]);
  });

  it('does not flag a document written by our own exporter', async () => {
    const out = await exportDocx({
      title: 'T',
      sections: [
        {
          pageWidthPx: 816,
          pageHeightPx: 1056,
          marginTopPx: 96,
          marginRightPx: 96,
          marginBottomPx: 96,
          marginLeftPx: 96,
          orientation: 'portrait',
          blocks: [
            { type: 'heading', level: 1, content: [{ type: 'text', text: 'H' }] },
            { type: 'listItem', listType: 'bullet', level: 0, content: [{ type: 'text', text: 'a' }] },
            { type: 'listItem', listType: 'numbered', level: 0, content: [{ type: 'text', text: 'b' }] },
          ],
        },
      ],
    });
    const ids = (await detectDocxLoss(out.bytes, { detail: 'full' })).items.map((i) => i.id);
    expect(ids).not.toContain('docx.listFormatting');
    expect(ids).not.toContain('docx.namedStyles');
    expect(ids).not.toContain('docx.sections');
    expect(ids).not.toContain('docx.tabsAndBreaks');
  });

  describe('docx.sections', () => {
    const sect = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>';
    it('counts sections beyond the first', async () => {
      const bytes = await docx(`${p('a', sect)}${p('b', sect)}${p('c')}${sect}`);
      expect(await item(bytes, 'docx.sections')).toMatchObject({ count: 2 });
    });
    it('is silent for a single section', async () => {
      expect(await item(await docx(p('a') + sect), 'docx.sections')).toBeUndefined();
    });
    it('ignores sectPrChange and sectPr outside the body flow', async () => {
      const bytes = await docx(
        `${p('a')}<w:tbl><w:tr><w:tc><w:p><w:pPr>${sect}</w:pPr></w:p></w:tc></w:tr></w:tbl>` +
          `${sect.replace('</w:sectPr>', '<w:sectPrChange w:id="1"><w:sectPr/></w:sectPrChange></w:sectPr>')}`
      );
      expect(await item(bytes, 'docx.sections')).toBeUndefined();
    });
  });

  describe('docx.namedStyles', () => {
    const parts = { 'word/styles.xml': styles(['Heading2', 'heading 2'], ['Callout', 'Callout'], ['ListParagraph', 'List Paragraph']) };
    it('counts paragraph and character style uses with names in where, skipping headings, Normal and List Paragraph', async () => {
      const bytes = await docx(
        p('a', '<w:pStyle w:val="Callout"/>') +
          p('b', '<w:pStyle w:val="Callout"/>') +
          p('c', '<w:pStyle w:val="Heading2"/>') +
          p('e', '<w:pStyle w:val="ListParagraph"/>') +
          '<w:p><w:r><w:rPr><w:rStyle w:val="Emphasis2"/></w:rPr><w:t>d</w:t></w:r></w:p>',
        parts
      );
      expect(await item(bytes, 'docx.namedStyles')).toMatchObject({ count: 3, where: ['Callout ×2', 'Emphasis2 ×1'] });
    });
    it('falls back to the id for a style missing from styles.xml', async () => {
      const bytes = await docx(p('a', '<w:pStyle w:val="NoSuchStyle"/>'));
      expect(await item(bytes, 'docx.namedStyles')).toMatchObject({ count: 1, where: ['NoSuchStyle ×1'] });
    });
    it('ignores a style recorded in a pPrChange', async () => {
      const bytes = await docx(p('a', '<w:pPrChange w:id="1"><w:pPr><w:pStyle w:val="Callout"/></w:pPr></w:pPrChange>'), parts);
      expect(await item(bytes, 'docx.namedStyles')).toBeUndefined();
    });
  });

  describe('docx.tableFormatting', () => {
    const tbl = (tblPr: string, tcPr = '', trPr = '') =>
      `<w:tbl><w:tblPr>${tblPr}</w:tblPr><w:tr>${trPr ? `<w:trPr>${trPr}</w:trPr>` : ''}<w:tc><w:tcPr>${tcPr}</w:tcPr>${p('x')}</w:tc></w:tr></w:tbl>`;
    it('is silent for a table with no formatting (auto width, plain shading)', async () => {
      const bytes = await docx(tbl('<w:tblW w:w="0" w:type="auto"/>', '<w:shd w:val="clear" w:fill="auto"/>'));
      expect(await item(bytes, 'docx.tableFormatting')).toBeUndefined();
    });
    it.each([
      ['borders', '<w:tblBorders><w:top w:val="single"/></w:tblBorders>', '', ''],
      ['borders', '', '<w:tcBorders><w:top w:val="single"/></w:tcBorders>', ''],
      ['shading', '', '<w:shd w:val="clear" w:fill="FFEE00"/>', ''],
      ['widths', '<w:tblW w:w="5000" w:type="pct"/>', '', ''],
      ['widths', '', '<w:tcW w:w="2400" w:type="dxa"/>', ''],
      ['repeating header rows', '', '', '<w:tblHeader/>'],
      ['vertical alignment', '', '<w:vAlign w:val="center"/>', ''],
      ['table style', '<w:tblStyle w:val="GridTable4"/>', '', ''],
    ])('detects %s', async (kind, tblPr, tcPr, trPr) => {
      const found = await item(await docx(tbl(tblPr, tcPr, trPr)), 'docx.tableFormatting');
      expect(found).toMatchObject({ count: 1, where: [`${kind} ×1`] });
    });
    it('counts each table once however many kinds it has', async () => {
      const bytes = await docx(tbl('<w:tblStyle w:val="X"/>', '<w:vAlign w:val="bottom"/>') + tbl('<w:tblBorders/>'));
      expect(await item(bytes, 'docx.tableFormatting')).toMatchObject({ count: 2 });
    });
    it('does not treat w:tblHeader w:val="0" as a header row', async () => {
      expect(await item(await docx(tbl('', '', '<w:tblHeader w:val="0"/>')), 'docx.tableFormatting')).toBeUndefined();
    });
  });

  describe('docx.tableStructure', () => {
    it('counts nested tables and cells with several non-empty paragraphs', async () => {
      const bytes = await docx(
        `<w:tbl><w:tr>` +
          `<w:tc>${p('one')}${p('two')}${p('')}</w:tc>` +
          `<w:tc>${p('solo')}<w:tbl><w:tr><w:tc>${p('inner')}</w:tc></w:tr></w:tbl></w:tc>` +
          `</w:tr></w:tbl>`
      );
      expect(await item(bytes, 'docx.tableStructure')).toMatchObject({
        count: 2,
        where: ['multi-paragraph cells ×1', 'nested tables ×1'],
      });
    });
    it('ignores empty trailing paragraphs in a cell', async () => {
      const bytes = await docx(`<w:tbl><w:tr><w:tc>${p('one')}<w:p/><w:p/></w:tc></w:tr></w:tbl>`);
      expect(await item(bytes, 'docx.tableStructure')).toBeUndefined();
    });
  });

  describe('docx.paragraphFormatting', () => {
    it.each([
      ['line spacing', '<w:spacing w:line="360" w:lineRule="auto"/>'],
      ['right/hanging indent', '<w:ind w:right="720"/>'],
      ['right/hanging indent', '<w:ind w:left="720" w:hanging="360"/>'],
      ['keep with next/lines', '<w:keepNext/>'],
      ['keep with next/lines', '<w:keepLines/>'],
      ['page break before', '<w:pageBreakBefore/>'],
    ])('detects %s', async (kind, pPr) => {
      expect(await item(await docx(p('a', pPr)), 'docx.paragraphFormatting')).toMatchObject({ count: 1, where: [`${kind} ×1`] });
    });
    it('detects hidden runs but not hidden paragraph marks', async () => {
      const bytes = await docx(
        '<w:p><w:r><w:rPr><w:vanish/></w:rPr><w:t>secret</w:t></w:r></w:p>' + p('mark', '<w:rPr><w:vanish/></w:rPr>')
      );
      expect(await item(bytes, 'docx.paragraphFormatting')).toMatchObject({ count: 1, where: ['hidden text ×1'] });
    });
    it('is silent for spacing before/after, left/first-line indents, switched-off flags and revisions', async () => {
      const bytes = await docx(
        p('a', '<w:spacing w:before="120" w:after="120"/><w:ind w:left="720" w:firstLine="360"/><w:keepNext w:val="0"/>') +
          p('b', '<w:pPrChange w:id="1"><w:pPr><w:keepNext/></w:pPr></w:pPrChange>')
      );
      expect(await item(bytes, 'docx.paragraphFormatting')).toBeUndefined();
    });
  });

  describe('docx.listFormatting', () => {
    const numbering =
      `<w:numbering ${NS}>` +
      `<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/></w:lvl></w:abstractNum>` +
      `<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="lowerRoman"/></w:lvl></w:abstractNum>` +
      `<w:abstractNum w:abstractNumId="2"><w:lvl w:ilvl="0"><w:start w:val="5"/><w:numFmt w:val="decimal"/></w:lvl></w:abstractNum>` +
      `<w:abstractNum w:abstractNumId="3"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/></w:lvl></w:abstractNum>` +
      `<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>` +
      `<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>` +
      `<w:num w:numId="3"><w:abstractNumId w:val="2"/></w:num>` +
      `<w:num w:numId="4"><w:abstractNumId w:val="3"/></w:num>` +
      `<w:num w:numId="5"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="3"/></w:lvlOverride></w:num>` +
      `<w:num w:numId="6"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>` +
      `</w:numbering>`;
    const li = (numId: string) => p('i', `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr>`);
    it('is silent for decimal lists, bullets, numId 0 and a restart at 1', async () => {
      const bytes = await docx(li('1') + li('4') + li('0') + li('6'), { 'word/numbering.xml': numbering });
      expect(await item(bytes, 'docx.listFormatting')).toBeUndefined();
    });
    it('detects other formats, custom start values and start overrides', async () => {
      const bytes = await docx(li('2') + li('3') + li('5'), { 'word/numbering.xml': numbering });
      expect(await item(bytes, 'docx.listFormatting')).toMatchObject({ count: 3, where: ['custom start ×2', 'lowerRoman ×1'] });
    });
  });

  describe('docx.tabsAndBreaks', () => {
    it('counts tabs, line/column breaks and hyphen characters, not tab stops or page breaks', async () => {
      const bytes = await docx(
        '<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="720"/></w:tabs></w:pPr>' +
          '<w:r><w:t>a</w:t><w:tab/><w:t>b</w:t><w:br/><w:br w:type="textWrapping"/><w:br w:type="column"/><w:br w:type="page"/>' +
          '<w:noBreakHyphen/><w:softHyphen/></w:r></w:p>'
      );
      expect(await item(bytes, 'docx.tabsAndBreaks')).toMatchObject({
        count: 6,
        where: ['line break ×2', 'column break ×1', 'non-breaking hyphen ×1', 'soft hyphen ×1', 'tab ×1'],
      });
    });
    it('is silent for a paragraph with only a page break and tab stops', async () => {
      const bytes = await docx('<w:p><w:pPr><w:tabs><w:tab w:pos="720"/></w:tabs></w:pPr><w:r><w:br w:type="page"/></w:r></w:p>');
      expect(await item(bytes, 'docx.tabsAndBreaks')).toBeUndefined();
    });
  });

  describe('docx.bookmarks', () => {
    it('counts bookmarks (not _GoBack) and internal-anchor hyperlinks (not external ones)', async () => {
      const bytes = await docx(
        '<w:p><w:bookmarkStart w:id="0" w:name="_GoBack"/><w:bookmarkEnd w:id="0"/>' +
          '<w:bookmarkStart w:id="1" w:name="_Toc1"/><w:bookmarkEnd w:id="1"/>' +
          '<w:hyperlink w:anchor="_Toc1"><w:r><w:t>go</w:t></w:r></w:hyperlink>' +
          '<w:hyperlink r:id="rId5"><w:r><w:t>web</w:t></w:r></w:hyperlink></w:p>'
      );
      expect(await item(bytes, 'docx.bookmarks')).toMatchObject({ count: 2, where: ['bookmarks ×1', 'internal links ×1'] });
    });
  });

  describe('docx.headerFooterVariants', () => {
    const ref = (type: string) => `<w:headerReference w:type="${type}" r:id="rId1"/>`;
    it('detects titlePg, first-page and even-page references', async () => {
      const bytes = await docx(
        `${p('a', `<w:sectPr>${ref('default')}${ref('first')}<w:titlePg/></w:sectPr>`)}` +
          `${p('b', `<w:sectPr>${ref('even')}</w:sectPr>`)}` +
          `<w:sectPr>${ref('default')}</w:sectPr>`
      );
      expect(await item(bytes, 'docx.headerFooterVariants')).toMatchObject({
        count: 2,
        where: ['even-page reference ×1', 'first-page header/footer (titlePg) ×1'],
      });
    });
    it('counts sections with headers when evenAndOddHeaders is on in settings', async () => {
      const bytes = await docx(`<w:sectPr>${ref('default')}</w:sectPr>`, {
        'word/settings.xml': `<w:settings ${NS}><w:evenAndOddHeaders/></w:settings>`,
      });
      expect(await item(bytes, 'docx.headerFooterVariants')).toMatchObject({ count: 1, where: ['even/odd headers setting ×1'] });
    });
    it('is silent for a plain default header', async () => {
      expect(await item(await docx(`<w:sectPr>${ref('default')}</w:sectPr>`), 'docx.headerFooterVariants')).toBeUndefined();
    });
  });

  describe('docx.sectionProperties', () => {
    it('detects page-number format/restart and vertical alignment', async () => {
      const bytes = await docx(
        p('a', '<w:sectPr><w:pgNumType w:start="1" w:fmt="lowerRoman"/></w:sectPr>') +
          `<w:sectPr><w:vAlign w:val="center"/></w:sectPr>`
      );
      expect(await item(bytes, 'docx.sectionProperties')).toMatchObject({
        count: 2,
        where: ['page numbering ×1', 'vertical alignment ×1'],
      });
    });
    it('is silent for a bare pgNumType, top alignment, and table cell vAlign', async () => {
      const bytes = await docx(
        '<w:tbl><w:tr><w:tc><w:tcPr><w:vAlign w:val="center"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>' +
          '<w:sectPr><w:pgNumType/><w:vAlign w:val="top"/></w:sectPr>'
      );
      expect(await item(bytes, 'docx.sectionProperties')).toBeUndefined();
    });
  });

  describe('docx.fields', () => {
    const fld = (...instr: string[]) =>
      `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>${instr.map((i) => `<w:r><w:instrText xml:space="preserve">${i}</w:instrText></w:r>`).join('')}` +
      `<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>cached</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`;
    it('breaks the count down by field kind and skips PAGE / NUMPAGES', async () => {
      const bytes = await docx(
        fld(' PAGEREF _Toc1 \\h ') + fld(' PAGEREF _Toc2 \\h ') + fld(' PAGE ') + fld(' NUMPAGES ') +
          fld(' SEQ Figure ') + '<w:p><w:fldSimple w:instr=" DATE \\@ &quot;yyyy&quot; "><w:r><w:t>x</w:t></w:r></w:fldSimple></w:p>'
      );
      expect(await item(bytes, 'docx.fields')).toMatchObject({ count: 4, where: ['PAGEREF ×2', 'DATE ×1', 'SEQ ×1'] });
    });
    it('counts a field once when its instruction is split across runs', async () => {
      const bytes = await docx(fld(' IF 1 ', '&gt;', ' 0 "a" "b" '));
      expect(await item(bytes, 'docx.fields')).toMatchObject({ count: 1, where: ['IF ×1'] });
    });
    it('counts nested fields separately, ignoring instruction text in a field result', async () => {
      const bytes = await docx(
        '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> TOC \\o "1-3" </w:instrText></w:r>' +
          '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
          '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> PAGEREF _Toc1 </w:instrText></w:r>' +
          '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>4</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>' +
          '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
      );
      expect(await item(bytes, 'docx.fields')).toMatchObject({ count: 2, where: ['PAGEREF ×1', 'TOC ×1'] });
    });
  });

  it('still counts a field with no end marker and an instruction outside any field', async () => {
    const bytes = await docx(
      '<w:p><w:r><w:instrText> DATE </w:instrText></w:r></w:p>' +
        '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> TOC </w:instrText></w:r></w:p>'
    );
    expect(await item(bytes, 'docx.fields')).toMatchObject({ count: 2, where: ['DATE ×1', 'TOC ×1'] });
  });

  describe("detail: 'summary' (the default, used by the import confirmation)", () => {
    it('leaves out the formatting/structure ids and the per-kind field breakdown', async () => {
      const bytes = await docx(
        p('styled', '<w:pStyle w:val="Callout"/><w:spacing w:line="360"/>') +
          '<w:p><w:r><w:t>a</w:t><w:tab/></w:r></w:p>' +
          '<w:p><w:fldSimple w:instr=" DATE "><w:r><w:t>x</w:t></w:r></w:fldSimple></w:p>' +
          '<w:sectPr><w:pgNumType w:start="3"/></w:sectPr>'
      );
      const summary = await detectDocxLoss(bytes);
      expect(summary.items).toEqual([
        { id: 'docx.fields', label: 'field(s) (table of contents, cross-references, dates…)', count: 1, effect: 'their displayed value is dropped' },
      ]);
      const full = await detectDocxLoss(bytes, { detail: 'full' });
      expect(full.items.map((i) => i.id).sort()).toEqual([
        'docx.fields',
        'docx.namedStyles',
        'docx.paragraphFormatting',
        'docx.sectionProperties',
        'docx.tabsAndBreaks',
      ]);
      expect(full.items.find((i) => i.id === 'docx.fields')?.where).toEqual(['DATE ×1']);
    });
  });

  it('does not throw on a document.xml that is not XML', async () => {
    const zip = new JSZip();
    zip.file('word/document.xml', 'not xml at all <<<');
    const bytes = await zip.generateAsync({ type: 'uint8array' });
    expect((await detectDocxLoss(bytes)).format).toBe('docx');
    expect((await detectDocxLoss(bytes, { detail: 'full' })).format).toBe('docx');
  });
});
