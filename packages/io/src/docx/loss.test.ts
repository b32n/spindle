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

  describe('named styles', () => {
    it('are kept on import, so they are no longer reported as lost', async () => {
      const parts = { 'word/styles.xml': styles(['Callout', 'Callout']) };
      const bytes = await docx(p('a', '<w:pStyle w:val="Callout"/>') + '<w:p><w:r><w:rPr><w:rStyle w:val="Emphasis2"/></w:rPr><w:t>d</w:t></w:r></w:p>', parts);
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

  describe('paragraph formatting (imported, so no longer a loss)', () => {
    it.each([
      ['line spacing', '<w:spacing w:line="360" w:lineRule="auto"/>'],
      ['right indent', '<w:ind w:right="720"/>'],
      ['hanging indent', '<w:ind w:left="720" w:hanging="360"/>'],
      ['keep with next', '<w:keepNext/>'],
      ['keep lines together', '<w:keepLines/>'],
      ['page break before', '<w:pageBreakBefore/>'],
    ])('does not report %s', async (_kind, pPr) => {
      expect(await item(await docx(p('a', pPr)), 'docx.paragraphFormatting')).toBeUndefined();
    });
    it('does not report hidden text', async () => {
      const bytes = await docx('<w:p><w:r><w:rPr><w:vanish/></w:rPr><w:t>secret</w:t></w:r></w:p>');
      expect((await detectDocxLoss(bytes, { detail: 'full' })).items).toEqual([]);
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
    it('detects other number formats only; custom starts and start overrides are imported', async () => {
      const bytes = await docx(li('2') + li('3') + li('5'), { 'word/numbering.xml': numbering });
      expect(await item(bytes, 'docx.listFormatting')).toMatchObject({ count: 1, where: ['lowerRoman ×1'] });
      expect(await item(await docx(li('3') + li('5'), { 'word/numbering.xml': numbering }), 'docx.listFormatting')).toBeUndefined();
    });
  });

  describe('bookmarks', () => {
    it('are no longer reported: bookmarks and internal links are imported', async () => {
      const bytes = await docx(
        '<w:p><w:bookmarkStart w:id="0" w:name="_GoBack"/><w:bookmarkEnd w:id="0"/>' +
          '<w:bookmarkStart w:id="1" w:name="_Toc1"/><w:bookmarkEnd w:id="1"/>' +
          '<w:hyperlink w:anchor="_Toc1"><w:r><w:t>go</w:t></w:r></w:hyperlink>' +
          '<w:hyperlink r:id="rId5"><w:r><w:t>web</w:t></w:r></w:hyperlink></w:p>'
      );
      expect(await item(bytes, 'docx.bookmarks')).toBeUndefined();
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
    it('counts only the fields that are not kept as fields (tables of contents, indexes, legacy form fields), by kind', async () => {
      const bytes = await docx(
        fld(' PAGEREF _Toc1 \\h ') + fld(' PAGEREF _Toc2 \\h ') + fld(' PAGE ') + fld(' NUMPAGES ') +
          fld(' SEQ Figure ') + '<w:p><w:fldSimple w:instr=" DATE \\@ &quot;yyyy&quot; "><w:r><w:t>x</w:t></w:r></w:fldSimple></w:p>' +
          fld(' TOC \\o "1-3" ') + fld(' INDEX ') + fld(' FORMTEXT ')
      );
      expect(await item(bytes, 'docx.fields')).toMatchObject({ count: 3, where: ['FORMTEXT ×1', 'INDEX ×1', 'TOC ×1'] });
    });
    it('counts a field once when its instruction is split across runs', async () => {
      const bytes = await docx(fld(' TOC ', '\\o', ' "1-3" '));
      expect(await item(bytes, 'docx.fields')).toMatchObject({ count: 1, where: ['TOC ×1'] });
    });
    it('counts nested fields separately, ignoring instruction text in a field result', async () => {
      const bytes = await docx(
        '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> TOC \\o "1-3" </w:instrText></w:r>' +
          '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
          '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> PAGEREF _Toc1 </w:instrText></w:r>' +
          '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>4</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>' +
          '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
      );
      expect(await item(bytes, 'docx.fields')).toMatchObject({ count: 1, where: ['TOC ×1'] });
    });
  });

  it('still counts a field with no end marker and an instruction outside any field', async () => {
    const bytes = await docx(
      '<w:p><w:r><w:instrText> INDEX </w:instrText></w:r></w:p>' +
        '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> TOC </w:instrText></w:r></w:p>'
    );
    expect(await item(bytes, 'docx.fields')).toMatchObject({ count: 2, where: ['INDEX ×1', 'TOC ×1'] });
  });

  describe("detail: 'summary' (the default, used by the import confirmation)", () => {
    it('leaves out the formatting/structure ids and the per-kind field breakdown', async () => {
      const bytes = await docx(
        p('styled', '<w:pStyle w:val="Callout"/><w:spacing w:line="360"/>') +
          '<w:p><w:fldSimple w:instr=" TOC "><w:r><w:t>x</w:t></w:r></w:fldSimple></w:p>' +
          '<w:sectPr><w:pgNumType w:start="3"/></w:sectPr>'
      );
      const summary = await detectDocxLoss(bytes);
      expect(summary.items).toEqual([
        {
          id: 'docx.fields',
          label: 'field(s) (table of contents, index, legacy form fields)',
          count: 1,
          effect: 'tables of contents and indexes are kept as static text, legacy form fields lose their value',
        },
      ]);
      const full = await detectDocxLoss(bytes, { detail: 'full' });
      expect(full.items.map((i) => i.id).sort()).toEqual([
        'docx.fields',
        'docx.sectionProperties',
      ]);
      expect(full.items.find((i) => i.id === 'docx.fields')?.where).toEqual(['TOC ×1']);
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

describe('detectDocxLoss: long-tail detectors', () => {
  const full = { detail: 'full' as const };
  const run = (bytes: Uint8Array, id: string) => detectDocxLoss(bytes, full).then((r) => r.items.find((i) => i.id === id));
  const drawing = (kind: 'inline' | 'anchor', uri: string) =>
    `<w:p><w:r><w:drawing><wp:${kind}><a:graphic><a:graphicData uri="${uri}"/></a:graphic></wp:${kind}></w:drawing></w:r></w:p>`;
  const OOXML = 'http://schemas.openxmlformats.org/drawingml/2006';

  it('docx.trackedMoves counts move marks, not the range markers', async () => {
    const bytes = await docx(
      '<w:p><w:moveFromRangeStart w:id="1" w:name="m1"/><w:moveFrom w:id="2"><w:r><w:t>a</w:t></w:r></w:moveFrom><w:moveFromRangeEnd w:id="1"/>' +
        '<w:moveTo w:id="3"><w:r><w:t>a</w:t></w:r></w:moveTo></w:p>'
    );
    expect(await run(bytes, 'docx.trackedMoves')).toMatchObject({ count: 2 });
  });

  it('docx.formattingRevisions counts the four change kinds by name', async () => {
    const bytes = await docx(
      '<w:p><w:pPr><w:pPrChange w:id="1"><w:pPr/></w:pPrChange></w:pPr><w:r><w:rPr><w:rPrChange w:id="2"><w:rPr/></w:rPrChange></w:rPr><w:t>a</w:t></w:r></w:p>' +
        '<w:tbl><w:tblPr><w:tblPrChange w:id="3"><w:tblPr/></w:tblPrChange></w:tblPr></w:tbl>' +
        '<w:sectPr><w:sectPrChange w:id="4"><w:sectPr/></w:sectPrChange></w:sectPr>'
    );
    expect(await run(bytes, 'docx.formattingRevisions')).toMatchObject({
      count: 4,
      where: ['pPrChange ×1', 'rPrChange ×1', 'sectPrChange ×1', 'tblPrChange ×1'],
    });
  });

  it('docx.dropCaps counts drop and margin drop caps but not plain text frames', async () => {
    const bytes = await docx(
      p('a', '<w:framePr w:dropCap="drop" w:lines="3" w:wrap="around"/>') +
        p('b', '<w:framePr w:dropCap="margin"/>') +
        p('c', '<w:framePr w:w="2000" w:h="1000" w:hRule="exact" w:wrap="around"/>') +
        p('d', '<w:framePr w:dropCap="none"/>')
    );
    expect(await run(bytes, 'docx.dropCaps')).toMatchObject({ count: 2 });
  });

  it('docx.watermarks looks in headers and footers for watermark shapes', async () => {
    const header = (id: string) => `<w:hdr ${NS} xmlns:v="x"><w:p><w:r><w:pict><v:shape id="${id}"><v:textpath string="DRAFT"/></v:shape></w:pict></w:r></w:p></w:hdr>`;
    const bytes = await docx(p('a'), {
      'word/header1.xml': header('PowerPlusWaterMarkObject357476642'),
      'word/header2.xml': header('Shape1'),
      'word/footer1.xml': header('WordPictureWatermark1'),
    });
    expect(await run(bytes, 'docx.watermarks')).toMatchObject({ count: 2 });
  });

  it('docx.pageBorders and docx.lineNumbers count section properties', async () => {
    const bytes = await docx(
      p('a', '<w:sectPr><w:pgBorders w:offsetFrom="page"><w:top w:val="single"/></w:pgBorders><w:lnNumType w:countBy="1"/></w:sectPr>') +
        '<w:sectPr><w:pgBorders/></w:sectPr>'
    );
    expect(await run(bytes, 'docx.pageBorders')).toMatchObject({ count: 2 });
    expect(await run(bytes, 'docx.lineNumbers')).toMatchObject({ count: 1 });
  });

  it('docx.hyphenation reads autoHyphenation from settings and honours val=0', async () => {
    const on = await docx(p('a'), { 'word/settings.xml': `<w:settings ${NS}><w:autoHyphenation/></w:settings>` });
    const off = await docx(p('a'), { 'word/settings.xml': `<w:settings ${NS}><w:autoHyphenation w:val="0"/></w:settings>` });
    expect(await run(on, 'docx.hyphenation')).toMatchObject({ count: 1 });
    expect(await run(off, 'docx.hyphenation')).toBeUndefined();
    expect(await run(await docx(p('a')), 'docx.hyphenation')).toBeUndefined();
  });

  it('docx.tabStops counts custom stops by leader and ignores cleared stops', async () => {
    const bytes = await docx(
      p('a', '<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9000"/><w:tab w:val="left" w:pos="720"/><w:tab w:val="clear" w:pos="1440"/><w:tab w:val="left" w:leader="none" w:pos="2000"/></w:tabs>')
    );
    expect(await run(bytes, 'docx.tabStops')).toMatchObject({ count: 3, where: ['no leader ×2', 'dot leader ×1'] });
  });

  it('docx.rtl counts paragraph, run and table direction and ignores switched-off ones', async () => {
    const bytes = await docx(
      p('a', '<w:bidi/>') +
        p('b', '<w:bidi w:val="0"/>') +
        '<w:p><w:r><w:rPr><w:rtl/></w:rPr><w:t>x</w:t></w:r></w:p>' +
        '<w:tbl><w:tblPr><w:bidiVisual/></w:tblPr></w:tbl>'
    );
    expect(await run(bytes, 'docx.rtl')).toMatchObject({
      count: 3,
      where: ['paragraph direction ×1', 'run direction ×1', 'table direction ×1'],
    });
  });

  it('docx.ruby and docx.altChunk count their elements; tabs, breaks and smart tags are no longer reported (their text is kept)', async () => {
    const bytes = await docx(
      '<w:p><w:r><w:ruby><w:rt/><w:rubyBase/></w:ruby></w:r><w:smartTag w:uri="u" w:element="place"><w:r><w:t>Paris</w:t></w:r></w:smartTag></w:p>' +
        '<w:altChunk r:id="rId9"/>'
    );
    expect(await run(bytes, 'docx.ruby')).toMatchObject({ count: 1 });
    expect(await run(bytes, 'docx.altChunk')).toMatchObject({ count: 1 });
    expect(await run(bytes, 'docx.smartTags')).toBeUndefined();
    const text = await docx('<w:p><w:r><w:t>a</w:t><w:tab/><w:br/><w:noBreakHyphen/><w:softHyphen/></w:r></w:p>');
    expect(await run(text, 'docx.tabsAndBreaks')).toBeUndefined();
  });

  it('docx.embeddedFonts counts font files under word/fonts', async () => {
    const bytes = await docx(p('a'), {
      'word/fonts/font1.odttf': 'x',
      'word/fonts/font2.odttf': 'x',
      'word/fonts/_rels/fontTable.xml.rels': '<Relationships/>',
      'word/media/image1.png': 'x',
    });
    expect(await run(bytes, 'docx.embeddedFonts')).toMatchObject({ count: 2 });
  });

  it('docx.legacyForms counts FORMTEXT/FORMCHECKBOX/FORMDROPDOWN, which stay in docx.fields too (a DATE field is kept as a field, so is in neither)', async () => {
    const fld = (i: string) => `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> ${i} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`;
    const bytes = await docx(fld('FORMTEXT') + fld('FORMCHECKBOX') + fld('FORMDROPDOWN') + fld('DATE'));
    expect(await run(bytes, 'docx.legacyForms')).toMatchObject({ count: 3, where: ['FORMCHECKBOX ×1', 'FORMDROPDOWN ×1', 'FORMTEXT ×1'] });
    expect(await run(bytes, 'docx.fields')).toMatchObject({ count: 3 });
  });

  describe('drawing counts (also in the summary report)', () => {
    it('counts an anchored picture as a floating image, but not inline pictures or anchored shapes/text boxes/charts', async () => {
      const bytes = await docx(
        drawing('anchor', `${OOXML}/picture`) +
          drawing('inline', `${OOXML}/picture`) +
          drawing('anchor', 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape') +
          drawing('anchor', `${OOXML}/chart`)
      );
      for (const options of [undefined, full]) {
        const items = (await detectDocxLoss(bytes, options)).items;
        expect(items.find((i) => i.id === 'docx.floatingImages')).toMatchObject({ count: 1 });
        expect(items.find((i) => i.id === 'docx.charts')).toMatchObject({ count: 1 });
      }
    });

    it('counts charts and SmartArt as objects, however many parts back them', async () => {
      const bytes = await docx(drawing('inline', `${OOXML}/chart`) + drawing('inline', `${OOXML}/chart`) + drawing('inline', `${OOXML}/diagram`), {
        'word/charts/chart1.xml': '<c/>',
        'word/charts/chart2.xml': '<c/>',
        'word/charts/style1.xml': '<c/>',
        'word/charts/colors1.xml': '<c/>',
        'word/diagrams/data1.xml': '<d/>',
        'word/diagrams/layout1.xml': '<d/>',
        'word/diagrams/quickStyle1.xml': '<d/>',
        'word/diagrams/colors1.xml': '<d/>',
        'word/diagrams/drawing1.xml': '<d/>',
      });
      expect((await detectDocxLoss(bytes)).items.find((i) => i.id === 'docx.charts')).toMatchObject({ count: 3 });
    });

    it('does not report charts for orphan parts that no drawing references', async () => {
      const bytes = await docx(p('a'), { 'word/charts/chart1.xml': '<c/>' });
      expect((await detectDocxLoss(bytes)).items).toEqual([]);
    });
  });

  it('summary mode leaves the long-tail ids out', async () => {
    const bytes = await docx(p('a', '<w:bidi/><w:framePr w:dropCap="drop"/>') + '<w:p><w:smartTag><w:r><w:t>x</w:t></w:r></w:smartTag></w:p>');
    expect((await detectDocxLoss(bytes)).items).toEqual([]);
  });
});
