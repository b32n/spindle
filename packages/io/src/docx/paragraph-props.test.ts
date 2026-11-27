import JSZip from 'jszip';
import { exportDocx } from './writer';
import { importDocx } from './reader';
import type { DocBlock, DocInline, DocSection, DocumentExportInput } from '@b32nio/spindle-shared/export/docx';

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

async function docx(body: string, extra: Record<string, string> = {}): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types/>');
  zip.file('word/document.xml', `<?xml version="1.0"?><w:document ${NS}><w:body>${body}<w:sectPr/></w:body></w:document>`);
  for (const [n, x] of Object.entries(extra)) zip.file(n, x);
  return zip.generateAsync({ type: 'uint8array' });
}
const styles = (inner: string) => `<?xml version="1.0"?><w:styles ${NS}>${inner}</w:styles>`;
const p = (pPr: string, text = 'x') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}<w:r><w:t>${text}</w:t></w:r></w:p>`;
const first = async (body: string, extra: Record<string, string> = {}) => (await importDocx(await docx(body, extra))).sections[0].blocks[0] as DocBlock & Record<string, unknown>;

describe('paragraph properties on import', () => {
  describe('line spacing', () => {
    it('maps auto spacing to the keywords or a multiple', async () => {
      expect(await first(p('<w:spacing w:line="240" w:lineRule="auto"/>'))).toMatchObject({ lineSpacing: 'single' });
      expect(await first(p('<w:spacing w:line="360" w:lineRule="auto"/>'))).toMatchObject({ lineSpacing: 'onePointFive' });
      expect(await first(p('<w:spacing w:line="480"/>'))).toMatchObject({ lineSpacing: 'double' }); // lineRule defaults to auto
      const odd = await first(p('<w:spacing w:line="276" w:lineRule="auto"/>'));
      expect(odd.lineSpacingCustom).toEqual({ rule: 'multiple', value: 1.15 });
      expect(odd.lineSpacing).toBeUndefined();
    });

    it('maps atLeast and exact to heights in points', async () => {
      expect((await first(p('<w:spacing w:line="300" w:lineRule="atLeast"/>'))).lineSpacingCustom).toEqual({ rule: 'atLeast', value: 15 });
      expect((await first(p('<w:spacing w:line="320" w:lineRule="exact"/>'))).lineSpacingCustom).toEqual({ rule: 'exactly', value: 16 });
    });

    it('ignores a line value that is not a positive number and merges spacing attributes through styles', async () => {
      const b = await first(p('<w:spacing w:line="0"/>'));
      expect(b.lineSpacing).toBeUndefined();
      expect(b.lineSpacingCustom).toBeUndefined();
      const s = styles('<w:docDefaults><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>');
      const c = await first(p('<w:spacing w:before="120"/>'), { 'word/styles.xml': s });
      expect(c).toMatchObject({ spaceBeforePt: 6, spaceAfterPt: 8 });
      expect(c.lineSpacingCustom).toEqual({ rule: 'multiple', value: 1.079 });
    });
  });

  describe('indents', () => {
    it('reads right, start/end and a hanging indent as a negative first-line indent', async () => {
      expect(await first(p('<w:ind w:left="720" w:right="360" w:hanging="360"/>'))).toMatchObject({ leftIndentPt: 36, rightIndentPt: 18, firstLineIndentPt: -18 });
      expect(await first(p('<w:ind w:start="720" w:end="240" w:firstLine="360"/>'))).toMatchObject({ leftIndentPt: 36, rightIndentPt: 12, firstLineIndentPt: 18 });
    });

    it('lets a derived layer replace a base hanging indent with a first-line one, and keeps it when only left is set', async () => {
      const s = styles(
        '<w:style w:type="paragraph" w:styleId="Hang"><w:name w:val="Hang"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:style>'
      );
      const extra = { 'word/styles.xml': s };
      expect((await first(p('<w:pStyle w:val="Hang"/><w:ind w:left="1000"/>'), extra)).firstLineIndentPt).toBe(-18);
      expect((await first(p('<w:pStyle w:val="Hang"/><w:ind w:firstLine="200"/>'), extra)).firstLineIndentPt).toBe(10);
    });
  });

  describe('pagination flags', () => {
    it('reads keepNext, keepLines and pageBreakBefore, directly and through a style, and respects off', async () => {
      expect(await first(p('<w:keepNext/><w:keepLines/><w:pageBreakBefore/>'))).toMatchObject({ keepWithNext: true, keepTogether: true, pageBreakBefore: true });
      const s = styles('<w:style w:type="paragraph" w:styleId="K"><w:name w:val="K"/><w:pPr><w:keepNext/><w:pageBreakBefore/></w:pPr></w:style>');
      const extra = { 'word/styles.xml': s };
      expect(await first(p('<w:pStyle w:val="K"/>'), extra)).toMatchObject({ keepWithNext: true, pageBreakBefore: true });
      const off = await first(p('<w:pStyle w:val="K"/><w:keepNext w:val="0"/><w:pageBreakBefore w:val="false"/>'), extra);
      expect(off.keepWithNext).toBeUndefined();
      expect(off.pageBreakBefore).toBeUndefined();
    });
  });
});

const run = (rPr: string, text = 'x') => `<w:p><w:r><w:rPr>${rPr}</w:rPr><w:t>${text}</w:t></w:r></w:p>`;
const runOf = async (body: string, extra: Record<string, string> = {}) => ((await first(body, extra)).content as DocInline[])[0] as unknown as Record<string, unknown>;

describe('run properties on import', () => {
  it('reads caps and hidden text, and toggles them through styles like bold', async () => {
    expect(await runOf(run('<w:caps/><w:vanish/>'))).toMatchObject({ allCaps: true, hidden: true });
    expect((await runOf(run('<w:caps w:val="0"/>'))).allCaps).toBeUndefined();
    const s = styles(
      '<w:style w:type="paragraph" w:styleId="Caps"><w:name w:val="Caps"/><w:rPr><w:caps/></w:rPr></w:style>' +
        '<w:style w:type="character" w:styleId="Lower"><w:name w:val="Lower"/><w:rPr><w:caps/></w:rPr></w:style>'
    );
    const body = (rPr: string) => `<w:p><w:pPr><w:pStyle w:val="Caps"/></w:pPr><w:r><w:rPr>${rPr}</w:rPr><w:t>x</w:t></w:r></w:p>`;
    expect(await runOf(body(''), { 'word/styles.xml': s })).toMatchObject({ allCaps: true });
    expect((await runOf(body('<w:rStyle w:val="Lower"/>'), { 'word/styles.xml': s })).allCaps).toBeUndefined(); // toggled off
  });

  it('uses the high-ANSI font when there is no ASCII one, and keeps East Asian and complex fonts only when they differ', async () => {
    expect((await runOf(run('<w:rFonts w:hAnsi="Verdana"/>'))).fontFamily).toBe('Verdana');
    expect((await runOf(run('<w:rFonts w:ascii="Georgia" w:hAnsi="Verdana"/>'))).fontFamily).toBe('Georgia');
    expect(await runOf(run('<w:rFonts w:ascii="Georgia" w:eastAsia="MS Mincho" w:cs="Arial"/>'))).toMatchObject({
      fontFamily: 'Georgia',
      fontFamilyEastAsia: 'MS Mincho',
      fontFamilyComplex: 'Arial',
    });
    const same = await runOf(run('<w:rFonts w:ascii="Georgia" w:eastAsia="Georgia" w:cs="Georgia" w:hAnsi="Georgia"/>'));
    expect(same.fontFamilyEastAsia).toBeUndefined();
    expect(same.fontFamilyComplex).toBeUndefined();
  });

  it('resolves East Asian and complex theme fonts that name a typeface and ignores empty ones', async () => {
    const theme = (ea: string) =>
      `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fontScheme name="f"><a:majorFont><a:latin typeface="Cambria"/><a:ea typeface="${ea}"/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface="Tahoma"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>`;
    const rPr = '<w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="majorEastAsia" w:cstheme="minorBidi"/>';
    const r = await runOf(run(rPr), { 'word/theme/theme1.xml': theme('MS Gothic') });
    expect(r).toMatchObject({ fontFamily: 'Calibri', fontFamilyEastAsia: 'MS Gothic', fontFamilyComplex: 'Tahoma' });
    const empty = await runOf(run('<w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia"/>'), { 'word/theme/theme1.xml': theme('') });
    expect(empty.fontFamilyEastAsia).toBeUndefined();
  });
});

describe('numbered list starts on import', () => {
  const numbering = (abstract: string, nums: string) => ({ 'word/numbering.xml': `<?xml version="1.0"?><w:numbering ${NS}>${abstract}${nums}</w:numbering>` });
  const abstractNum = (id: number, start: number, fmt = 'decimal') =>
    `<w:abstractNum w:abstractNumId="${id}"><w:lvl w:ilvl="0"><w:start w:val="${start}"/><w:numFmt w:val="${fmt}"/></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/></w:lvl></w:abstractNum>`;
  const num = (id: number, abstract: number, override?: number) =>
    `<w:num w:numId="${id}"><w:abstractNumId w:val="${abstract}"/>${override !== undefined ? `<w:lvlOverride w:ilvl="0"><w:startOverride w:val="${override}"/></w:lvlOverride>` : ''}</w:num>`;
  const li = (numId: number, ilvl = 0, text = 'i') => p(`<w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr>`, text);
  const items = async (body: string, extra: Record<string, string>) =>
    (await importDocx(await docx(body, extra))).sections[0].blocks.map((b) => (b.type === 'listItem' ? [b.listType, b.level, b.start] : [b.type]));

  it('keeps a first number other than 1 on the first item only', async () => {
    const r = await items(li(1) + li(1) + li(1), numbering(abstractNum(0, 5), num(1, 0)));
    expect(r).toEqual([['numbered', 0, 5], ['numbered', 0, undefined], ['numbered', 0, undefined]]);
  });

  it('leaves a list that starts at 1 without a start', async () => {
    expect(await items(li(1) + li(1), numbering(abstractNum(0, 1), num(1, 0)))).toEqual([['numbered', 0, undefined], ['numbered', 0, undefined]]);
  });

  it('carries on counting after an interruption, as Word does for a list sharing its numbering', async () => {
    const r = await items(li(1) + li(1) + p('', 'break') + li(1) + li(1), numbering(abstractNum(0, 1), num(1, 0)));
    expect(r).toEqual([['numbered', 0, undefined], ['numbered', 0, undefined], ['paragraph'], ['numbered', 0, 3], ['numbered', 0, undefined]]);
  });

  it('restarts at the first number of a startOverride, also directly after another list', async () => {
    const r = await items(li(1) + li(1) + li(2) + li(2), numbering(abstractNum(0, 1), num(1, 0) + num(2, 0, 1)));
    expect(r).toEqual([['numbered', 0, undefined], ['numbered', 0, undefined], ['numbered', 0, 1], ['numbered', 0, undefined]]);
    const at9 = await items(li(1) + li(1) + li(2), numbering(abstractNum(0, 1), num(1, 0) + num(2, 0, 9)));
    expect(at9[2]).toEqual(['numbered', 0, 9]);
  });

  it('restarts for a separate list definition that follows directly', async () => {
    const r = await items(li(1) + li(1) + li(2), numbering(abstractNum(0, 1) + abstractNum(1, 1), num(1, 0) + num(2, 1)));
    expect(r[2]).toEqual(['numbered', 0, 1]);
  });

  it('counts nested levels separately and starts them again after each parent', async () => {
    const r = await items(li(1, 0) + li(1, 1) + li(1, 1) + li(1, 0) + li(1, 1), numbering(abstractNum(0, 1), num(1, 0)));
    // only top-level items may carry a start; here nothing restarts
    expect(r.every((x) => x[2] === undefined)).toBe(true);
  });

  it('does not number bullets and breaks a numbered run at a bullet', async () => {
    const r = await items(
      li(1) + li(3) + li(1),
      numbering(abstractNum(0, 1) + abstractNum(1, 1, 'bullet'), num(1, 0) + num(3, 1))
    );
    expect(r).toEqual([['numbered', 0, undefined], ['bullet', 0, undefined], ['numbered', 0, 2]]);
  });

  it('ignores a numId with no definition', async () => {
    expect(await items(li(9), numbering(abstractNum(0, 1), num(1, 0)))).toEqual([['numbered', 0, undefined]]);
  });

  it('continues the count across a section break', async () => {
    const sect = '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:sectPr/></w:pPr><w:r><w:t>last</w:t></w:r></w:p>';
    const result = await importDocx(await docx(li(1) + sect + li(1), numbering(abstractNum(0, 1), num(1, 0))));
    expect(result.sections).toHaveLength(2);
    expect(result.sections[1].blocks[0]).toMatchObject({ type: 'listItem', start: 3 });
  });
});

// ---------------------------------------------------------------------------
// Export and round trips
// ---------------------------------------------------------------------------

const SECTION: Omit<DocSection, 'blocks'> = {
  pageWidthPx: 816, pageHeightPx: 1056, marginTopPx: 96, marginRightPx: 96, marginBottomPx: 96, marginLeftPx: 96, orientation: 'portrait',
};
const input = (blocks: DocBlock[]): DocumentExportInput => ({ title: 'T', sections: [{ ...SECTION, blocks }] });
const t = (text: string, extra: object = {}) => ({ type: 'text' as const, text, ...extra });
const xmlOf = async (i: DocumentExportInput, part: string) => {
  const zip = await JSZip.loadAsync((await exportDocx(i)).bytes);
  return (await zip.file(part)?.async('string')) ?? '';
};
const roundTrip = async (blocks: DocBlock[]) => (await importDocx((await exportDocx(input(blocks))).bytes)).sections[0].blocks;

describe('paragraph properties on export', () => {
  it('writes line spacing, indents and pagination flags', async () => {
    const doc = await xmlOf(
      input([
        { type: 'paragraph', content: [t('a')], lineSpacingCustom: { rule: 'multiple', value: 1.15 }, rightIndentPt: 18, firstLineIndentPt: -18, leftIndentPt: 36, keepWithNext: true, keepTogether: true, pageBreakBefore: true },
        { type: 'paragraph', content: [t('b')], lineSpacingCustom: { rule: 'atLeast', value: 15 } },
        { type: 'paragraph', content: [t('c')], lineSpacingCustom: { rule: 'exactly', value: 16 } },
        { type: 'paragraph', content: [t('d')], lineSpacing: 'double', firstLineIndentPt: 12 },
      ]),
      'word/document.xml'
    );
    expect(doc).toContain('<w:keepNext/><w:keepLines/><w:pageBreakBefore/>');
    expect(doc).toContain('<w:spacing w:line="276" w:lineRule="auto"/>');
    expect(doc).toContain('<w:ind w:left="720" w:right="360" w:hanging="360"/>');
    expect(doc).toContain('<w:spacing w:line="300" w:lineRule="atLeast"/>');
    expect(doc).toContain('<w:spacing w:line="320" w:lineRule="exact"/>');
    expect(doc).toMatch(/<w:spacing w:line="480"[^>]*\/><w:ind w:firstLine="240"\/>/);
  });

  it('round-trips the paragraph properties', async () => {
    const block = {
      type: 'paragraph' as const,
      content: [t('a')],
      lineSpacingCustom: { rule: 'multiple' as const, value: 1.15 },
      rightIndentPt: 18,
      firstLineIndentPt: -18,
      leftIndentPt: 36,
      keepWithNext: true,
      keepTogether: true,
      pageBreakBefore: true,
    };
    const heading = { type: 'heading' as const, level: 2 as const, content: [t('h')], lineSpacing: 'onePointFive' as const, spaceBeforePt: 12, spaceAfterPt: 0, keepWithNext: true };
    const out = await roundTrip([block, heading]);
    expect(out[0]).toMatchObject(block);
    expect(out[1]).toMatchObject(heading);
  });

  it('writes caps, hidden text and per-script fonts, and round-trips them', async () => {
    const content = [
      t('a', { allCaps: true, hidden: true }),
      t('b', { fontFamily: 'Georgia', fontFamilyEastAsia: 'MS Mincho', fontFamilyComplex: 'Arial' }),
      t('c', { fontFamilyEastAsia: 'MS Mincho' }),
    ];
    const doc = await xmlOf(input([{ type: 'paragraph', content }]), 'word/document.xml');
    expect(doc).toContain('<w:caps/><w:vanish/>');
    // (attribute order is the library's)
    expect(doc).toMatch(/<w:rFonts(?=[^>]*w:ascii="Georgia")(?=[^>]*w:hAnsi="Georgia")(?=[^>]*w:eastAsia="MS Mincho")(?=[^>]*w:cs="Arial")[^>]*\/>/);
    const out = (await roundTrip([{ type: 'paragraph', content }]))[0] as unknown as { content: Array<Record<string, unknown>> };
    expect(out.content[0]).toMatchObject({ allCaps: true, hidden: true });
    expect(out.content[1]).toMatchObject({ fontFamily: 'Georgia', fontFamilyEastAsia: 'MS Mincho', fontFamilyComplex: 'Arial' });
    expect(out.content[2]).toMatchObject({ fontFamilyEastAsia: 'MS Mincho' });
  });
});

describe('numbered lists on export', () => {
  const item = (text: string, extra: object = {}): DocBlock => ({ type: 'listItem', listType: 'numbered', level: 0, content: [t(text)], ...extra }) as DocBlock;
  const para = (text: string): DocBlock => ({ type: 'paragraph', content: [t(text)] });
  const numIds = (doc: string) => [...doc.matchAll(/<w:numId w:val="(\d+)"\/>/g)].map((m) => m[1]);

  it('starts every separate numbered list at its own number instead of carrying on counting', async () => {
    const doc = await xmlOf(input([item('a'), item('b'), para('between'), item('c'), item('d')]), 'word/document.xml');
    const ids = numIds(doc);
    expect(ids[0]).toBe(ids[1]);
    expect(ids[2]).toBe(ids[3]);
    expect(ids[0]).not.toBe(ids[2]);
  });

  it('writes the first number of a list and restarts at an item that states its own', async () => {
    const blocks = [item('a', { start: 5 }), item('b'), item('c', { start: 1 }), item('d')];
    const numberingXml = await xmlOf(input(blocks), 'word/numbering.xml');
    expect(numberingXml).toContain('<w:startOverride w:val="5"/>');
    const out = await roundTrip(blocks);
    expect(out.map((b) => (b.type === 'listItem' ? b.start : 'x'))).toEqual([5, undefined, 1, undefined]);
  });

  it('keeps a numbered list continuing from an earlier one (a start above 1 after a break)', async () => {
    const out = await roundTrip([item('a'), item('b'), para('between'), item('c', { start: 3 }), item('d')]);
    expect(out.map((b) => (b.type === 'listItem' ? b.start : 'x'))).toEqual([undefined, undefined, 'x', 3, undefined]);
  });

  it('ignores a start that is not a positive whole number and leaves bullets alone', async () => {
    const out = await roundTrip([item('a', { start: 0 }), item('b'), { type: 'listItem', listType: 'bullet', level: 0, content: [t('bullet')], start: 4 } as DocBlock]);
    expect(out.map((b) => (b.type === 'listItem' ? [b.listType, b.start] : 'x'))).toEqual([['numbered', undefined], ['numbered', undefined], ['bullet', undefined]]);
  });
});
