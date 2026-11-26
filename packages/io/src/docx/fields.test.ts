import JSZip from 'jszip';
import { exportDocx } from './writer';
import { importDocx } from './reader';
import type { DocBlock, DocInline, DocSection } from '@b32nio/spindle-shared/export/docx';

const BASE: Omit<DocSection, 'blocks'> = {
  pageWidthPx: 816,
  pageHeightPx: 1056,
  marginTopPx: 96,
  marginRightPx: 96,
  marginBottomPx: 96,
  marginLeftPx: 96,
  orientation: 'portrait',
};
const W = 'xmlns:w="x"';
const r = (inner: string, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}${inner}</w:r>`;
const t = (text: string) => `<w:t xml:space="preserve">${text}</w:t>`;
const chr = (type: string, extra = '') => r(`<w:fldChar w:fldCharType="${type}"${extra}/>`);
const instr = (text: string) => r(`<w:instrText xml:space="preserve">${text}</w:instrText>`);
const complex = (instruction: string, result: string, begin = '') => chr('begin', begin) + instr(instruction) + chr('separate') + result + chr('end');
const para = (inner: string) => `<w:p ${W}>${inner}</w:p>`;

async function importBody(bodyXml: string) {
  const exported = await exportDocx({ title: 'T', sections: [{ ...BASE, blocks: [{ type: 'paragraph', content: [{ type: 'text', text: 'x' }] }] }] });
  const zip = await JSZip.loadAsync(exported.bytes);
  const xml = (await zip.file('word/document.xml')!.async('string')).replace(/<w:p>[\s\S]*?<\/w:p>/, bodyXml);
  zip.file('word/document.xml', xml);
  return importDocx(await zip.generateAsync({ type: 'uint8array' }));
}
const contentOf = async (bodyXml: string, block = 0): Promise<DocInline[]> =>
  ((await importBody(bodyXml)).sections[0].blocks[block] as Extract<DocBlock, { type: 'paragraph' }>).content;

async function roundTrip(content: DocInline[]) {
  const exported = await exportDocx({ title: 'T', sections: [{ ...BASE, blocks: [{ type: 'paragraph', content }] }] });
  const zip = await JSZip.loadAsync(exported.bytes);
  const xml = await zip.file('word/document.xml')!.async('string');
  const imported = await importDocx(exported.bytes);
  return { xml, imported, content: (imported.sections[0].blocks[0] as Extract<DocBlock, { type: 'paragraph' }>).content };
}

describe('fields', () => {
  it('keeps a complex field as a generic field with its instruction and cached result', async () => {
    const content = await contentOf(para(r(t('See ')) + complex(' REF _Ref42 \\h ', r(t('Figure 3'))) + r(t('.'))));
    expect(content).toEqual([
      { type: 'text', text: 'See ' },
      { type: 'field', instruction: 'REF _Ref42 \\h', kind: 'REF', result: 'Figure 3' },
      { type: 'text', text: '.' },
    ]);
  });

  it('reads a field whose instruction and result share one run (the shape this writer emits)', async () => {
    const content = await contentOf(
      para(r('<w:fldChar w:fldCharType="begin"/><w:instrText> SEQ Table </w:instrText><w:fldChar w:fldCharType="separate"/>' + t('7') + '<w:fldChar w:fldCharType="end"/>'))
    );
    expect(content).toEqual([{ type: 'field', instruction: 'SEQ Table', kind: 'SEQ', result: '7' }]);
  });

  it('keeps fldSimple fields with their result', async () => {
    const content = await contentOf(para(`<w:fldSimple w:instr=" AUTHOR "><w:r><w:t>Ada</w:t></w:r></w:fldSimple>`));
    expect(content).toEqual([{ type: 'field', instruction: 'AUTHOR', kind: 'AUTHOR', result: 'Ada' }]);
  });

  it('keeps DATE, TIME and TITLE as cached text (never recalculated)', async () => {
    const content = await contentOf(para(complex(' DATE \\@ "d MMM yyyy" ', r(t('5 Mar 2024'))) + complex(' TITLE ', r(t('Harbour')))));
    expect(content).toMatchObject([
      { type: 'field', kind: 'DATE', result: '5 Mar 2024' },
      { type: 'field', kind: 'TITLE', result: 'Harbour' },
    ]);
  });

  it('keeps formatting inside a cached result as runs', async () => {
    const content = await contentOf(para(complex(' REF x ', r(t('bold'), '<w:b/>') + r(t(' plain')))));
    expect(content).toEqual([
      {
        type: 'field',
        instruction: 'REF x',
        kind: 'REF',
        result: [
          { type: 'text', text: 'bold', bold: true },
          { type: 'text', text: ' plain' },
        ],
      },
    ]);
  });

  it('keeps a nested field as part of the outer field result', async () => {
    const content = await contentOf(para(complex(' REF outer ', r(t('see page ')) + complex(' PAGEREF x ', r(t('9'))))));
    expect(content).toEqual([{ type: 'field', instruction: 'REF outer', kind: 'REF', result: 'see page 9' }]);
  });

  it('keeps the cached value of a field nested in an instruction and says so', async () => {
    const { warnings, sections } = await importBody(
      para(chr('begin') + instr(' IF ') + complex(' PAGE ', r(t('1'))) + instr(' > 0 "a" "b" ') + chr('separate') + r(t('a')) + chr('end'))
    );
    expect((sections[0].blocks[0] as Extract<DocBlock, { type: 'paragraph' }>).content).toMatchObject([{ type: 'field', kind: 'IF', instruction: 'IF 1 > 0 "a" "b"', result: 'a' }]);
    expect(warnings.some((w) => w.includes('nested in the instruction'))).toBe(true);
  });

  it('keeps PAGE and NUMPAGES as page-number fields', async () => {
    const content = await contentOf(para(complex(' PAGE ', r(t('3'))) + complex(' NUMPAGES ', r(t('9'))) + complex(' PAGEREF _Toc1 \\h ', r(t('4')))));
    expect(content).toMatchObject([{ type: 'pageNumber' }, { type: 'totalPages' }, { type: 'field', kind: 'PAGEREF', result: '4' }]);
  });

  it('turns a HYPERLINK field into links', async () => {
    const content = await contentOf(para(complex(' HYPERLINK "https://example.com/a" ', r(t('example'), '<w:b/>'))));
    expect(content).toEqual([{ type: 'link', text: 'example', href: 'https://example.com/a', bold: true }]);
  });

  it('turns a HYPERLINK field to an anchor only into an internal link', async () => {
    const content = await contentOf(para(complex(' HYPERLINK \\l "_Toc1" ', r(t('Intro')))));
    expect(content).toEqual([{ type: 'link', text: 'Intro', href: '#_Toc1' }]);
  });

  it('keeps the text of a table of contents as static paragraphs, and fields inside it', async () => {
    const { sections, warnings } = await importBody(
      para(chr('begin') + instr(' TOC \\o "1-3" ') + chr('separate') + r(t('Intro')) + r('<w:tab/>')) +
        para(r(t('Chapter')) + complex(' PAGEREF _Toc2 ', r(t('4')))) +
        para(chr('end'))
    );
    const blocks = sections[0].blocks as Array<Extract<DocBlock, { type: 'paragraph' }>>;
    expect(blocks[0].content).toEqual([{ type: 'text', text: 'Intro' }, { type: 'text', text: '\t' }]);
    expect(blocks[1].content).toMatchObject([{ type: 'text', text: 'Chapter' }, { type: 'field', kind: 'PAGEREF', result: '4' }]);
    expect(warnings.filter((w) => w.includes('Unsupported field'))).toEqual([]);
  });

  it('keeps the text a legacy form field displays as plain text, without a warning', async () => {
    const { sections, warnings } = await importBody(para(complex(' FORMTEXT ', r(t('typed'))) + complex(' FORMCHECKBOX ', '') + complex(' FORMDROPDOWN ', r(t('Red')))));
    expect((sections[0].blocks[0] as Extract<DocBlock, { type: 'paragraph' }>).content).toEqual([
      { type: 'text', text: 'typed' },
      { type: 'text', text: 'Red' },
    ]);
    expect(warnings.filter((w) => w.includes('Unsupported field'))).toEqual([]);
  });

  it('reads w:dirty', async () => {
    const content = await contentOf(para(complex(' REF x ', r(t('1')), ' w:dirty="true"')));
    expect(content).toMatchObject([{ type: 'field', dirty: true }]);
  });

  it('keeps no field result in headers and footers, as before', async () => {
    const exported = await exportDocx({
      title: 'T',
      sections: [{ ...BASE, header: { paragraphs: [{ content: [{ type: 'text', text: 'h' }] }] }, blocks: [{ type: 'paragraph', content: [] }] }],
    });
    const zip = await JSZip.loadAsync(exported.bytes);
    const name = Object.keys(zip.files).find((n) => /word\/header\d*\.xml$/.test(n))!;
    const xml = (await zip.file(name)!.async('string')).replace(/<w:p>[\s\S]*?<\/w:p>/, `<w:p>${complex(' STYLEREF "Heading 1" ', r(t('Outlook')))}</w:p>`);
    zip.file(name, xml);
    const imported = await importDocx(await zip.generateAsync({ type: 'uint8array' }));
    expect(imported.sections[0].header).toEqual({ paragraphs: [{ content: [] }] });
    expect(imported.warnings).toContain('Unsupported field code "STYLEREF "Heading 1"" — its cached value was dropped.');
  });
});

describe('field export', () => {
  it('writes begin, instruction, separate, cached result and end so Word can refresh the field', async () => {
    const { xml } = await roundTrip([{ type: 'field', instruction: 'REF _Ref1 \\h', kind: 'REF', result: 'Figure 1' }]);
    const order = ['fldCharType="begin"', 'REF _Ref1 \\h', 'fldCharType="separate"', 'Figure 1', 'fldCharType="end"'].map((s) => xml.indexOf(s));
    expect(order.every((i) => i >= 0)).toBe(true);
    expect([...order].sort((a, b) => a - b)).toEqual(order);
  });

  it('round-trips plain, formatted and dirty fields unchanged', async () => {
    const content: DocInline[] = [
      { type: 'text', text: 'a ' },
      { type: 'field', instruction: 'SEQ Figure \\* ARABIC', kind: 'SEQ', result: '2' },
      { type: 'field', instruction: 'REF x', kind: 'REF', result: [{ type: 'text', text: 'B', bold: true, color: '#ff0000' }, { type: 'text', text: 'c' }], dirty: true },
      { type: 'field', instruction: 'MERGEFIELD Name', kind: 'MERGEFIELD', result: '' },
    ];
    const { content: back, xml } = await roundTrip(content);
    expect(back).toEqual(content);
    expect(xml).toContain('w:dirty="true"');
  });
});
