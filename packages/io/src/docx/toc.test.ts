import JSZip from 'jszip';
import { readFileSync } from 'node:fs';
import path from 'node:path';
import { exportDocx } from './writer';
import { importDocx } from './reader';
import { findTocRanges, isEmptyParagraph, tocLists, withoutMarkers } from './toc-field';
import { childrenOf, find, parseXml, tagOf } from './xml-tree';
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
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
type Toc = Extract<DocBlock, { type: 'toc' }>;
type Paragraph = Extract<DocBlock, { type: 'paragraph' }>;
type Control = Extract<DocBlock, { type: 'contentControl' }>;

const chr = (type: string) => `<w:r><w:fldChar w:fldCharType="${type}"/></w:r>`;
const instr = (text: string) => `<w:r><w:instrText xml:space="preserve">${text}</w:instrText></w:r>`;
const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const tab = '<w:r><w:tab/></w:r>';
const pageRef = (anchor: string, page: string) => chr('begin') + instr(` PAGEREF ${anchor} \\h `) + chr('separate') + run(page) + chr('end');
const entry = (style: string, anchor: string, title: string, page: string, lead = '', trail = '') =>
  `<w:p><w:pPr><w:pStyle w:val="${style}"/></w:pPr>${lead}<w:hyperlink w:anchor="${anchor}" w:history="1">${run(title)}${tab}${pageRef(anchor, page)}</w:hyperlink>${trail}</w:p>`;
const start = chr('begin') + instr(' TOC \\o "1-3" \\h \\z \\u ') + chr('separate');
const sdt = (pr: string, content: string) => `<w:sdt><w:sdtPr>${pr}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
const docPart = '<w:docPartObj><w:docPartGallery w:val="Table of Contents"/><w:docPartUnique/></w:docPartObj>';

async function importBody(bodyXml: string) {
  const exported = await exportDocx({ title: 'T', sections: [{ ...BASE, blocks: [{ type: 'paragraph', content: [{ type: 'text', text: 'x' }] }] }] });
  const zip = await JSZip.loadAsync(exported.bytes);
  const xml = await zip.file('word/document.xml')!.async('string');
  zip.file('word/document.xml', xml.replace(/<w:p>[\s\S]*?<\/w:p>/, bodyXml.replace(/<w:p>/g, '<w:p>')));
  return importDocx(await zip.generateAsync({ type: 'uint8array' }));
}
const blocksOf = async (bodyXml: string) => (await importBody(bodyXml)).sections[0].blocks;

async function exportXml(blocks: DocBlock[], section: Partial<DocSection> = {}) {
  const out = await exportDocx({ title: 'T', sections: [{ ...BASE, ...section, blocks }] });
  const zip = await JSZip.loadAsync(out.bytes);
  return { xml: await zip.file('word/document.xml')!.async('string'), settings: await zip.file('word/settings.xml')!.async('string'), bytes: out.bytes };
}

describe('reading a table of contents', () => {
  it('reads a bare field that starts in the first entry and ends in the last as one block of entries', async () => {
    const blocks = await blocksOf(
      `<w:p ${W}><w:r><w:t>Before</w:t></w:r></w:p>` +
        entry('TOC1', '_Toc9000', 'Overview', '1', start) +
        entry('TOC2', '_Toc9001', 'Scope', '1') +
        entry('TOC1', '_Toc9002', 'Findings', '2', '', chr('end')) +
        '<w:p><w:r><w:t>After</w:t></w:r></w:p>'
    );
    expect(blocks.map((b) => b.type)).toEqual(['paragraph', 'toc', 'paragraph']);
    const toc = blocks[1] as Toc;
    expect(toc.instruction).toBe('TOC \\o "1-3" \\h \\z \\u');
    expect(toc.blocks).toHaveLength(3);
    const first = toc.blocks[0] as Paragraph;
    expect(first.namedStyleId).toBe('TOC1');
    // A link to the heading's bookmark, the tab as plain text, and the page number as a field of its own.
    expect(first.content).toMatchObject([
      { type: 'link', text: 'Overview', href: '#_Toc9000' },
      { type: 'text', text: '\t' },
      { type: 'field', kind: 'PAGEREF', instruction: 'PAGEREF _Toc9000 \\h', result: '1' },
    ]);
    expect((toc.blocks[1] as Paragraph).namedStyleId).toBe('TOC2');
  });

  it('reads the table of contents building block control Word writes, with a title before it and the end of the field in a paragraph of its own', async () => {
    const blocks = await blocksOf(
      sdt(docPart, `<w:p ${W}><w:pPr><w:pStyle w:val="TOCHeading"/></w:pPr>${run('Contents')}</w:p>` + entry('TOC1', '_Toc1', 'One', '1', start) + entry('TOC1', '_Toc2', 'Two', '2') + `<w:p>${chr('end')}</w:p>`).replace('<w:sdt>', `<w:sdt ${W}>`)
    );
    expect(blocks).toHaveLength(1);
    const control = blocks[0] as Control;
    expect(control.control).toEqual({ type: 'docPart', gallery: 'Table of Contents' });
    expect(control.blocks.map((b) => b.type)).toEqual(['paragraph', 'toc']);
    expect((control.blocks[1] as Toc).blocks).toHaveLength(2);
  });

  it('reads a table of contents whose entries are one PAGEREF field around the text, the tab and the number (as LibreOffice writes it)', async () => {
    const libre = (title: string, page: string, first = false) =>
      `<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>${first ? start : ''}<w:hyperlink w:anchor="_Toc9">${chr('begin')}${instr('PAGEREF _Toc9 \\h')}${chr('separate')}<w:r><w:t>${title}</w:t><w:tab/><w:t>${page}</w:t></w:r>${chr('end')}</w:hyperlink></w:p>`;
    const blocks = await blocksOf(libre('Programme overview', '1', true).replace('<w:p>', `<w:p ${W}>`) + libre('Budget', '2').replace('</w:p>', `${chr('end')}</w:p>`));
    const toc = blocks[0] as Toc;
    expect(toc.blocks).toHaveLength(2);
    expect((toc.blocks[0] as Paragraph).content).toMatchObject([{ type: 'field', kind: 'PAGEREF', result: 'Programme overview\t1' }]);
  });

  it('reads two tables of contents, and one that starts and ends in a single paragraph, and one with no entries', async () => {
    const one = `<w:p ${W}>${start}${run('Only entry')}${chr('end')}</w:p>`;
    const none = `<w:p>${start}${chr('end')}</w:p>`;
    const blocks = await blocksOf(one + '<w:p><w:r><w:t>between</w:t></w:r></w:p>' + none);
    expect(blocks.map((b) => b.type)).toEqual(['toc', 'paragraph', 'toc']);
    expect((blocks[0] as Toc).blocks).toHaveLength(1);
    expect((blocks[2] as Toc).blocks).toEqual([]);
  });

  it('keeps the instruction as written, with spaces tidied and the pieces of it joined', async () => {
    const blocks = await blocksOf(`<w:p ${W}>${chr('begin')}${instr(' TOC ')}${instr('\\o "2-4"')}${instr('  \\n ')}${chr('separate')}${run('x')}${chr('end')}</w:p>`);
    expect((blocks[0] as Toc).instruction).toBe('TOC \\o "2-4" \\n');
  });

  it('makes a paragraph of each entry, whatever its style: a numbered or heading style does not make it a list item or a heading', async () => {
    const numbered = `<w:p ${W}><w:pPr><w:pStyle w:val="Heading1"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>${start}${run('Numbered')}${chr('end')}</w:p>`;
    const toc = (await blocksOf(numbered))[0] as Toc;
    expect(toc.blocks[0]).toMatchObject({ type: 'paragraph', content: [{ type: 'text', text: 'Numbered' }] });
    expect(toc.blocks[0]).not.toHaveProperty('listType');
    expect(toc.blocks[0]).not.toHaveProperty('level');
  });

  it('splits a link around the tab and the number into the link and plain text, so the tab can be drawn as a leader', async () => {
    const link = `<w:p ${W}>${start}<w:hyperlink w:anchor="_Toc3"><w:r><w:t>Title</w:t><w:tab/><w:t>7</w:t></w:r></w:hyperlink>${chr('end')}</w:p>`;
    const toc = (await blocksOf(link))[0] as Toc;
    expect((toc.blocks[0] as Paragraph).content).toMatchObject([
      { type: 'link', text: 'Title', href: '#_Toc3' },
      { type: 'text', text: '\t' },
      { type: 'link', text: '7', href: '#_Toc3' },
    ]);
  });

  it('keeps the named styles the entries use', async () => {
    const result = await importBody(entry('TOC1', '_Toc1', 'One', '1', start, chr('end')).replace('<w:p>', `<w:p ${W}>`));
    expect(result.sections[0].blocks[0].type).toBe('toc');
    // The styles part of the test file does not define TOC1, so none is reported.
    expect(Object.keys(result.namedStyles ?? {})).not.toContain('Missing');
  });

  it('leaves a table of contents as static paragraphs when it cannot be one block', async () => {
    const entries = start + run('entry');
    const withTable = await blocksOf(`<w:p ${W}>${entries}</w:p><w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl><w:p>${chr('end')}</w:p>`);
    expect(withTable.some((b) => b.type === 'toc')).toBe(false);
    const withBreak = await blocksOf(`<w:p ${W}>${entries}</w:p><w:p><w:pPr><w:sectPr/></w:pPr></w:p><w:p>${chr('end')}</w:p>`);
    expect(withBreak.some((b) => b.type === 'toc')).toBe(false);
    const noEnd = await blocksOf(`<w:p ${W}>${entries}</w:p>`);
    expect(noEnd.some((b) => b.type === 'toc')).toBe(false);
    expect((noEnd[0] as Paragraph).content).toEqual([{ type: 'text', text: 'entry' }]);
  });

  it('is not misled by other fields: one that is not a table of contents, and a table of contents inside a field', async () => {
    const other = `<w:p ${W}>${chr('begin')}${instr(' INDEX ')}${chr('separate')}${run('index')}${chr('end')}</w:p>`;
    expect((await blocksOf(other)).some((b) => b.type === 'toc')).toBe(false);
    const nested = `<w:p ${W}>${chr('begin')}${instr(' IF 1 = 1 ')}${chr('separate')}${chr('begin')}${instr(' TOC ')}${chr('separate')}${run('inner')}${chr('end')}${chr('end')}</w:p>`;
    expect((await blocksOf(nested)).some((b) => b.type === 'toc')).toBe(false);
  });

  it('reads the tables of contents of the fixtures and writes them back so that reading them again gives the same blocks', async () => {
    const fixtures = path.resolve(__dirname, '../../../../fixtures/docx');
    for (const name of ['notes-toc-fields', 'libreoffice-complex', 'long-document', 'kitchen-sink', 'word-compat-kitchen']) {
      const first = await importDocx(new Uint8Array(readFileSync(path.join(fixtures, `${name}.docx`))));
      const tocs = JSON.stringify(first.sections.map((s) => s.blocks)).match(/"type":"toc"/g) ?? [];
      expect(tocs.length).toBeGreaterThan(0);
      const out = await exportDocx({ title: 'T', sections: first.sections, comments: first.comments, namedStyles: first.namedStyles });
      const second = await importDocx(out.bytes);
      const tocsOf = (sections: DocSection[]) => {
        const found: Toc[] = [];
        const visit = (b: DocBlock) => {
          if (b.type === 'toc') found.push(b);
          else if (b.type === 'contentControl') b.blocks.forEach(visit);
        };
        sections.forEach((s) => s.blocks.forEach(visit));
        return found;
      };
      expect(tocsOf(second.sections)).toEqual(tocsOf(first.sections));
    }
  });
});

describe('toc-field helpers', () => {
  const body = (xml: string) => childrenOf(find(parseXml(`<w:body ${W}>${xml}</w:body>`), 'w:body')!);

  it('finds the field and says which runs are its own markup', () => {
    const nodes = body(entry('TOC1', '_Toc1', 'A', '1', start) + entry('TOC1', '_Toc2', 'B', '2', '', chr('end')));
    const [range] = findTocRanges(nodes);
    expect(range).toMatchObject({ start: 0, end: 1, instruction: 'TOC \\o "1-3" \\h \\z \\u' });
    // Begin, instruction, separate and end: four runs, not the entries' PAGEREF fields.
    expect(range.markers.size).toBe(4);
    const stripped = withoutMarkers(nodes[0], range.markers);
    expect(stripped).not.toBe(nodes[0]);
    expect(withoutMarkers(nodes[0], new Set())).toBe(nodes[0]);
    expect(isEmptyParagraph(withoutMarkers(body(`<w:p><w:pPr/>${chr('end')}</w:p>`)[0], range.markers))).toBe(false);
  });

  it('finds nothing without a table of contents field, and ignores bookmarks between its paragraphs', () => {
    expect(findTocRanges(body('<w:p><w:r><w:t>x</w:t></w:r></w:p>'))).toEqual([]);
    const withBookmark = body(`<w:p>${start}${run('a')}</w:p><w:bookmarkStart w:id="1" w:name="b"/><w:bookmarkEnd w:id="1"/><w:p>${chr('end')}</w:p>`);
    expect(findTocRanges(withBookmark)).toHaveLength(1);
  });

  it('lists the places a table of contents can be: the body and the content controls that are kept', () => {
    const nodes = body(`<w:p><w:r><w:t>x</w:t></w:r></w:p>` + sdt(docPart, '<w:p/>') + sdt('<w:group/>', '<w:p><w:r><w:t>unwrapped</w:t></w:r></w:p>') + '<w:customXml w:element="e"><w:p/></w:customXml>');
    const lists = tocLists(nodes);
    expect(lists).toHaveLength(2);
    // The body: the paragraph, the kept control, the content of the unsupported one and of the custom XML.
    const names = (list: typeof lists[0]) => list.map((n) => tagOf(n));
    expect(names(lists[1])).toEqual(['w:p', 'w:sdt', 'w:p', 'w:p']);
    expect(names(lists[0])).toEqual(['w:p']);
  });
});

describe('writing a table of contents', () => {
  const link = (text: string, href: string): DocInline => ({ type: 'link', text, href });
  const text = (t: string): DocInline => ({ type: 'text', text: t });
  const editorEntries = (): DocBlock[] => [
    { type: 'paragraph', namedStyleId: 'TOC1', content: [link('Chapter one', '#_Toc100000001'), text('\t1')] },
    { type: 'paragraph', namedStyleId: 'TOC2', leftIndentPt: 18, content: [link('Section', '#_Toc100000002'), text('\t2')] },
  ];

  it('writes the field across the entries: starting in the first, ending in a paragraph of its own', async () => {
    const { xml } = await exportXml([{ type: 'toc', instruction: 'TOC \\o "1-3" \\h \\z \\u', blocks: editorEntries() }]);
    expect(xml).not.toContain('<w:sdt>');
    const content = xml.slice(xml.indexOf('<w:body>'), xml.indexOf('<w:sectPr>'));
    const paragraphs = content.match(/<w:p>[\s\S]*?<\/w:p>/g)!;
    expect(paragraphs).toHaveLength(3);
    expect(paragraphs[0]).toMatch(/<w:fldChar w:fldCharType="begin"\/>.*<w:instrText xml:space="preserve"> TOC \\o &quot;1-3&quot; \\h \\z \\u <\/w:instrText>.*<w:fldChar w:fldCharType="separate"\/>.*<w:hyperlink [^>]*w:anchor="_Toc100000001"/s);
    expect(paragraphs[1]).not.toContain('w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC');
    expect(paragraphs[2]).toBe('<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>');
    expect(content.match(/w:fldCharType="end"/g)!.length).toBe(content.match(/w:fldCharType="begin"/g)!.length);
  });

  it('gives the number after the tab as a PAGEREF field of the heading’s bookmark, so Word can refresh it, and a right tab with a dot leader at the right margin', async () => {
    const { xml } = await exportXml([{ type: 'toc', instruction: 'TOC \\o "1-3" \\h', blocks: editorEntries() }]);
    expect(xml).toContain('<w:instrText xml:space="preserve"> PAGEREF _Toc100000001 \\h </w:instrText>');
    expect(xml).toContain('<w:instrText xml:space="preserve"> PAGEREF _Toc100000002 \\h </w:instrText>');
    // 816 - 2 * 96 = 624 px = 9360 twips.
    expect(xml).toMatch(/<w:tab w:val="right" w:pos="9360" w:leader="dot"\/>/);
    expect(xml).not.toContain('updateFields');
  });

  it('puts the right tab at the right margin of the section (a landscape page, other margins)', async () => {
    const { xml } = await exportXml([{ type: 'toc', instruction: 'TOC', blocks: editorEntries() }], { pageWidthPx: 1056, pageHeightPx: 816, orientation: 'landscape', marginLeftPx: 48, marginRightPx: 48 });
    expect(xml).toContain('w:pos="14400"');
  });

  it('leaves entries as they are when they have no link to a bookmark, a number after a tab, or fields already', async () => {
    const already: DocBlock = { type: 'paragraph', content: [link('A', '#_Toc1'), text('\t'), { type: 'field', instruction: 'PAGEREF _Toc1 \\h', kind: 'PAGEREF', result: '5' }] };
    const { xml } = await exportXml([
      {
        type: 'toc',
        instruction: 'TOC \\o "1-3"',
        blocks: [
          { type: 'paragraph', content: [text('Plain\t3')] },
          { type: 'paragraph', content: [link('NoNumber', '#_Toc2'), text('\t')] },
          { type: 'paragraph', content: [link('External', 'https://example.com'), text('\t4')] },
          already,
        ],
      },
    ]);
    expect(xml.match(/PAGEREF/g)).toHaveLength(1);
    expect(xml).toContain('Plain');
  });

  it('writes a table of contents with no entries as a field of its own', async () => {
    const { xml } = await exportXml([{ type: 'toc', instruction: 'TOC \\o "1-3"', blocks: [] }]);
    expect(xml).toMatch(/<w:instrText[^>]*> TOC \\o &quot;1-3&quot; <\/w:instrText>/);
    expect(xml).toContain('<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>');
  });

  it('writes the field as paragraphs in whatever control surrounds it, and adds no control of its own', async () => {
    const toc: DocBlock = { type: 'toc', instruction: 'TOC', blocks: editorEntries() };
    const inside = await exportXml([{ type: 'contentControl', control: { type: 'docPart', gallery: 'Table of Contents' }, blocks: [{ type: 'paragraph', content: [text('Contents')] }, toc] }]);
    expect(inside.xml.match(/<w:docPartGallery/g)).toHaveLength(1);
    expect(inside.xml.match(/<w:sdt>/g)).toHaveLength(1);
    const other = await exportXml([{ type: 'contentControl', control: { type: 'richText', alias: 'Wrapper' }, blocks: [toc] }]);
    expect(other.xml.match(/<w:docPartGallery/g)).toBeNull();
    expect(other.xml.match(/<w:sdt>/g)).toHaveLength(1);
  });

  it('writes a table of contents in a table cell as paragraphs of the cell', async () => {
    const { xml } = await exportXml([{ type: 'table', rows: [{ cells: [{ content: [], blocks: [{ type: 'toc', instruction: 'TOC', blocks: editorEntries() }] }] }] }]);
    expect(xml.match(/<w:instrText[^>]*> TOC <\/w:instrText>/g)).toHaveLength(1);
  });

  it('writes tables of contents in several sections, each with the tab of its own page', async () => {
    const out = await exportDocx({
      title: 'T',
      sections: [
        { ...BASE, blocks: [{ type: 'toc', instruction: 'TOC', blocks: editorEntries() }] },
        { ...BASE, marginRightPx: 144, blocks: [{ type: 'toc', instruction: 'TOC', blocks: editorEntries() }] },
      ],
    });
    const xml = await (await JSZip.loadAsync(out.bytes)).file('word/document.xml')!.async('string');
    expect(xml).toContain('w:pos="9360"');
    expect(xml).toContain('w:pos="8640"');
  });

  it('round-trips an entry made in the editor: the number comes back as the PAGEREF field it was written as', async () => {
    const { bytes } = await exportXml([{ type: 'toc', instruction: 'TOC \\o "1-3" \\h \\z \\u', blocks: editorEntries() }]);
    const toc = (await importDocx(bytes)).sections[0].blocks[0] as Toc;
    expect(toc.type).toBe('toc');
    expect(toc.blocks).toHaveLength(2);
    expect((toc.blocks[0] as Paragraph).content).toMatchObject([
      { type: 'link', text: 'Chapter one', href: '#_Toc100000001' },
      { type: 'text', text: '\t' },
      { type: 'field', kind: 'PAGEREF', result: '1' },
    ]);
  });
});
