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
const r = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const start = (name: string, id = 1) => `<w:bookmarkStart w:id="${id}" w:name="${name}"/>`;
const end = (id = 1) => `<w:bookmarkEnd w:id="${id}"/>`;
const para = (inner: string) => `<w:p xmlns:w="x">${inner}</w:p>`;
type Paragraph = Extract<DocBlock, { type: 'paragraph' }>;

async function importBody(bodyXml: string) {
  const exported = await exportDocx({ title: 'T', sections: [{ ...BASE, blocks: [{ type: 'paragraph', content: [{ type: 'text', text: 'x' }] }] }] });
  const zip = await JSZip.loadAsync(exported.bytes);
  const xml = (await zip.file('word/document.xml')!.async('string')).replace(/<w:p>[\s\S]*?<\/w:p>/, bodyXml);
  zip.file('word/document.xml', xml);
  return importDocx(await zip.generateAsync({ type: 'uint8array' }));
}
const blocksOf = async (bodyXml: string) => (await importBody(bodyXml)).sections[0].blocks;
const contentOf = async (bodyXml: string, block = 0): Promise<DocInline[]> => ((await blocksOf(bodyXml))[block] as Paragraph).content;

async function roundTrip(blocks: DocBlock[]) {
  const exported = await exportDocx({ title: 'T', sections: [{ ...BASE, blocks }] });
  const zip = await JSZip.loadAsync(exported.bytes);
  const xml = await zip.file('word/document.xml')!.async('string');
  const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
  return { xml, rels, imported: await importDocx(exported.bytes) };
}

describe('bookmarks', () => {
  it('keeps a bookmark as a point where it starts, including _Toc and _Ref names', async () => {
    const content = await contentOf(para(start('_Toc1') + r('Intro') + end() + start('_Ref2', 2) + r(' tail') + end(2)));
    expect(content).toEqual([
      { type: 'bookmark', name: '_Toc1' },
      { type: 'text', text: 'Intro' },
      { type: 'bookmark', name: '_Ref2' },
      { type: 'text', text: ' tail' },
    ]);
  });

  it("drops Word's _GoBack cursor marker and unnamed bookmarks", async () => {
    const content = await contentOf(para(start('_GoBack') + end() + '<w:bookmarkStart w:id="3"/>' + r('a')));
    expect(content).toEqual([{ type: 'text', text: 'a' }]);
  });

  it('finds bookmarks inside a hyperlink and a bookmark that spans paragraphs (start only)', async () => {
    const blocks = await blocksOf(
      para(`<w:hyperlink w:anchor="x">${start('inner')}${r('go')}</w:hyperlink>`) + para(start('spans') + r('one')) + para(r('two') + end())
    );
    expect((blocks[0] as Paragraph).content).toEqual([
      { type: 'bookmark', name: 'inner' },
      { type: 'link', text: 'go', href: '#x' },
    ]);
    expect((blocks[1] as Paragraph).content).toEqual([{ type: 'bookmark', name: 'spans' }, { type: 'text', text: 'one' }]);
    expect((blocks[2] as Paragraph).content).toEqual([{ type: 'text', text: 'two' }]);
  });

  it('moves a bookmark that starts between paragraphs to the start of the next paragraph', async () => {
    const blocks = await blocksOf(`<w:p xmlns:w="x">${r('a')}</w:p>${start('between')}<w:p xmlns:w="x">${r('b')}</w:p>`);
    expect((blocks[0] as Paragraph).content).toEqual([{ type: 'text', text: 'a' }]);
    expect((blocks[1] as Paragraph).content).toEqual([{ type: 'bookmark', name: 'between' }, { type: 'text', text: 'b' }]);
  });

  it('does not let a bookmark-only paragraph turn into a page break or hide a table cell\'s content', async () => {
    const blocks = await blocksOf(
      para(start('a') + '<w:r><w:br w:type="page"/></w:r>') +
        `<w:tbl xmlns:w="x"><w:tr><w:tc><w:p>${start('c', 5)}</w:p><w:p>${r('cell text')}</w:p></w:tc></w:tr></w:tbl>`
    );
    expect(blocks[0]).toEqual({ type: 'pageBreak' });
    const cell = (blocks[1] as Extract<DocBlock, { type: 'table' }>).rows[0].cells[0];
    expect(cell.content).toEqual([{ type: 'bookmark', name: 'c' }, { type: 'text', text: 'cell text' }]);
  });

  it('drops a bookmark in a header without a warning', async () => {
    const exported = await exportDocx({
      title: 'T',
      sections: [{ ...BASE, blocks: [{ type: 'paragraph', content: [{ type: 'text', text: 'x' }] }], header: { paragraphs: [{ content: [{ type: 'text', text: 'H' }] }] } }],
    });
    const zip = await JSZip.loadAsync(exported.bytes);
    const name = Object.keys(zip.files).find((f) => /word\/header\d*\.xml$/.test(f))!;
    const xml = (await zip.file(name)!.async('string')).replace(/<w:p>/, '<w:p><w:bookmarkStart w:id="9" w:name="hb"/>');
    zip.file(name, xml);
    const { sections, warnings } = await importDocx(await zip.generateAsync({ type: 'uint8array' }));
    expect(sections[0].header?.paragraphs[0].content).toEqual([{ type: 'text', text: 'H' }]);
    expect(warnings).toEqual([]);
  });
});

describe('internal links', () => {
  it('turns a hyperlink with an anchor into a #name link', async () => {
    expect(await contentOf(para(`<w:hyperlink w:anchor="_Toc1">${r('Intro')}</w:hyperlink>`))).toEqual([{ type: 'link', text: 'Intro', href: '#_Toc1' }]);
  });

  it('keeps an external target and its anchor together', async () => {
    const exported = await exportDocx({
      title: 'T',
      sections: [{ ...BASE, blocks: [{ type: 'paragraph', content: [{ type: 'link', text: 'l', href: 'https://e.com/p' }] }] }],
    });
    const zip = await JSZip.loadAsync(exported.bytes);
    const xml = (await zip.file('word/document.xml')!.async('string')).replace('<w:hyperlink ', '<w:hyperlink w:anchor="sec" ');
    zip.file('word/document.xml', xml);
    const { sections } = await importDocx(await zip.generateAsync({ type: 'uint8array' }));
    expect((sections[0].blocks[0] as Paragraph).content).toMatchObject([{ type: 'link', href: 'https://e.com/p#sec' }]);
  });
});

describe('bookmarks and internal links in the export', () => {
  const content: DocInline[] = [
    { type: 'bookmark', name: '_Toc99' },
    { type: 'text', text: 'Heading' },
    { type: 'link', text: 'jump', href: '#_Toc99', bold: true },
    { type: 'link', text: 'web', href: 'https://example.com/#frag' },
  ];

  it('writes a bookmark and an anchor hyperlink, and an external link with a relationship', async () => {
    const { xml, rels } = await roundTrip([{ type: 'paragraph', content }]);
    expect(xml).toMatch(/<w:bookmarkStart [^>]*w:name="_Toc99"[^>]*\/><w:bookmarkEnd [^>]*\/>/);
    expect(xml).toMatch(/<w:hyperlink [^>]*w:anchor="_Toc99"/);
    const internal = xml.match(/<w:hyperlink [^>]*w:anchor="_Toc99"[^>]*>[\s\S]*?<\/w:hyperlink>/)![0];
    expect(internal).not.toContain('Hyperlink'); // an internal link keeps the look of its text
    // The writer library puts the fragment in w:anchor and the rest in the relationship; the reader joins them again.
    expect(rels).toContain('Target="https://example.com/"');
    expect(rels).not.toContain('Target="#_Toc99"');
  });

  it('round-trips bookmarks and internal links, with unique bookmark ids', async () => {
    const { xml, imported } = await roundTrip([
      { type: 'paragraph', content },
      { type: 'heading', level: 1, content: [{ type: 'bookmark', name: 'b2' }, { type: 'text', text: 'H' }] },
    ]);
    // (The writer gives every link Word's hyperlink look, which comes back as run formatting.)
    expect((imported.sections[0].blocks[0] as Paragraph).content).toMatchObject([
      { type: 'bookmark', name: '_Toc99' },
      { type: 'text', text: 'Heading' },
      { type: 'link', text: 'jump', href: '#_Toc99', bold: true },
      { type: 'link', text: 'web', href: 'https://example.com/#frag' },
    ]);
    expect((imported.sections[0].blocks[1] as { content: DocInline[] }).content[0]).toEqual({ type: 'bookmark', name: 'b2' });
    const ids = [...xml.matchAll(/<w:bookmarkStart [^>]*w:id="(\d+)"/g)].map((m) => m[1]);
    expect(new Set(ids).size).toBe(2);
  });

  it('skips a bookmark with no name and keeps a lone # as an external link', async () => {
    const { xml } = await roundTrip([{ type: 'paragraph', content: [{ type: 'bookmark', name: '' }, { type: 'link', text: 'x', href: '#' }] }]);
    expect(xml).not.toContain('<w:bookmarkStart');
    expect(xml).not.toContain('w:anchor');
  });
});
