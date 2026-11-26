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
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
type Paragraph = Extract<DocBlock, { type: 'paragraph' }>;
type Table = Extract<DocBlock, { type: 'table' }>;

const t = (text: string) => `<w:t xml:space="preserve">${text}</w:t>`;
const r = (inner: string, props = '') => `<w:r>${props && `<w:rPr>${props}</w:rPr>`}${inner}</w:r>`;
const who = (id: string, author = 'Ada') => `w:id="${id}" w:author="${author}" w:date="2024-05-06T07:08:09Z"`;

async function importBody(bodyXml: string) {
  const exported = await exportDocx({ title: 'T', sections: [{ ...BASE, blocks: [{ type: 'paragraph', content: [{ type: 'text', text: 'x' }] }] }] });
  const zip = await JSZip.loadAsync(exported.bytes);
  const xml = await zip.file('word/document.xml')!.async('string');
  zip.file('word/document.xml', xml.replace(/<w:p>[\s\S]*?<\/w:p>/, bodyXml));
  return importDocx(await zip.generateAsync({ type: 'uint8array' }));
}
const blocksOf = async (bodyXml: string) => (await importBody(bodyXml)).sections[0].blocks;
const paragraph = (inner: string, pPr = '') => `<w:p ${W}>${pPr && `<w:pPr>${pPr}</w:pPr>`}${inner}</w:p>`;

async function exportXml(blocks: DocBlock[]) {
  const out = await exportDocx({ title: 'T', sections: [{ ...BASE, blocks }] });
  const zip = await JSZip.loadAsync(out.bytes);
  return { xml: await zip.file('word/document.xml')!.async('string'), bytes: out.bytes };
}
const roundTrip = async (blocks: DocBlock[]) => (await importDocx((await exportXml(blocks)).bytes)).sections[0].blocks;

const META = { id: '5', author: 'Ada', date: '2024-05-06T07:08:09Z' };

describe('reading moves', () => {
  it('reads the two ends of a move as a deletion and an insertion that share the name of the move', async () => {
    const blocks = await blocksOf(
      paragraph(
        r(t('a ')) +
          `<w:moveFromRangeStart w:id="1" w:name="move9"/><w:moveFrom ${who('2')}>${r('<w:delText>moved</w:delText>')}</w:moveFrom><w:moveFromRangeEnd w:id="1"/>` +
          `<w:moveToRangeStart w:id="3" w:name="move9"/><w:moveTo ${who('4')}>${r(t('moved'))}</w:moveTo><w:moveToRangeEnd w:id="3"/>` +
          `<w:ins ${who('6')}>${r(t('plain'))}</w:ins>`
      )
    );
    expect((blocks[0] as Paragraph).content).toMatchObject([
      { text: 'a ' },
      { text: 'moved', revision: { deletion: { id: '2', author: 'Ada', moveId: 'move9' } } },
      { text: 'moved', revision: { insertion: { id: '4', moveId: 'move9' } } },
      { text: 'plain', revision: { insertion: { id: '6' } } },
    ]);
    expect(JSON.stringify((blocks[0] as Paragraph).content[3])).not.toContain('moveId');
  });

  it('keeps a move whose range markers are missing, without a name', async () => {
    const blocks = await blocksOf(paragraph(`<w:moveTo ${who('4')}>${r(t('x'))}</w:moveTo>`));
    expect((blocks[0] as Paragraph).content[0]).toEqual({ type: 'text', text: 'x', revision: { insertion: { id: '4', author: 'Ada', date: META.date } } });
  });

  it('warns once when a moved run has no range name because the range starts in another paragraph', async () => {
    const result = await importBody(
      paragraph(`<w:moveFromRangeStart w:id="1" w:name="move9"/>${r(t('a'))}`) +
        paragraph(`<w:moveFrom ${who('2')}>${r('<w:delText>b</w:delText>')}</w:moveFrom><w:moveFromRangeEnd w:id="1"/>`) +
        paragraph(`<w:moveTo ${who('3')}>${r(t('c'))}</w:moveTo>`)
    );
    expect(((result.sections[0].blocks[1] as Paragraph).content[0] as { revision: unknown }).revision).toEqual({ deletion: { id: '2', author: 'Ada', date: META.date } });
    expect(result.warnings.filter((w) => w.includes('move range starts in another paragraph'))).toHaveLength(1);
  });

  it('reads the mark of a moved paragraph as a deletion or an insertion of the mark', async () => {
    const blocks = await blocksOf(paragraph(r(t('a')), `<w:rPr><w:moveFrom ${who('7')}/></w:rPr>`) + paragraph(r(t('b')), `<w:rPr><w:moveTo ${who('8')}/></w:rPr>`));
    expect(blocks[0]).toMatchObject({ revision: { deletion: { id: '7' } } });
    expect(blocks[1]).toMatchObject({ revision: { insertion: { id: '8' } } });
  });
});

describe('reading formatting changes', () => {
  it('keeps the formatting a run had before, as effective formatting, and the formatting it has now', async () => {
    const blocks = await blocksOf(
      paragraph(
        r(t('now bold'), `<w:b/><w:rPrChange ${who('3')}><w:rPr><w:i/><w:color w:val="C00000"/></w:rPr></w:rPrChange>`) +
          r(t('plain now'), `<w:rPrChange ${who('4', 'Bo')}><w:rPr><w:b/></w:rPr></w:rPrChange>`)
      )
    );
    const [first, second] = (blocks[0] as Paragraph).content as Array<Extract<DocInline, { type: 'text' }>>;
    expect(first).toMatchObject({ bold: true, revision: { formatChange: { id: '3', author: 'Ada', date: META.date, previous: { italic: true, color: '#c00000' } } } });
    expect(first.revision!.formatChange!.previous.bold).toBeFalsy();
    expect(second.bold).toBeFalsy();
    expect(second.revision!.formatChange).toMatchObject({ id: '4', author: 'Bo', previous: { bold: true } });
  });

  it('combines a formatting change with an insertion, and keeps the character style the run had', async () => {
    const blocks = await blocksOf(paragraph(`<w:ins ${who('1')}>${r(t('x'), `<w:rPrChange ${who('2')}><w:rPr/></w:rPrChange>`)}</w:ins>`));
    expect((blocks[0] as Paragraph).content[0]).toMatchObject({ revision: { insertion: { id: '1' }, formatChange: { id: '2', previous: {} } } });
  });

  it('says so when a formatting change is on a run that has no text left (a field result, an image)', async () => {
    const change = `<w:rPrChange ${who('2')}><w:rPr><w:b/></w:rPr></w:rPrChange>`;
    const fieldResult =
      r('<w:fldChar w:fldCharType="begin"/>') +
      r('<w:instrText> REF a </w:instrText>') +
      r('<w:fldChar w:fldCharType="separate"/>') +
      r(t('3'), `<w:i/>${change}`) +
      r('<w:fldChar w:fldCharType="end"/>');
    const inField = await importBody(paragraph(fieldResult));
    const lost = /formatting change of a run without text of its own/;
    expect(inField.warnings.filter((w) => lost.test(w))).toHaveLength(1);
    expect(JSON.stringify(inField.sections[0].blocks)).not.toContain('formatChange');
    const image = await importBody(paragraph(r('<w:drawing/>', change)));
    expect(image.warnings.filter((w) => lost.test(w))).toHaveLength(1);
    // A run with text keeps its change and does not warn.
    const plain = await importBody(paragraph(r(t('x'), change)));
    expect(plain.warnings.filter((w) => lost.test(w))).toHaveLength(0);
    expect(JSON.stringify(plain.sections[0].blocks)).toContain('formatChange');
  });

  it('reads the change of a paragraph with its old formatting, including the heading it was', async () => {
    const blocks = await blocksOf(
      paragraph(r(t('centred')), `<w:jc w:val="center"/><w:pPrChange ${who('9')}><w:pPr><w:ind w:left="720"/></w:pPr></w:pPrChange>`) +
        paragraph(r(t('title')), `<w:pStyle w:val="Heading1"/><w:pPrChange ${who('10')}><w:pPr/></w:pPrChange>`)
    );
    expect(blocks[0]).toMatchObject({ type: 'paragraph', alignment: 'center', revision: { formatChange: { id: '9', previous: { leftIndentPt: 36 } } } });
    expect((blocks[0] as Paragraph).revision!.formatChange!.previous.alignment).toBeUndefined();
    expect(blocks[1]).toMatchObject({ type: 'heading', revision: { formatChange: { id: '10' } } });
  });
});

describe('reading changes of paragraph marks and table rows', () => {
  it('reads an inserted and a deleted paragraph mark', async () => {
    const blocks = await blocksOf(paragraph(r(t('a')), `<w:rPr><w:ins ${who('1')}/></w:rPr>`) + paragraph(r(t('b')), `<w:rPr><w:del ${who('2', 'Bo')}/></w:rPr>`));
    expect(blocks[0]).toMatchObject({ revision: { insertion: { id: '1', author: 'Ada' } } });
    expect(blocks[1]).toMatchObject({ revision: { deletion: { id: '2', author: 'Bo' } } });
  });

  it('reads inserted and deleted rows', async () => {
    const row = (pr: string) => `<w:tr><w:trPr>${pr}</w:trPr><w:tc><w:p>${r(t('c'))}</w:p></w:tc></w:tr>`;
    const blocks = await blocksOf(`<w:tbl ${W}>${row(`<w:ins ${who('1')}/>`)}${row(`<w:del ${who('2', 'Bo')}/>`)}${row('')}</w:tbl>`);
    const rows = (blocks[0] as Table).rows;
    expect(rows[0].revision).toEqual({ insertion: { id: '1', author: 'Ada', date: META.date } });
    expect(rows[1].revision).toEqual({ deletion: { id: '2', author: 'Bo', date: META.date } });
    expect(rows[2].revision).toBeUndefined();
  });

  it('numbers changes without an id', async () => {
    const blocks = await blocksOf(paragraph(r(t('a')), '<w:rPr><w:ins/></w:rPr>') + paragraph(r(t('b')), '<w:rPr><w:del/></w:rPr>'));
    expect(blocks[0]).toMatchObject({ revision: { insertion: { id: 'r1', author: 'Unknown' } } });
    expect(blocks[1]).toMatchObject({ revision: { deletion: { id: 'r2' } } });
  });
});

describe('writing', () => {
  it('writes the two ends of a move with their range markers and reads them back', async () => {
    const content: DocInline[] = [
      { type: 'text', text: 'from', revision: { deletion: { id: '2', author: 'Ada', date: META.date, moveId: 'm1' } } },
      { type: 'text', text: 'to', revision: { insertion: { id: '4', author: 'Ada', moveId: 'm1' } } },
    ];
    const { xml } = await exportXml([{ type: 'paragraph', content }]);
    expect(xml).toMatch(/<w:moveFromRangeStart [^>]*w:name="m1"[^>]*\/><w:moveFrom [^>]*w:author="Ada"[^>]*><w:r><w:delText[^>]*>from<\/w:delText><\/w:r><\/w:moveFrom><w:moveFromRangeEnd [^>]*\/>/);
    expect(xml).toMatch(/<w:moveToRangeStart [^>]*w:name="m1"[^>]*\/><w:moveTo [^>]*><w:r><w:t[^>]*>to<\/w:t><\/w:r><\/w:moveTo><w:moveToRangeEnd [^>]*\/>/);
    const ids = [...xml.matchAll(/<w:(?:moveFrom|moveTo|ins|del)(?:RangeStart)? w:id="(\d+)"/g)].map((m) => m[1]);
    expect(new Set(ids).size).toBe(ids.length);
    expect(((await roundTrip([{ type: 'paragraph', content }]))[0] as Paragraph).content).toMatchObject([
      { text: 'from', revision: { deletion: { id: '2', moveId: 'm1' } } },
      { text: 'to', revision: { insertion: { id: '4', moveId: 'm1' } } },
    ]);
  });

  it('writes a move that something splits as two ranges of the same name, and reads both back as one move', async () => {
    const move = { deletion: { id: '2', author: 'Ada', moveId: 'm1' } };
    const content: DocInline[] = [
      { type: 'text', text: 'a', bold: true, revision: move },
      { type: 'text', text: 'b', italic: true, revision: move },
      { type: 'bookmark', name: 'here' },
      { type: 'text', text: 'c', revision: move },
    ];
    const { xml } = await exportXml([{ type: 'paragraph', content }]);
    // Runs of different formatting stay in one range; the bookmark splits it into two (as Word does for a bookmark in a move).
    expect([...xml.matchAll(/<w:moveFromRangeStart [^>]*w:name="m1"/g)]).toHaveLength(2);
    const ids = [...xml.matchAll(/<w:moveFromRange(?:Start|End) w:id="(\d+)"/g)].map((m) => m[1]);
    expect(new Set(ids).size).toBe(2);
    const back = ((await roundTrip([{ type: 'paragraph', content }]))[0] as Paragraph).content.filter((c) => c.type === 'text');
    expect(back).toMatchObject([
      { text: 'a', revision: { deletion: { id: '2', moveId: 'm1' } } },
      { text: 'b', revision: { deletion: { id: '2', moveId: 'm1' } } },
      { text: 'c', revision: { deletion: { id: '2', moveId: 'm1' } } },
    ]);
  });

  it('writes a formatting change of a run as w:rPrChange with the old formatting, and reads it back', async () => {
    const content: DocInline[] = [{ type: 'text', text: 'now', bold: true, revision: { formatChange: { ...META, previous: { italic: true, fontSize: 14, color: '#c00000' } } } }];
    const { xml } = await exportXml([{ type: 'paragraph', content }]);
    expect(xml).toMatch(/<w:rPr><w:b\/>[\s\S]*<w:rPrChange [^>]*w:author="Ada"[^>]*><w:rPr><w:i\/>[\s\S]*<w:sz w:val="28"\/>[\s\S]*<\/w:rPrChange><\/w:rPr>/);
    expect(((await roundTrip([{ type: 'paragraph', content }]))[0] as Paragraph).content[0]).toMatchObject({
      bold: true,
      revision: { formatChange: { id: '5', author: 'Ada', previous: { italic: true, fontSize: 14, color: '#c00000' } } },
    });
  });

  it('writes an inserted and deleted paragraph mark and the formatting before a paragraph change', async () => {
    const blocks: DocBlock[] = [
      { type: 'paragraph', content: [{ type: 'text', text: 'a' }], revision: { insertion: { id: '1', author: 'Ada' } } },
      { type: 'paragraph', content: [{ type: 'text', text: 'b' }], revision: { deletion: { id: '2', author: 'Bo', date: META.date } } },
      { type: 'paragraph', alignment: 'center', content: [{ type: 'text', text: 'c' }], revision: { formatChange: { ...META, previous: { alignment: 'right', leftIndentPt: 36 } } } },
      { type: 'heading', level: 2, content: [{ type: 'text', text: 'd' }], revision: { formatChange: { id: '6', author: 'Cy', previous: {} } } },
      { type: 'listItem', listType: 'bullet', level: 0, content: [{ type: 'text', text: 'e' }], revision: { insertion: { id: '7', author: 'Di' } } },
    ];
    const { xml } = await exportXml(blocks);
    expect(xml).toMatch(/<w:p><w:pPr>[\s\S]*?<w:rPr><w:ins [^>]*w:author="Ada"[^>]*\/><\/w:rPr>/);
    expect(xml).toMatch(/<w:rPr><w:del [^>]*w:author="Bo"[^>]*\/><\/w:rPr>/);
    expect(xml).toMatch(/<w:pPrChange [^>]*w:author="Ada"[^>]*><w:pPr>[\s\S]*?<w:ind w:left="720"\/>[\s\S]*?<w:jc w:val="right"\/>[\s\S]*?<\/w:pPrChange>/);
    const back = await roundTrip(blocks);
    expect(back[0]).toMatchObject({ revision: { insertion: { id: '1', author: 'Ada' } } });
    expect(back[1]).toMatchObject({ revision: { deletion: { id: '2', author: 'Bo' } } });
    expect(back[2]).toMatchObject({ alignment: 'center', revision: { formatChange: { id: '5', previous: { alignment: 'right', leftIndentPt: 36 } } } });
    expect(back[3]).toMatchObject({ type: 'heading', revision: { formatChange: { id: '6', author: 'Cy' } } });
    expect(back[4]).toMatchObject({ type: 'listItem', revision: { insertion: { id: '7', author: 'Di' } } });
  });

  it('writes the old heading level of a paragraph that became something else', async () => {
    const blocks: DocBlock[] = [{ type: 'paragraph', content: [{ type: 'text', text: 'was a heading' }], revision: { formatChange: { id: '1', author: 'A', previous: { headingLevel: 3 } } } }];
    const { xml } = await exportXml(blocks);
    expect(xml).toMatch(/<w:pPrChange [^>]*><w:pPr><w:pStyle w:val="Heading3"\/>/);
  });

  it('writes inserted and deleted table rows and reads them back', async () => {
    const cells = [{ content: [{ type: 'text', text: 'c' } as DocInline] }];
    const blocks: DocBlock[] = [
      {
        type: 'table',
        rows: [{ cells, revision: { insertion: { id: '1', author: 'Ada', date: META.date } } }, { cells, revision: { deletion: { id: '2', author: 'Bo' } } }, { cells }],
      },
    ];
    const { xml } = await exportXml(blocks);
    expect(xml).toMatch(/<w:trPr><w:ins [^>]*w:author="Ada"[^>]*\/><\/w:trPr>/);
    expect(xml).toMatch(/<w:trPr><w:del [^>]*w:author="Bo"[^>]*\/><\/w:trPr>/);
    const rows = ((await roundTrip(blocks))[0] as Table).rows;
    expect(rows.map((row) => row.revision)).toEqual([
      { insertion: { id: '1', author: 'Ada', date: META.date } },
      { deletion: { id: '2', author: 'Bo' } },
      undefined,
    ]);
  });

  it('writes a changed link with both a change and a formatting change inside the hyperlink', async () => {
    const content: DocInline[] = [{ type: 'link', text: 'go', href: 'https://example.com/', revision: { insertion: { id: '1', author: 'Ada' }, formatChange: { ...META, previous: { bold: true } } } }];
    const back = ((await roundTrip([{ type: 'paragraph', content }]))[0] as Paragraph).content[0];
    expect(back).toMatchObject({ type: 'link', revision: { insertion: { id: '1' }, formatChange: { previous: { bold: true } } } });
  });
});
