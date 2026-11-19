import JSZip from 'jszip';
import { exportDocx } from './writer';
import { importDocx } from './reader';
import type { DocBlock, DocCommentThread, DocSection, DocumentExportInput } from '@b32nio/spindle-shared/export/docx';

const SECTION: Omit<DocSection, 'blocks'> = {
  pageWidthPx: 816,
  pageHeightPx: 1056,
  marginTopPx: 96,
  marginRightPx: 96,
  marginBottomPx: 96,
  marginLeftPx: 96,
  orientation: 'portrait',
};

const t = (text: string, commentThreadId?: string, extra: object = {}) => ({ type: 'text' as const, text, ...(commentThreadId && { commentThreadId }), ...extra });
const para = (...content: ReturnType<typeof t>[]): DocBlock => ({ type: 'paragraph', content });

function input(blocks: DocBlock[], comments?: DocCommentThread[], extra: Partial<DocSection> = {}): DocumentExportInput {
  return { title: 'T', sections: [{ ...SECTION, blocks, ...extra }], ...(comments && { comments }) };
}

async function parts(i: DocumentExportInput) {
  const out = await exportDocx(i);
  const zip = await JSZip.loadAsync(out.bytes);
  const text = async (name: string) => (await zip.file(name)?.async('string')) ?? '';
  return { out, zip, text };
}

const thread = (id: string, texts: string[], extra: Partial<DocCommentThread> = {}): DocCommentThread => ({
  id,
  quote: '',
  comments: texts.map((text, i) => ({ id: `${id}-c${i}`, author: i === 0 ? 'Alice' : 'Bob', initials: i === 0 ? 'A' : 'B', date: `2026-01-0${i + 1}T10:00:00Z`, text })),
  ...extra,
});

describe('exportDocx: comments', () => {
  it('writes no comment parts when there are no comments', async () => {
    const { text } = await parts(input([para(t('plain'))]));
    expect(await text('word/commentsExtended.xml')).toBe('');
    expect(await text('word/people.xml')).toBe('');
    expect(await text('word/document.xml')).not.toContain('comment');
  });

  it('writes the comment, extended and people parts with their content types and relationships', async () => {
    const { text } = await parts(input([para(t('see '), t('this', 'th1'), t(' now'))], [thread('th1', ['Root', 'Reply'])]));
    const types = await text('[Content_Types].xml');
    expect(types).toContain('PartName="/word/comments.xml"');
    expect(types).toContain('wordprocessingml.commentsExtended+xml" PartName="/word/commentsExtended.xml"');
    expect(types).toContain('wordprocessingml.people+xml" PartName="/word/people.xml"');
    const rels = await text('word/_rels/document.xml.rels');
    expect(rels).toMatch(/relationships\/comments" Target="comments.xml"/);
    expect(rels).toMatch(/2011\/relationships\/commentsExtended" Target="commentsExtended.xml"/);
    expect(rels).toMatch(/2011\/relationships\/people" Target="people.xml"/);
    const ids = [...rels.matchAll(/Id="(rId\d+)"/g)].map((m) => m[1]);
    expect(new Set(ids).size).toBe(ids.length);

    const comments = await text('word/comments.xml');
    expect(comments).toContain('<w:comment w:id="0" w:author="Alice" w:date="2026-01-01T10:00:00Z" w:initials="A">');
    expect(comments).toContain('<w:comment w:id="1" w:author="Bob" w:date="2026-01-02T10:00:00Z" w:initials="B">');
    expect(comments).toContain('<w:t xml:space="preserve">Root</w:t>');
    expect(await text('word/people.xml')).toContain('<w15:person w15:author="Alice"><w15:presenceInfo w15:providerId="None" w15:userId="Alice"/>');
  });

  it('wraps the run range in markers for every comment of the thread and adds the reference runs', async () => {
    const { text } = await parts(input([para(t('see '), t('this', 'th1'), t(' now'))], [thread('th1', ['Root', 'Reply'])]));
    const doc = await text('word/document.xml');
    expect(doc).toMatch(
      /see <\/w:t><\/w:r><w:commentRangeStart w:id="0"\/><w:commentRangeStart w:id="1"\/><w:r><w:t xml:space="preserve">this<\/w:t><\/w:r><w:commentRangeEnd w:id="0"\/><w:r><w:commentReference w:id="0"\/><\/w:r><w:commentRangeEnd w:id="1"\/><w:r><w:commentReference w:id="1"\/><\/w:r><w:r><w:t xml:space="preserve"> now/
    );
  });

  it('links replies to their root and writes the resolved state on every comment of a resolved thread', async () => {
    const { text } = await parts(
      input([para(t('a', 'th1'), t('b', 'th2'))], [thread('th1', ['Root', 'Reply'], { resolved: true }), thread('th2', ['Other'])])
    );
    const ext = [...(await text('word/commentsExtended.xml')).matchAll(/<w15:commentEx ([^>]*)\/>/g)].map((m) => m[1]);
    expect(ext).toEqual([
      'w15:paraId="10000001" w15:done="1"',
      'w15:paraId="10000002" w15:paraIdParent="10000001" w15:done="1"',
      'w15:paraId="10000003" w15:done="0"',
    ]);
  });

  it('writes multi-paragraph bodies as paragraphs with tabs, keeps the last paragraph id for threading, and escapes markup', async () => {
    const { text } = await parts(input([para(t('x', 'th1'))], [thread('th1', ['one\ttwo\nsecond <b> & "q"'])]));
    const comments = await text('word/comments.xml');
    expect(comments).toContain('<w:p w14:paraId="10000001"><w:r><w:annotationRef/></w:r><w:r><w:t xml:space="preserve">one</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t xml:space="preserve">two</w:t></w:r></w:p>');
    expect(comments).toContain('<w:p w14:paraId="10000002"><w:r><w:t xml:space="preserve">second &lt;b&gt; &amp; &quot;q&quot;</w:t></w:r></w:p>');
    expect(await text('word/commentsExtended.xml')).toContain('w15:paraId="10000002"');
  });

  it('removes control characters that XML cannot hold', async () => {
    const { text } = await parts(input([para(t('x', 'th1'))], [thread('th1', ['a\u0001b\u0008c'])]));
    expect(await text('word/comments.xml')).toContain('>abc<');
  });

  it('drops date fractions, omits an unparseable date and uses the author id in people.xml', async () => {
    const th = thread('th1', ['a', 'b']);
    th.comments[0].date = '2026-03-04T05:06:07.891Z';
    th.comments[1].date = 'yesterday';
    th.comments[0].authorId = 'S::alice@example.com::1';
    const { text } = await parts(input([para(t('x', 'th1'))], [th]));
    const comments = await text('word/comments.xml');
    expect(comments).toContain('w:date="2026-03-04T05:06:07Z"');
    expect(comments).toMatch(/<w:comment w:id="1" w:author="Bob" w:initials="B">/);
    expect(await text('word/people.xml')).toContain('w15:userId="S::alice@example.com::1"');
  });

  it('spans a thread over paragraphs and a table cell from its first run to its last', async () => {
    const table: DocBlock = { type: 'table', rows: [{ cells: [{ content: [t('cell', 'th1')] }, { content: [t('other')] }] }] };
    const { text } = await parts(input([para(t('first', 'th1')), para(t('between')), table], [thread('th1', ['c'])]));
    const doc = await text('word/document.xml');
    const start = doc.indexOf('<w:commentRangeStart w:id="0"/>');
    const end = doc.indexOf('<w:commentRangeEnd w:id="0"/>');
    expect(start).toBeGreaterThan(0);
    expect(end).toBeGreaterThan(doc.indexOf('between'));
    expect(end).toBeGreaterThan(doc.indexOf('cell'));
    expect(end).toBeLessThan(doc.indexOf('other'));
    expect(doc.match(/commentRangeStart/g)).toHaveLength(1);
    expect(doc.match(/commentReference/g)).toHaveLength(1);
  });

  it('writes overlapping ranges and a hyperlink inside a range', async () => {
    const link = { type: 'link' as const, text: 'site', href: 'https://example.com/', commentThreadId: 'th2' };
    const { text } = await parts(
      input([{ type: 'paragraph', content: [t('a', 'th1'), t('b', 'th2'), link, t('c', 'th1')] }], [thread('th1', ['one']), thread('th2', ['two'])])
    );
    const doc = await text('word/document.xml');
    const order = [...doc.matchAll(/<w:commentRange(Start|End) w:id="(\d)"\/>/g)].map((m) => `${m[1]}${m[2]}`);
    expect(order).toEqual(['Start0', 'Start1', 'End1', 'End0']);
  });

  it('anchors a thread without a tagged run at the start of the first body paragraph', async () => {
    const { text } = await parts(input([para(t('first')), para(t('second'))], [thread('th1', ['orphan', 'reply'])]));
    const doc = await text('word/document.xml');
    expect(doc).toMatch(
      /<w:p>(?:<w:pPr>.*?<\/w:pPr>)?<w:commentRangeStart w:id="0"\/><w:commentRangeStart w:id="1"\/><w:commentRangeEnd w:id="0"\/><w:r><w:commentReference w:id="0"\/><\/w:r><w:commentRangeEnd w:id="1"\/><w:r><w:commentReference w:id="1"\/><\/w:r><w:r><w:t xml:space="preserve">first/
    );
    expect(doc.match(/commentRangeStart/g)).toHaveLength(2);
  });

  it('warns and writes no comment parts when orphans have no paragraph to attach to', async () => {
    const { out, zip } = await parts(input([{ type: 'horizontalRule' }], [thread('th1', ['lost'])]));
    expect(out.warnings).toEqual(['1 comment(s) could not be anchored (the document has no paragraph) and were not written.']);
    expect(zip.file('word/commentsExtended.xml')).toBeNull();
  });

  it('ignores a run tagged with an unknown thread, empty threads and duplicate thread ids', async () => {
    const { text } = await parts(
      input([para(t('x', 'ghost'), t('y', 'th1'))], [thread('th1', ['kept']), thread('th1', ['duplicate']), { id: 'empty', quote: '', comments: [] }])
    );
    const doc = await text('word/document.xml');
    expect(doc.match(/commentRangeStart/g)).toHaveLength(1);
    const comments = await text('word/comments.xml');
    expect(comments).toContain('kept');
    expect(comments).not.toContain('duplicate');
  });

  it('does not anchor comments in headers and footers', async () => {
    const { text } = await parts(
      input([para(t('body'))], [thread('th1', ['c'])], { header: { paragraphs: [{ content: [t('head', 'th1')] }] } })
    );
    expect(await text('word/header1.xml')).not.toContain('commentRange');
    expect(await text('word/document.xml')).toContain('<w:commentRangeStart w:id="0"/>'); // falls back to the orphan anchor
  });

  describe('round trip through importDocx', () => {
    it('keeps threads, replies, resolved state, authors, initials, dates and anchors', async () => {
      const threads = [
        { ...thread('th1', ['Root', 'Reply one', 'Reply two'], { resolved: true, quote: 'this' }) },
        { ...thread('th2', ['Second']) },
      ];
      threads[0].comments[0].authorId = 'alice-id';
      const result = await importDocx((await exportDocx(input([para(t('see '), t('this', 'th1')), para(t('para two', 'th2'))], threads))).bytes);
      expect(result.warnings).toEqual([]);
      expect(result.comments).toHaveLength(2);
      const [a, b] = result.comments!;
      expect(a).toMatchObject({ resolved: true, quote: 'this' });
      expect(a.comments.map((c) => [c.author, c.initials, c.date, c.text])).toEqual([
        ['Alice', 'A', '2026-01-01T10:00:00Z', 'Root'],
        ['Bob', 'B', '2026-01-02T10:00:00Z', 'Reply one'],
        ['Bob', 'B', '2026-01-03T10:00:00Z', 'Reply two'],
      ]);
      expect(a.comments[0].authorId).toBe('alice-id');
      expect(b).toMatchObject({ quote: 'para two' });
      expect(b.resolved).toBeUndefined();
      const runs = result.sections[0].blocks.flatMap((blk) => (blk.type === 'paragraph' ? blk.content : []));
      expect(runs.map((r) => (r.type === 'text' ? [r.text, r.commentThreadId === a.id ? 'A' : r.commentThreadId === b.id ? 'B' : '-'] : []))).toEqual([
        ['see ', '-'],
        ['this', 'A'],
        ['para two', 'B'],
      ]);
    });

    it('is stable: exporting an imported file again gives the same threads', async () => {
      const first = await importDocx((await exportDocx(input([para(t('x', 'th1'))], [thread('th1', ['a\nb', 'c'], { resolved: true })]))).bytes);
      const again = await importDocx(
        (
          await exportDocx({
            title: 'T',
            sections: first.sections,
            comments: first.comments,
          })
        ).bytes
      );
      const strip = (r: typeof first) => r.comments!.map((th) => ({ ...th, id: '', comments: th.comments.map((c) => ({ ...c, id: '' })) }));
      expect(strip(again)).toEqual(strip(first));
    });

    it('keeps an orphan thread as an unanchored thread', async () => {
      const result = await importDocx((await exportDocx(input([para(t('first'))], [thread('th1', ['orphan'])]))).bytes);
      expect(result.comments).toHaveLength(1);
      expect(result.comments![0].quote).toBe('');
      expect(result.sections[0].blocks[0]).toMatchObject({ content: [{ type: 'text', text: 'first' }] });
    });
  });
});
