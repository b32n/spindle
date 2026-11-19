import JSZip from 'jszip';
import { importDocx } from './reader';
import { detectDocxLoss } from './loss';
import type { DocBlock, DocInline } from '@b32nio/spindle-shared/export/docx';

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" ' +
  'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"';

async function docx(body: string, parts: Record<string, string> = {}): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('word/document.xml', `<?xml version="1.0"?><w:document ${NS}><w:body>${body}<w:sectPr/></w:body></w:document>`);
  for (const [name, xml] of Object.entries(parts)) zip.file(name, xml);
  return zip.generateAsync({ type: 'uint8array' });
}

const comment = (id: number, author: string, text: string, extra = '', paraId = `0000000${id}`) =>
  `<w:comment w:id="${id}" w:author="${author}" w:date="2026-01-0${id + 1}T10:00:00Z" w:initials="${author[0]}${extra}">` +
  `<w:p w14:paraId="${paraId}"><w:r><w:annotationRef/></w:r><w:r><w:t>${text}</w:t></w:r></w:p></w:comment>`;
const comments = (...c: string[]) => `<w:comments ${NS}>${c.join('')}</w:comments>`;
const ext = (...e: string[]) => `<w15:commentsEx ${NS}>${e.join('')}</w15:commentsEx>`;
const ex = (paraId: string, done: 0 | 1, parent?: string) =>
  `<w15:commentEx w15:paraId="${paraId}"${parent ? ` w15:paraIdParent="${parent}"` : ''} w15:done="${done}"/>`;
const run = (t: string) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
const start = (id: number) => `<w:commentRangeStart w:id="${id}"/>`;
const end = (id: number) => `<w:commentRangeEnd w:id="${id}"/><w:r><w:commentReference w:id="${id}"/></w:r>`;
const p = (...inner: string[]) => `<w:p>${inner.join('')}</w:p>`;

const inlines = (blocks: DocBlock[]): DocInline[] => blocks.flatMap((b) => ('content' in b ? b.content : b.type === 'table' ? b.rows.flatMap((r) => r.cells.flatMap((c) => c.content)) : []));
const tagged = (blocks: DocBlock[]) =>
  inlines(blocks).flatMap((i) => (i.type === 'text' || i.type === 'link' ? [{ text: i.text, id: i.commentThreadId }] : []));

describe('importDocx: comments', () => {
  it('anchors a thread to the commented run and keeps author, initials and date', async () => {
    const result = await importDocx(
      await docx(p(run('before '), start(1), run('commented'), end(1), run(' after')), {
        'word/comments.xml': comments(comment(1, 'Alice', 'Check this')),
      })
    );
    expect(result.warnings).toEqual([]);
    expect(tagged(result.sections[0].blocks)).toEqual([
      { text: 'before ', id: undefined },
      { text: 'commented', id: 'docx-thread-1' },
      { text: ' after', id: undefined },
    ]);
    expect(result.comments).toEqual([
      {
        id: 'docx-thread-1',
        quote: 'commented',
        comments: [{ id: '1', author: 'Alice', authorId: 'Alice', initials: 'A', date: '2026-01-02T10:00:00Z', text: 'Check this' }],
      },
    ]);
  });

  it('has no comments key when the file has none', async () => {
    const result = await importDocx(await docx(p(run('plain'))));
    expect(result.comments).toBeUndefined();
  });

  it('groups replies under the root through commentsExtended and takes the resolved state from the root', async () => {
    const result = await importDocx(
      await docx(p(start(1), start(2), start(3), run('text'), end(1), end(2), end(3)), {
        'word/comments.xml': comments(comment(1, 'Alice', 'Root'), comment(2, 'Bob', 'Reply one'), comment(3, 'Carol', 'Reply two')),
        'word/commentsExtended.xml': ext(ex('00000001', 1), ex('00000002', 0, '00000001'), ex('00000003', 1, '00000001')),
      })
    );
    expect(result.comments).toHaveLength(1);
    const thread = result.comments![0];
    expect(thread.id).toBe('docx-thread-1');
    expect(thread.resolved).toBe(true);
    expect(thread.comments.map((c) => `${c.author}: ${c.text}`)).toEqual(['Alice: Root', 'Bob: Reply one', 'Carol: Reply two']);
    expect(tagged(result.sections[0].blocks)).toEqual([{ text: 'text', id: 'docx-thread-1' }]);
  });

  it('does not mark a thread resolved when only a reply is done', async () => {
    const result = await importDocx(
      await docx(p(start(1), run('x'), end(1)), {
        'word/comments.xml': comments(comment(1, 'A', 'r'), comment(2, 'B', 'reply')),
        'word/commentsExtended.xml': ext(ex('00000001', 0), ex('00000002', 1, '00000001')),
      })
    );
    expect(result.comments![0].resolved).toBeUndefined();
  });

  it('keeps every comment its own thread without commentsExtended', async () => {
    const result = await importDocx(
      await docx(p(start(1), run('a'), end(1), start(2), run('b'), end(2)), {
        'word/comments.xml': comments(comment(1, 'A', 'one'), comment(2, 'B', 'two')),
      })
    );
    expect(result.comments!.map((t) => t.id)).toEqual(['docx-thread-1', 'docx-thread-2']);
  });

  it('survives parent links that loop or point nowhere', async () => {
    const result = await importDocx(
      await docx(p(start(1), run('a'), end(1)), {
        'word/comments.xml': comments(comment(1, 'A', 'one'), comment(2, 'B', 'two')),
        'word/commentsExtended.xml': ext(ex('00000001', 0, '00000002'), ex('00000002', 0, '00000001')),
      })
    );
    expect(result.comments).toHaveLength(1);
    expect(result.comments![0].comments).toHaveLength(2);
    const dangling = await importDocx(
      await docx(p(run('a')), {
        'word/comments.xml': comments(comment(1, 'A', 'one')),
        'word/commentsExtended.xml': ext(ex('00000001', 0, '99999999')),
      })
    );
    expect(dangling.comments).toHaveLength(1);
  });

  it('spans paragraphs and builds the quote with a space between them', async () => {
    const result = await importDocx(
      await docx(p(start(1), run('first')) + p(run('second')) + p(run('third'), end(1), run(' tail')), {
        'word/comments.xml': comments(comment(1, 'A', 'across')),
      })
    );
    const blocks = result.sections[0].blocks;
    expect(tagged(blocks)).toEqual([
      { text: 'first', id: 'docx-thread-1' },
      { text: 'second', id: 'docx-thread-1' },
      { text: 'third', id: 'docx-thread-1' },
      { text: ' tail', id: undefined },
    ]);
    expect(result.comments![0].quote).toBe('first second third');
  });

  it('anchors across a link and inside a tracked insertion', async () => {
    const rels = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9" Type="x" Target="https://example.com/" TargetMode="External"/></Relationships>';
    const result = await importDocx(
      await docx(
        p(start(1), '<w:hyperlink r:id="rId9" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' + run('link') + '</w:hyperlink>', '<w:ins w:id="5">' + run('inserted') + '</w:ins>', end(1)),
        { 'word/comments.xml': comments(comment(1, 'A', 'c')), 'word/_rels/document.xml.rels': rels }
      )
    );
    const items = inlines(result.sections[0].blocks);
    expect(items.map((i) => (i.type === 'link' || i.type === 'text' ? [i.type, i.text, i.commentThreadId] : []))).toEqual([
      ['link', 'link', 'docx-thread-1'],
      ['text', 'inserted', 'docx-thread-1'],
    ]);
  });

  it('follows ranges that open and close between paragraphs, table rows and cells', async () => {
    const cell = (inner: string) => `<w:tc>${inner}</w:tc>`;
    const body =
      start(1) +
      '<w:tbl><w:tr>' + cell(p(run('A1'))) + end(1) + cell(p(run('B1'))) + '</w:tr></w:tbl>' +
      p(run('outside'));
    const result = await importDocx(await docx(body, { 'word/comments.xml': comments(comment(1, 'A', 'table')) }));
    expect(tagged(result.sections[0].blocks)).toEqual([
      { text: 'A1', id: 'docx-thread-1' },
      { text: 'B1', id: undefined },
      { text: 'outside', id: undefined },
    ]);
  });

  it('closes a range whose end sits in a cell paragraph that is not imported', async () => {
    const cell = `<w:tc>${p(run('first'))}${p(run('second'), end(1))}</w:tc>`;
    const result = await importDocx(
      await docx(p(start(1), run('before')) + `<w:tbl><w:tr>${cell}</w:tr></w:tbl>` + p(run('after')), {
        'word/comments.xml': comments(comment(1, 'A', 'c')),
      })
    );
    expect(tagged(result.sections[0].blocks)).toEqual([
      { text: 'before', id: 'docx-thread-1' },
      { text: 'first', id: 'docx-thread-1' },
      { text: 'after', id: undefined },
    ]);
  });

  it('tags the innermost of overlapping ranges but quotes the text for both threads', async () => {
    const result = await importDocx(
      await docx(p(start(1), run('outer '), start(2), run('inner'), end(2), run(' end'), end(1)), {
        'word/comments.xml': comments(comment(1, 'A', 'outer'), comment(2, 'B', 'inner')),
      })
    );
    expect(tagged(result.sections[0].blocks).map((t) => t.id)).toEqual(['docx-thread-1', 'docx-thread-2', 'docx-thread-1']);
    expect(Object.fromEntries(result.comments!.map((t) => [t.id, t.quote]))).toEqual({
      'docx-thread-1': 'outer inner end',
      'docx-thread-2': 'inner',
    });
  });

  it('keeps a comment with a collapsed range, no range, or an unterminated range as a thread without anchor', async () => {
    const result = await importDocx(
      await docx(p(start(1), end(1), run('a')) + p(run('b')) + p(start(3), run('never closed')) + p(run('after')), {
        'word/comments.xml': comments(comment(1, 'A', 'collapsed'), comment(2, 'B', 'no range'), comment(3, 'C', 'open')),
      })
    );
    expect(result.comments!.map((t) => [t.id, t.quote])).toEqual([
      ['docx-thread-1', ''],
      ['docx-thread-2', ''],
      ['docx-thread-3', ''],
    ]);
    expect(tagged(result.sections[0].blocks).every((t) => t.id === undefined)).toBe(true);
  });

  it('ignores range markers whose comment does not exist', async () => {
    const result = await importDocx(await docx(p(start(7), run('x'), end(7)), { 'word/comments.xml': comments(comment(1, 'A', 'c')) }));
    expect(tagged(result.sections[0].blocks)).toEqual([{ text: 'x', id: undefined }]);
  });

  it('truncates a long quote', async () => {
    const long = 'word '.repeat(100);
    const result = await importDocx(await docx(p(start(1), run(long), end(1)), { 'word/comments.xml': comments(comment(1, 'A', 'c')) }));
    const quote = result.comments![0].quote;
    expect(quote.length).toBe(201);
    expect(quote.endsWith('…')).toBe(true);
  });

  it('joins multi-paragraph bodies with newlines and keeps tabs, breaks and skips deleted text', async () => {
    const body =
      '<w:comment w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z">' +
      '<w:p><w:r><w:t>one</w:t><w:tab/><w:t>two</w:t><w:br/><w:t>three</w:t></w:r><w:del w:id="9"><w:r><w:delText>gone</w:delText></w:r></w:del></w:p>' +
      '<w:p><w:r><w:t>second</w:t></w:r></w:p></w:comment>';
    const result = await importDocx(await docx(p(start(1), run('x'), end(1)), { 'word/comments.xml': comments(body) }));
    expect(result.comments![0].comments[0].text).toBe('one\ttwo\nthree\nsecond');
  });

  it('defaults a missing author, id and date', async () => {
    const result = await importDocx(
      await docx(p(start(1), run('x'), end(1)), {
        'word/comments.xml': comments('<w:comment w:id="1"><w:p><w:r><w:t>hi</w:t></w:r></w:p></w:comment><w:comment><w:p/></w:comment>'),
      })
    );
    expect(result.comments![0].comments[0]).toEqual({ id: '1', author: 'Unknown', authorId: 'Unknown', text: 'hi' });
    expect(result.comments).toHaveLength(1);
  });

  it('uses the user id from people.xml as the author id', async () => {
    const people = `<w15:people ${NS}><w15:person w15:author="Alice"><w15:presenceInfo w15:providerId="AD" w15:userId="S::alice@example.com::1"/></w15:person></w15:people>`;
    const result = await importDocx(
      await docx(p(start(1), run('x'), end(1)), { 'word/comments.xml': comments(comment(1, 'Alice', 'c')), 'word/people.xml': people })
    );
    expect(result.comments![0].comments[0]).toMatchObject({ author: 'Alice', authorId: 'S::alice@example.com::1' });
  });

  it('keeps comments with a warning-free import when commentsExtended or people are garbage', async () => {
    const result = await importDocx(
      await docx(p(start(1), run('x'), end(1)), {
        'word/comments.xml': comments(comment(1, 'A', 'c')),
        'word/commentsExtended.xml': 'not xml at all \u0000\u0001',
        'word/people.xml': '<<<>',
      })
    );
    expect(result.warnings).toEqual([]);
    expect(result.comments).toHaveLength(1);
  });

  it('warns and imports the text when the comments part is not a comments part', async () => {
    const result = await importDocx(await docx(p(start(1), run('x'), end(1)), { 'word/comments.xml': '<nope/>' }));
    expect(result.comments).toBeUndefined();
    expect(result.warnings).toEqual(['The comments part could not be read, so its comments were not imported.']);
    expect(tagged(result.sections[0].blocks)).toEqual([{ text: 'x', id: undefined }]);
  });

  it('keeps what a part truncated inside a comment still holds', async () => {
    const whole = comments(comment(1, 'A', 'one'), comment(2, 'B', 'two'));
    const truncated = whole.slice(0, whole.indexOf('two') + 2);
    const result = await importDocx(await docx(p(start(1), run('x'), end(1)), { 'word/comments.xml': truncated }));
    expect(result.warnings).toEqual([]);
    expect(result.comments?.map((t) => t.comments[0].text)).toEqual(['one', '']); // the unterminated text itself is not recoverable
  });

  it('warns, and keeps the text, when the comments part is cut inside a tag', async () => {
    const whole = comments(comment(1, 'A', 'one'));
    const result = await importDocx(await docx(p(start(1), run('x'), end(1)), { 'word/comments.xml': whole.slice(0, whole.indexOf('w:author') + 4) }));
    expect(result.comments).toBeUndefined();
    expect(result.warnings).toHaveLength(1);
    expect(tagged(result.sections[0].blocks)).toEqual([{ text: 'x', id: undefined }]);
  });

  it('does not report comments as lost by the loss detector once they are imported', async () => {
    const bytes = await docx(p(start(1), run('x'), end(1)), { 'word/comments.xml': comments(comment(1, 'A', 'c')) });
    expect((await detectDocxLoss(bytes)).items.find((i) => i.id === 'docx.comments')).toBeUndefined();
  });

  it('reports the comments of an unreadable part as lost', async () => {
    const bytes = await docx(p(run('x')), { 'word/comments.xml': '<nope><w:comment w:id="1"/><w:comment w:id="2"/></nope>' });
    expect((await detectDocxLoss(bytes)).items.find((i) => i.id === 'docx.comments')).toMatchObject({ count: 2 });
  });
});
