import JSZip from 'jszip';
import type { PptxCommentThread, PptxSlide } from '@b32nio/spindle-shared/export/pptx';
import { importPptx } from './reader';
import { OoxmlPackage } from '../common/ooxml-package';
import { isoTime } from './comments';

// Speaker notes and comments (plan K10c), read from synthetic decks that are real zips.

const NS =
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"';
const P188 = 'xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main"';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const MS_REL = 'http://schemas.microsoft.com/office/2018/10/relationships';
const PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const PRES_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml';

const rel = (id: string, type: string, target: string, external = false, ns = REL): string => `<Relationship Id="${id}" Type="${ns}/${type}" Target="${target}"${external ? ' TargetMode="External"' : ''}/>`;
const rels = (list: string[]): string => `<Relationships xmlns="${PKG_REL}">${list.join('')}</Relationships>`;

interface Deck {
  /** The shape tree of slide 1. */
  tree?: string;
  slideRels?: string[];
  presRels?: string[];
  /** Parts besides the standard ones. */
  extra?: Record<string, string>;
}

async function build(d: Deck): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', `<Types xmlns="${CT}"><Override PartName="/ppt/presentation.xml" ContentType="${PRES_TYPE}"/></Types>`);
  zip.file('_rels/.rels', rels([rel('rId1', 'officeDocument', 'ppt/presentation.xml')]));
  zip.file('ppt/_rels/presentation.xml.rels', rels([rel('rId1', 'slide', 'slides/slide1.xml'), ...(d.presRels ?? [])]));
  zip.file('ppt/presentation.xml', `<p:presentation ${NS}><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>`);
  zip.file('ppt/slides/slide1.xml', `<p:sld ${NS}><p:cSld><p:spTree>${d.tree ?? ''}</p:spTree></p:cSld></p:sld>`);
  zip.file('ppt/slides/_rels/slide1.xml.rels', rels(d.slideRels ?? []));
  for (const [name, content] of Object.entries(d.extra ?? {})) zip.file(name, content);
  return zip.generateAsync({ type: 'uint8array' });
}

const slideOf = async (d: Deck): Promise<{ slide: PptxSlide; warnings: string[] }> => {
  const result = await importPptx(await build(d));
  return { slide: result.slides[0], warnings: result.warnings };
};

// ---- notes ---------------------------------------------------------------------------------------------------------

const ph = (type: string, id = 2): string => `<p:nvSpPr><p:cNvPr id="${id}" name="n"/><p:cNvSpPr/><p:nvPr><p:ph type="${type}"/></p:nvPr></p:nvSpPr>`;
const body = (inner: string, type = 'body', id = 2): string => `<p:sp>${ph(type, id)}<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>${inner}</p:txBody></p:sp>`;
const para = (text: string, rPr = ''): string => `<a:p><a:r>${rPr}<a:t>${text}</a:t></a:r></a:p>`;
const notesPage = (tree: string, extra = ''): string => `<p:notes ${NS}><p:cSld><p:spTree>${tree}</p:spTree></p:cSld>${extra}</p:notes>`;
const NOTES_MASTER = `<p:notesMaster ${NS}><p:cSld><p:spTree><p:sp>${ph('body')}<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle><a:lvl2pPr marL="100"><a:defRPr sz="1400"/></a:lvl2pPr></a:lstStyle><a:p><a:r><a:t>Click to edit Master text styles</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld><p:clrMap bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:notesStyle><a:lvl1pPr marL="0"><a:defRPr sz="1200" b="1"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:defRPr></a:lvl1pPr></p:notesStyle></p:notesMaster>`;

const withNotes = (page: string, notesRels: string[] = [rel('rId1', 'notesMaster', '../notesMasters/notesMaster1.xml')], master: string | undefined = NOTES_MASTER): Deck => ({
  slideRels: [rel('rIdN', 'notesSlide', '../notesSlides/notesSlide1.xml')],
  extra: {
    'ppt/notesSlides/notesSlide1.xml': page,
    'ppt/notesSlides/_rels/notesSlide1.xml.rels': rels(notesRels),
    ...(master !== undefined ? { 'ppt/notesMasters/notesMaster1.xml': master } : {}),
  },
});

describe('speaker notes', () => {
  it('reads the body placeholder of the notes page, with its text and run formatting', async () => {
    const { slide } = await slideOf(withNotes(notesPage(body(para('Say hello.', '<a:rPr i="1"/>') + para('Pause.')) + body(para('3'), 'sldNum', 3))));
    expect(slide.notes?.paragraphs.map((p) => p.runs.map((r) => r.text))).toEqual([['Say hello.'], ['Pause.']]);
    expect(slide.notes?.paragraphs[0].runs[0].italic).toBe(true);
  });

  it('takes the text styles of the notes master (its notesStyle, its body placeholder list style and its colour map) as the base', async () => {
    const { slide } = await slideOf(withNotes(notesPage(body(para('A') + `<a:p><a:pPr lvl="1"/><a:r><a:t>B</a:t></a:r></a:p>`))));
    const [first, second] = slide.notes?.paragraphs ?? [];
    expect(first.style).toMatchObject({ fontSize: 12, bold: true });
    // the master's colour map sends tx1 to lt1
    expect(first.style.color).toEqual({ kind: 'theme', slot: 'lt1' });
    // level 2 has the master body placeholder's own list style (marL 100 EMU, 14 pt)
    expect(second.style).toMatchObject({ fontSize: 14 });
  });

  it('applies the notes page colour map override over the master\'s', async () => {
    const page = notesPage(body(para('A')), '<p:clrMapOvr><a:overrideClrMapping bg1="lt1" tx1="dk1" bg2="lt2" tx2="lt2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:clrMapOvr>');
    const { slide } = await slideOf(withNotes(page));
    expect(slide.notes?.paragraphs[0].style.color).toEqual({ kind: 'theme', slot: 'dk1' });
  });

  it('keeps hyperlinks of notes text against the notes page relationships', async () => {
    const link = '<a:rPr><a:hlinkClick r:id="rL"/></a:rPr>';
    const { slide } = await slideOf(withNotes(notesPage(body(para('site', link))), [rel('rId1', 'notesMaster', '../notesMasters/notesMaster1.xml'), rel('rL', 'hyperlink', 'https://example.com/', true)]));
    expect(slide.notes?.paragraphs[0].runs[0].link).toEqual({ kind: 'url', href: 'https://example.com/' });
  });

  it('joins several body placeholders and ignores a body without text body', async () => {
    const tree = body(para('One')) + `<p:sp>${ph('body', 4)}<p:spPr/></p:sp>` + body(para('Two'), 'body', 5);
    const { slide } = await slideOf(withNotes(notesPage(tree)));
    expect(slide.notes?.paragraphs.map((p) => p.runs[0].text)).toEqual(['One', 'Two']);
  });

  it('finds the body placeholder inside alternate content', async () => {
    const tree = `<mc:AlternateContent><mc:Choice Requires="x">${body(para('In choice'))}</mc:Choice></mc:AlternateContent>`;
    const { slide } = await slideOf(withNotes(notesPage(tree)));
    expect(slide.notes?.paragraphs[0].runs[0].text).toBe('In choice');
  });

  it('reads notes without a notes master (default colours, no inherited style)', async () => {
    const { slide } = await slideOf(withNotes(notesPage(body(para('Plain'))), [], undefined));
    expect(slide.notes?.paragraphs[0].style).toEqual({});
    const dangling = await slideOf(withNotes(notesPage(body(para('Plain'))), [rel('rId1', 'notesMaster', '../notesMasters/gone.xml')], undefined));
    expect(dangling.slide.notes?.paragraphs[0].runs[0].text).toBe('Plain');
  });

  it('reads a notes master that has no body placeholder', async () => {
    const master = `<p:notesMaster ${NS}><p:cSld><p:spTree/></p:cSld></p:notesMaster>`;
    const { slide } = await slideOf(withNotes(notesPage(body(para('A'))), undefined, master));
    expect(slide.notes?.paragraphs[0].runs[0].text).toBe('A');
  });

  it('gives a field run as text and counts a break-only or blank page as no notes', async () => {
    const field = `<a:p><a:fld type="slidenum"><a:t>7</a:t></a:fld></a:p>`;
    expect((await slideOf(withNotes(notesPage(body(field))))).slide.notes?.paragraphs[0].runs[0]).toMatchObject({ kind: 'field', text: '7' });
    for (const inner of [`<a:p><a:br/></a:p>`, para('   '), '<a:p/>'] as const) expect((await slideOf(withNotes(notesPage(body(inner))))).slide.notes).toBeUndefined();
  });

  it('has no notes for a page without a body placeholder, an unreadable page, a page of the wrong kind or no notes page', async () => {
    expect((await slideOf(withNotes(notesPage(body(para('only a header'), 'hdr'))))).slide.notes).toBeUndefined();
    expect((await slideOf(withNotes('<p:notes'))).slide.notes).toBeUndefined();
    expect((await slideOf(withNotes(`<p:sld ${NS}/>`))).slide.notes).toBeUndefined();
    expect((await slideOf({ slideRels: [rel('rIdN', 'notesSlide', '../notesSlides/gone.xml')] })).slide.notes).toBeUndefined();
    expect((await slideOf({})).slide.notes).toBeUndefined();
    expect((await slideOf({ slideRels: [rel('rIdN', 'notesSlide', 'http://x/n.xml', true)] })).slide.notes).toBeUndefined();
  });
});

// ---- comments ------------------------------------------------------------------------------------------------------

const LEGACY_AUTHORS = `<p:cmAuthorLst ${NS}><p:cmAuthor id="0" name="Ann" initials="A" lastIdx="2" clrIdx="0"/><p:cmAuthor id="1" name="  " lastIdx="1" clrIdx="1"/></p:cmAuthorLst>`;
const lcm = (author: string, idx: string, text: string, extra = '', dt = '2026-01-15T09:00:00.000'): string => `<p:cm authorId="${author}" dt="${dt}" idx="${idx}"><p:pos x="600" y="1200"/><p:text>${text}</p:text>${extra}</p:cm>`;
const parent = (author: string, idx: string): string => `<p:extLst><p:ext uri="{C676402C-5697-4E1C-873F-D02D1690AC5C}"><p15:threadingInfo xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main" timeZoneBias="-60"><p15:parentCm authorId="${author}" idx="${idx}"/></p15:threadingInfo></p:ext></p:extLst>`;
const legacyDeck = (cms: string, authors: string | null = LEGACY_AUTHORS): Deck => ({
  presRels: authors !== null ? [rel('rIdA', 'commentAuthors', 'commentAuthors.xml')] : [],
  slideRels: [rel('rIdC', 'comments', '../comments/comment1.xml')],
  extra: { 'ppt/comments/comment1.xml': `<p:cmLst ${NS}>${cms}</p:cmLst>`, ...(authors !== null ? { 'ppt/commentAuthors.xml': authors } : {}) },
});
const threadsOf = async (d: Deck): Promise<PptxCommentThread[]> => (await slideOf(d)).slide.threads ?? [];

describe('legacy comments', () => {
  it('reads a comment with its author, time and position (1/576 inch to px)', async () => {
    expect(await threadsOf(legacyDeck(lcm('0', '1', ' Check this. ')))).toEqual([
      { id: '256-0.1', format: 'legacy', anchor: { kind: 'slide' }, position: { x: 100, y: 200 }, comments: [{ author: 'Ann', initials: 'A', body: 'Check this.', createdAt: '2026-01-15T09:00:00.000Z' }] },
    ]);
  });

  it('puts a reply (a cm with a parentCm) into the thread of the comment it answers', async () => {
    const threads = await threadsOf(legacyDeck(lcm('0', '1', 'Q') + lcm('1', '1', 'A', parent('0', '1')) + lcm('0', '2', 'Next')));
    expect(threads.map((t) => t.comments.map((c) => c.body))).toEqual([['Q', 'A'], ['Next']]);
  });

  it('puts a reply that comes before its parent into the parent\'s thread, and a reply to a reply into the thread at the top', async () => {
    const early = await threadsOf(legacyDeck(lcm('1', '1', 'A', parent('0', '1')) + lcm('0', '1', 'Q')));
    expect(early.map((t) => t.comments.map((c) => c.body))).toEqual([['Q', 'A']]);
    const chain = await threadsOf(legacyDeck(lcm('0', '1', 'Q') + lcm('1', '1', 'A', parent('0', '1')) + lcm('0', '2', 'B', parent('1', '1')) + lcm('1', '2', 'C', parent('0', '2'))));
    expect(chain.map((t) => t.comments.map((c) => c.body))).toEqual([['Q', 'A', 'B', 'C']]);
  });

  it('keeps comments whose parents loop, or point at themselves, as threads of their own', async () => {
    const loop = await threadsOf(legacyDeck(lcm('0', '1', 'a', parent('0', '2')) + lcm('0', '2', 'b', parent('0', '1')) + lcm('0', '3', 'c', parent('0', '3')) + lcm('0', '4', 'd', parent('0', '1'))));
    expect(loop.map((t) => t.comments.map((c) => c.body))).toEqual([['a'], ['b'], ['c'], ['d']]);
  });

  it('keeps a reply to a thread whose own parent is missing in that thread', async () => {
    const t = await threadsOf(legacyDeck(lcm('0', '5', 'x', parent('0', '1')) + lcm('0', '6', 'y', parent('0', '5'))));
    expect(t.map((x) => x.comments.map((c) => c.body))).toEqual([['x', 'y']]);
  });

  it('keeps a reply whose parent is missing as a thread of its own', async () => {
    expect((await threadsOf(legacyDeck(lcm('0', '5', 'orphan', parent('0', '1'))))).map((t) => t.comments.length)).toEqual([1]);
  });

  it('names an unlisted or nameless author "Unknown author" and tolerates a missing position, text and time', async () => {
    const threads = await threadsOf(legacyDeck('<p:cm authorId="9" idx="1"/>' + lcm('1', '1', 'x', '', 'yesterday'), LEGACY_AUTHORS));
    expect(threads[0]).toEqual({ id: '256-9.1', format: 'legacy', anchor: { kind: 'slide' }, comments: [{ author: 'Unknown author', body: '' }] });
    expect(threads[1].comments[0]).toEqual({ author: 'Unknown author', body: 'x' });
  });

  it('keeps thread ids unique when a part repeats an author and index', async () => {
    const threads = await threadsOf(legacyDeck(lcm('0', '1', 'a') + lcm('0', '1', 'b') + lcm('0', '1', 'c')));
    expect(threads.map((t) => t.id)).toEqual(['256-0.1', '256-0.1-2', '256-0.1-3']);
  });

  it('tolerates a comment and a reply without author, index or parent attributes', async () => {
    const bare = '<p:cm><p:text>a</p:text></p:cm><p:cm><p:text>b</p:text><p:extLst><p:ext><p15:threadingInfo xmlns:p15="urn:p"><p15:parentCm/></p15:threadingInfo></p:ext></p:extLst></p:cm>';
    expect((await threadsOf(legacyDeck(bare))).map((t) => t.comments.map((c) => c.body))).toEqual([['a', 'b']]);
  });

  it('works without an author list', async () => {
    expect((await threadsOf(legacyDeck(lcm('0', '1', 'x'), null)))[0].comments[0].author).toBe('Unknown author');
  });

  it('has no threads for a part that is missing, is not a comment list or is linked externally', async () => {
    expect(await threadsOf({ slideRels: [rel('rIdC', 'comments', '../comments/none.xml')] })).toEqual([]);
    expect(await threadsOf({ slideRels: [rel('rIdC', 'comments', '../comments/c.xml')], extra: { 'ppt/comments/c.xml': `<p:sld ${NS}/>` } })).toEqual([]);
    expect(await threadsOf({ slideRels: [rel('rIdC', 'comments', 'http://x/c.xml', true)] })).toEqual([]);
    expect((await slideOf(legacyDeck(''))).slide.threads).toBeUndefined();
  });
});

const AUTHOR_A = '{00000000-0000-4000-8000-0000000000A1}';
const AUTHOR_B = '{00000000-0000-4000-8000-0000000000A2}';
const MODERN_AUTHORS = `<p188:authorLst ${NS} ${P188}><p188:author id="${AUTHOR_A}" name="Bea" initials="B" userId="b" providerId="None"/><p188:author id="${AUTHOR_B}" name="Cy" userId="c" providerId="None"/><p188:author name="No id"/><p188:author id="{ANON}"/><p188:author id="${AUTHOR_A}" name="Dup"/></p188:authorLst>`;
const tx = (...paras: string[]): string => `<p188:txBody><a:bodyPr/><a:lstStyle/>${paras.map((t) => `<a:p><a:r><a:t>${t}</a:t></a:r></a:p>`).join('')}</p188:txBody>`;
const SLIDE_MK = '<pc:sldMkLst xmlns:pc="http://schemas.microsoft.com/office/powerpoint/2013/main/command"><pc:docMk/><pc:sldMk cId="1" sldId="256"/></pc:sldMkLst>';
const shapeMk = (mk: string, id: string): string => `<ac:deMkLst xmlns:ac="http://schemas.microsoft.com/office/drawing/2013/main/command" xmlns:pc="http://schemas.microsoft.com/office/powerpoint/2013/main/command"><pc:docMk/><pc:sldMk cId="1" sldId="256"/><ac:${mk} id="${id}" creationId="{X}"/></ac:deMkLst>`;
const mcm = (id: string, author: string, inner: string, attrs = ''): string => `<p188:cm id="${id}" authorId="${author}" created="2026-02-03T04:05:06.000"${attrs}>${inner}</p188:cm>`;
const modernDeck = (cms: string, tree = ''): Deck => ({
  tree,
  presRels: [rel('rIdA', 'authors', 'authors.xml', false, MS_REL)],
  slideRels: [rel('rIdC', 'comments', '../comments/modernComment_1.xml', false, MS_REL)],
  extra: { 'ppt/comments/modernComment_1.xml': `<p188:cmLst ${NS} ${P188}>${cms}</p188:cmLst>`, 'ppt/authors.xml': MODERN_AUTHORS },
});
const shape = (id: number, inner = ''): string => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="S${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp>${inner}`;

describe('modern comments', () => {
  it('reads a comment with its replies, authors and slide anchor', async () => {
    const reply = `<p188:replyLst><p188:reply id="{R1}" authorId="${AUTHOR_B}" created="2026-02-03T04:05:07Z">${tx('Yes', 'really')}</p188:reply></p188:replyLst>`;
    expect(await threadsOf(modernDeck(mcm('{C1}', AUTHOR_A, SLIDE_MK + reply + tx('Final?'))))).toEqual([
      {
        id: '256-{C1}',
        format: 'modern',
        anchor: { kind: 'slide' },
        comments: [
          { author: 'Bea', initials: 'B', body: 'Final?', createdAt: '2026-02-03T04:05:06.000Z' },
          { author: 'Cy', body: 'Yes\nreally', createdAt: '2026-02-03T04:05:07.000Z' },
        ],
      },
    ]);
  });

  it('anchors to a shape the slide holds, at any depth of its groups, and to the slide for an id it does not hold', async () => {
    const tree = shape(2) + `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="9" name="G"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${shape(10)}</p:grpSp>`;
    const threads = await threadsOf(
      modernDeck(mcm('{A}', AUTHOR_A, shapeMk('spMk', '2') + tx('on shape')) + mcm('{B}', AUTHOR_A, shapeMk('picMk', '10') + tx('in group')) + mcm('{C}', AUTHOR_A, shapeMk('spMk', '99') + tx('gone')) + mcm('{D}', AUTHOR_A, '<ac:deMkLst xmlns:ac="urn:x"><ac:txMk/></ac:deMkLst>' + tx('no id')), tree)
    );
    expect(threads.map((t) => t.anchor)).toEqual([{ kind: 'element', pptxId: '2' }, { kind: 'element', pptxId: '10' }, { kind: 'slide' }, { kind: 'slide' }]);
  });

  it('marks resolved and closed comments resolved and leaves active ones open', async () => {
    const threads = await threadsOf(
      modernDeck(mcm('{A}', AUTHOR_A, tx('a'), ' status="resolved"') + mcm('{B}', AUTHOR_A, tx('b'), ' status="closed"') + mcm('{C}', AUTHOR_A, tx('c'), ' status="active"') + mcm('{D}', AUTHOR_A, tx('d')))
    );
    expect(threads.map((t) => t.resolved)).toEqual([true, true, undefined, undefined]);
  });

  it('tolerates a reply without an author and authors without a name, initials or a duplicate id', async () => {
    const reply = `<p188:replyLst><p188:reply id="{R1}">${tx('r')}</p188:reply></p188:replyLst>`;
    const [t] = await threadsOf(modernDeck(mcm('{C1}', '{ANON}', reply + tx('q'))));
    expect(t.comments.map((c) => c.author)).toEqual(['Unknown author', 'Unknown author']);
    expect((await threadsOf(modernDeck(`<p188:cm>${tx('anon')}</p188:cm>`)))[0].comments[0].author).toBe('Unknown author');
    // the first author with an id wins
    expect((await threadsOf(modernDeck(mcm('{C1}', AUTHOR_A, tx('q')))))[0].comments[0].author).toBe('Bea');
  });

  it('gives a comment with no id one of its own, keeps ids unique and names an unlisted author', async () => {
    const noId = `<p188:cm authorId="${AUTHOR_B}">${tx('x')}</p188:cm>`;
    const threads = await threadsOf(modernDeck(noId + noId + mcm('{D}', '{nobody}', tx('y')) + mcm('{D}', AUTHOR_A, tx('z'))));
    expect(threads.map((t) => t.id)).toEqual(['256-m', '256-m-2', '256-{D}', '256-{D}-2']);
    expect(threads[2].comments[0].author).toBe('Unknown author');
    expect(threads[0].comments[0].createdAt).toBeUndefined();
  });

  it('reads a legacy and a modern part on the same slide, modern first', async () => {
    const d = modernDeck(mcm('{A}', AUTHOR_A, tx('m')));
    d.slideRels = [...(d.slideRels ?? []), rel('rIdL', 'comments', '../comments/comment1.xml')];
    d.extra = { ...d.extra, 'ppt/comments/comment1.xml': `<p:cmLst ${NS}>${lcm('0', '1', 'l')}</p:cmLst>` };
    expect((await threadsOf(d)).map((t) => t.format)).toEqual(['modern', 'legacy']);
  });

  it('stops at 10,000 threads of a slide with a warning', async () => {
    const many = Array.from({ length: 10_001 }, (_, i) => mcm(`{${i}}`, AUTHOR_A, tx('x'))).join('');
    const { slide, warnings } = await slideOf(modernDeck(many));
    expect(slide.threads).toHaveLength(10_000);
    expect(warnings.some((w) => w.includes('10000'))).toBe(true);
  });

  it('caps replies too: one thread with a very large number of replies is cut at 10,000 comments, later threads dropped', async () => {
    const replies = Array.from({ length: 10_005 }, (_, i) => `<p188:reply id="{r${i}}" authorId="${AUTHOR_A}">${tx('r')}</p188:reply>`).join('');
    const big = `<p188:cm id="{big}" authorId="${AUTHOR_A}">${tx('q')}<p188:replyLst>${replies}</p188:replyLst></p188:cm>`;
    const { slide, warnings } = await slideOf(modernDeck(big + mcm('{next}', AUTHOR_A, tx('n'))));
    expect(slide.threads).toHaveLength(1);
    expect(slide.threads?.[0].comments).toHaveLength(10_000);
    expect(warnings.filter((w) => w.includes('more than 10000 comments'))).toHaveLength(1);
  });
});

describe('isoTime', () => {
  it.each([
    ['2026-01-15T09:00:00.000', '2026-01-15T09:00:00.000Z'],
    ['2026-01-15T09:00', '2026-01-15T09:00:00.000Z'],
    ['2026-01-15T09:00:00Z', '2026-01-15T09:00:00.000Z'],
    ['2026-01-15T09:00:00+02:00', '2026-01-15T07:00:00.000Z'],
    ['2026-01-15T09:00:00-0130', '2026-01-15T10:30:00.000Z'],
    [' 2026-01-15T09:00:00.5 ', '2026-01-15T09:00:00.500Z'],
    ['2026-13-45T09:00:00', undefined],
    ['2024-02-31T10:00:00', undefined],
    ['2023-02-29T10:00:00Z', undefined],
    ['2024-02-29T10:00:00Z', '2024-02-29T10:00:00.000Z'],
    ['2024-04-31T10:00:00', undefined],
    ['2024-01-15T24:00:00', undefined],
    ['2024-01-15T10:60:00', undefined],
    ['2024-01-15T23:59:59', '2024-01-15T23:59:59.000Z'],
    ['2024-00-10T10:00:00', undefined],
    ['yesterday', undefined],
    ['', undefined],
    [undefined, undefined],
  ])('%s -> %s', (input, expected) => expect(isoTime(input)).toBe(expected));
});

describe('a slide that fails to read has neither notes nor threads', () => {
  it('is kept as an error placeholder', async () => {
    const d = { ...withNotes(notesPage(body(para('x')))), ...legacyDeck(lcm('0', '1', 'x')) };
    const slide = (await slideOf({ ...d, slideRels: [...(withNotes('').slideRels ?? []), ...(legacyDeck('').slideRels ?? [])], extra: { ...withNotes(notesPage(body(para('x')))).extra, ...legacyDeck(lcm('0', '1', 'x')).extra }, tree: '<p:sp' })).slide;
    expect(slide.error).toBeDefined();
    expect(slide.notes).toBeUndefined();
    expect(slide.threads).toBeUndefined();
  });
});

describe('notes and comments that cannot be read cost the slide nothing else', () => {
  const text = `<p:sp><p:nvSpPr><p:cNvPr id="2" name="s"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>${para('kept')}</p:txBody></p:sp>`;
  const deck = (): Deck => {
    const n = withNotes(notesPage(body(para('n'))));
    const c = legacyDeck(lcm('0', '1', 'c'));
    return { tree: text, slideRels: [...(n.slideRels ?? []), ...(c.slideRels ?? [])], presRels: c.presRels ?? [], extra: { ...n.extra, ...c.extra } };
  };
  /** Makes one method of the package throw for the part named by `part`. */
  const failOn = (method: 'rels' | 'xml', part: string, error: unknown): void => {
    const original = OoxmlPackage.prototype[method] as (this: OoxmlPackage, name: string) => Promise<unknown>;
    jest.spyOn(OoxmlPackage.prototype, method).mockImplementation(async function (this: OoxmlPackage, name: string) {
      if (name === part) throw error;
      return original.call(this, name);
    } as never);
  };

  afterEach(() => jest.restoreAllMocks());

  it('drops unreadable notes, with a warning, and keeps the slide elements and comments', async () => {
    failOn('rels', 'ppt/notesSlides/notesSlide1.xml', new Error('boom'));
    expect((await slideOf(deck())).warnings.some((w) => w.includes('boom'))).toBe(true);
    jest.restoreAllMocks();
    failOn('rels', 'ppt/notesSlides/notesSlide1.xml', 'plain failure');
    expect((await slideOf(deck())).warnings.some((w) => w.includes('speaker notes of slide 1') && w.includes('unknown error'))).toBe(true);
    jest.restoreAllMocks();
    failOn('rels', 'ppt/notesSlides/notesSlide1.xml', new Error('boom'));
    const { slide, warnings } = await slideOf(deck());
    expect(slide.error).toBeUndefined();
    expect(slide.elements).toHaveLength(1);
    expect(slide.notes).toBeUndefined();
    expect(slide.threads).toHaveLength(1);
    expect(warnings.some((w) => w.includes('speaker notes of slide 1') && w.includes('boom'))).toBe(true);
  });

  it('drops unreadable comments, with a warning, and keeps the slide elements and notes', async () => {
    failOn('xml', 'ppt/comments/comment1.xml', new Error('bang'));
    expect((await slideOf(deck())).warnings.some((w) => w.includes('comments of slide 1') && w.includes('bang'))).toBe(true);
    jest.restoreAllMocks();
    failOn('xml', 'ppt/comments/comment1.xml', 'plain failure');
    const { slide, warnings } = await slideOf(deck());
    expect(slide.error).toBeUndefined();
    expect(slide.elements).toHaveLength(1);
    expect(slide.notes?.paragraphs).toHaveLength(1);
    expect(slide.threads).toBeUndefined();
    expect(warnings.some((w) => w.includes('comments of slide 1') && w.includes('unknown error'))).toBe(true);
  });
});

describe('a notes page shared by several slides', () => {
  it('is read for each slide', async () => {
    const zip = await JSZip.loadAsync(await build({ ...withNotes(notesPage(body(para('shared')))) }));
    zip.file('ppt/_rels/presentation.xml.rels', rels([rel('rId1', 'slide', 'slides/slide1.xml'), rel('rId2', 'slide', 'slides/slide2.xml')]));
    zip.file('ppt/presentation.xml', `<p:presentation ${NS}><p:sldIdLst><p:sldId id="256" r:id="rId1"/><p:sldId id="257" r:id="rId2"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>`);
    zip.file('ppt/slides/slide2.xml', await (zip.file('ppt/slides/slide1.xml') as JSZip.JSZipObject).async('string'));
    zip.file('ppt/slides/_rels/slide2.xml.rels', await (zip.file('ppt/slides/_rels/slide1.xml.rels') as JSZip.JSZipObject).async('string'));
    const result = await importPptx(await zip.generateAsync({ type: 'uint8array' }));
    expect(result.slides.map((s) => s.notes?.paragraphs[0].runs[0].text)).toEqual(['shared', 'shared']);
  });
});
