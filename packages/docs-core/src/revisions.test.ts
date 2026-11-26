// Tracked changes (insertion and deletion marks): the model, the ProseMirror marks, the Y document, collaboration
// blockers and the DOCX mapping.

import * as Y from 'yjs';
import { EditorState } from 'prosemirror-state';
import { updateYFragment, yXmlFragmentToProseMirrorRootNode } from 'y-prosemirror';
import JSZip from 'jszip';
import { readFileSync } from 'node:fs';
import path from 'node:path';
import { docsSchema } from './prosemirror/schema';
import { blocksToPmDoc, proseMirrorToDocument, revisionOf } from './prosemirror/sync';
import { readBlockRevision } from './prosemirror/block-revision';
import { DocumentImpl } from './document';
import { getCollabBlockers } from './collab/blockers';
import { COLLAB_SCHEMA_VERSION, CollabSchemaVersionError, assertSchemaVersion, getYDocFields, hydrateYDocFromData, serializeYDocToData } from './collab/y-schema';
import { exportToDocx, importFromDocx } from './export/docx';
import { REVISION_COLORS, assignAuthorColors, authorColor, collectRevisions, documentRevisionAuthors, revisionAuthors, sameRevisionAuthors } from './revisions';
import { DEFAULT_PAGE_CONFIG } from './types';
import type { Block, DocumentData, InlineContent, ParagraphBlock, RunFormat } from './types';

const ADA = { id: '1', author: 'Ada', date: '2024-05-06T07:08:09Z' };
const BO = { id: '2', author: 'Bo' };

const CONTENT: InlineContent[] = [
  { type: 'text', text: 'keep ' },
  { type: 'text', text: 'new ', bold: true, revision: { insertion: ADA } },
  { type: 'text', text: 'old ', revision: { deletion: BO } },
  { type: 'link', text: 'link', href: 'https://example.com/', revision: { insertion: ADA, deletion: BO } },
  { type: 'text', text: ' end' },
];
const PARAGRAPH_PLAIN: ParagraphBlock = { id: 'p0', type: 'paragraph', content: [{ type: 'text', text: 'plain' }] };
const PARAGRAPH: ParagraphBlock = { id: 'p1', type: 'paragraph', content: CONTENT };

function dataWith(blocks: Block[]): DocumentData {
  return {
    id: 'doc',
    title: 'Doc',
    sections: [{ id: 's1', pageConfig: DEFAULT_PAGE_CONFIG, blocks }],
    defaultPageConfig: DEFAULT_PAGE_CONFIG,
    textStylePool: {},
    paragraphStylePool: {},
  };
}
// Block ids differ between copies; the ids of changes (which have no `type`) must survive.
const strip = (blocks: Block[]) =>
  JSON.parse(JSON.stringify(blocks, function (this: object, k, v) { return k === 'id' && 'type' in this ? undefined : v; }));
const viaPm = (blocks: Block[]) => proseMirrorToDocument(blocksToPmDoc(blocks), dataWith([]) as never).sections[0].blocks;
const viaYjs = (blocks: Block[]) => {
  const ydoc = new Y.Doc();
  hydrateYDocFromData(ydoc, dataWith(blocks));
  return serializeYDocToData(ydoc).sections[0].blocks;
};

describe('tracked changes in the ProseMirror document', () => {
  it('become insertion and deletion marks and come back as the same runs', () => {
    const pm = blocksToPmDoc([PARAGRAPH]);
    const marks = (i: number) => pm.firstChild!.child(i).marks.map((m) => [m.type.name, m.attrs.revisionId, m.attrs.author, m.attrs.date]);
    expect(marks(1)).toContainEqual(['insertion', '1', 'Ada', '2024-05-06T07:08:09Z']);
    expect(marks(2)).toEqual([['deletion', '2', 'Bo', null]]);
    expect(strip(viaPm([PARAGRAPH]))).toEqual(strip([PARAGRAPH]));
  });

  it('are read from marks without a date or with a stray one', () => {
    const { insertion, deletion } = docsSchema.marks;
    expect(revisionOf([])).toBeUndefined();
    expect(revisionOf([insertion.create({ revisionId: 5, author: 'A', date: '' })])).toEqual({ insertion: { id: '5', author: 'A' } });
    expect(revisionOf([deletion.create({ revisionId: '6', author: 'B', date: 'x' })])).toEqual({ deletion: { id: '6', author: 'B', date: 'x' } });
  });

  it('have DOM forms that read back, and ignore an element without an id or an author', () => {
    const { insertion, deletion } = docsSchema.marks;
    const mark = insertion.create({ revisionId: '9', author: 'Ada', date: '2024-01-01' });
    expect(mark.type.spec.toDOM!(mark, true)).toEqual([
      'span',
      { 'data-insertion': '9', 'data-revision-author': 'Ada', 'data-revision-date': '2024-01-01', class: 'revision-insertion' },
      0,
    ]);
    const bare = deletion.create({ revisionId: '9', author: 'Bo' });
    expect((bare.type.spec.toDOM!(bare, true) as readonly unknown[])[1]).not.toHaveProperty('data-revision-date');
    const read = (kind: typeof insertion, attrs: Record<string, string | null>) =>
      (kind.spec.parseDOM![0].getAttrs as (d: unknown) => unknown)({ getAttribute: (n: string) => attrs[n] ?? null });
    expect(read(insertion, { 'data-insertion': '9', 'data-revision-author': 'Ada', 'data-revision-date': 'd' })).toEqual({ revisionId: '9', author: 'Ada', date: 'd', moveId: null });
    expect(read(deletion, { 'data-deletion': '9', 'data-revision-author': 'Bo', 'data-revision-move': 'm1' })).toMatchObject({ moveId: 'm1' });
    expect(read(deletion, { 'data-deletion': '9' })).toBe(false);
    // Pasted HTML is not trusted for size: an id, author, date or move name is cut at 200 characters.
    const long = 'y'.repeat(900);
    expect(read(deletion, { 'data-deletion': long, 'data-revision-author': long, 'data-revision-date': long, 'data-revision-move': long })).toEqual({ revisionId: 'y'.repeat(200), author: 'y'.repeat(200), date: 'y'.repeat(200), moveId: 'y'.repeat(200) });
    expect(read(deletion, { 'data-revision-author': 'x' })).toBe(false);
  });

  it('are not inherited by text typed next to them', () => {
    expect(docsSchema.marks.insertion.spec.inclusive).toBe(false);
    expect(docsSchema.marks.deletion.spec.inclusive).toBe(false);
  });
});

describe('line breaks inside tracked runs', () => {
  const BREAKS: [string, InlineContent][] = [
    ['insertion', { type: 'text', text: 'a\nb', revision: { insertion: ADA } }],
    ['deletion', { type: 'text', text: 'a\nb', revision: { deletion: BO } }],
    ['both', { type: 'text', text: 'a\nb', revision: { insertion: ADA, deletion: BO } }],
    ['a lone break', { type: 'text', text: '\n', revision: { deletion: BO } }],
    ['a link', { type: 'link', text: 'a\nb', href: 'https://example.com/', revision: { insertion: ADA } }],
  ];
  it.each(BREAKS)('keep their change on %s, through the editor model, Y and DOCX', async (_name, run) => {
    const paragraph: ParagraphBlock = { id: 'p', type: 'paragraph', content: [{ type: 'text', text: 'x' }, run, { type: 'text', text: 'y' }] };
    const expected = strip([paragraph]);
    expect(strip(viaPm([paragraph]))).toEqual(expected);
    expect(strip(viaYjs([paragraph]))).toEqual(expected);
    const doc = new DocumentImpl('doc', 'Doc');
    doc.setData(dataWith([paragraph]));
    const { bytes } = await exportToDocx(doc);
    const reopened = new DocumentImpl('doc2', 'Doc2');
    await importFromDocx(bytes, reopened);
    expect(strip(reopened.getSections()[0].blocks)[0].content).toMatchObject(expected[0].content);
  });

  it('are written as a break inside the change', async () => {
    const doc = new DocumentImpl('doc', 'Doc');
    doc.setData(dataWith([{ id: 'p', type: 'paragraph', content: [{ type: 'text', text: 'a\nb', revision: { deletion: BO } }] }]));
    const xml = await (await JSZip.loadAsync((await exportToDocx(doc)).bytes)).file('word/document.xml')!.async('string');
    expect(xml).toMatch(/<w:del [^>]*>.*<w:delText[^>]*>a<\/w:delText>.*<w:br\/>.*<w:delText[^>]*>b<\/w:delText>/s);
  });

  it('are kept when the break node carries the change itself', () => {
    const { nodes, marks } = docsSchema;
    const mark = marks.deletion.create({ revisionId: '2', author: 'Bo' });
    const pm = docsSchema.node('doc', null, [nodes.paragraph.create(null, [docsSchema.text('a'), nodes.hard_break.create(null, null, [mark]), docsSchema.text('b')])]);
    expect(strip(proseMirrorToDocument(pm, dataWith([]) as never).sections[0].blocks)[0].content).toEqual([
      { type: 'text', text: 'a' },
      { type: 'text', text: '\n', revision: { deletion: BO } },
      { type: 'text', text: 'b' },
    ]);
  });
});

describe('tracked changes in the Y document and in collaboration', () => {
  it('survive hydrate and serialize', () => {
    expect(strip(viaYjs([PARAGRAPH]))).toEqual(strip([PARAGRAPH]));
  });

  it('are not a blocker for collaboration', () => {
    expect(getCollabBlockers(dataWith([PARAGRAPH]))).toEqual([]);
  });

  describe('survive two clients editing', () => {
    const merge = (a: Y.Doc, b: Y.Doc) => {
      Y.applyUpdate(b, Y.encodeStateAsUpdate(a, Y.encodeStateVector(b)));
      Y.applyUpdate(a, Y.encodeStateAsUpdate(b, Y.encodeStateVector(a)));
    };
    const pair = () => {
      const a = new Y.Doc();
      hydrateYDocFromData(a, dataWith([PARAGRAPH]));
      const b = new Y.Doc();
      Y.applyUpdate(b, Y.encodeStateAsUpdate(a));
      return [a, b] as const;
    };
    // Each client types in the paragraph with a plain ProseMirror transaction pushed back into its Y document; `at` is an
    // offset in the paragraph text, 'start' and 'end' the two ends.
    const type = (ydoc: Y.Doc, at: 'start' | 'end' | number, text: string) => {
      const fragment = getYDocFields(ydoc).content;
      const state = EditorState.create({ doc: yXmlFragmentToProseMirrorRootNode(fragment, docsSchema) });
      const paragraph = state.doc.firstChild!;
      const pos = at === 'start' ? 1 : at === 'end' ? paragraph.nodeSize - 1 : 1 + at;
      const tr = state.tr.insertText(text, pos);
      updateYFragment(ydoc, fragment, tr.doc, { mapping: new Map(), isOMark: new Map() });
    };
    const contentOf = (d: Y.Doc) => (serializeYDocToData(d).sections[0].blocks[0] as ParagraphBlock).content;

    it('on either side of the changes: ids, authors and dates of every run are kept', () => {
      const [a, b] = pair();
      type(a, 'start', '>> ');
      type(b, 'end', ' <<');
      merge(a, b);
      expect(strip([{ ...PARAGRAPH, content: contentOf(a) }])).toEqual(strip([{ ...PARAGRAPH, content: contentOf(b) }]));
      expect(strip([{ ...PARAGRAPH, content: contentOf(a) }])).toEqual(
        strip([{ ...PARAGRAPH, content: [{ type: 'text', text: '>> keep ' }, ...CONTENT.slice(1, 4), { type: 'text', text: ' end <<' }] }])
      );
    });

    it('with one client typing inside a deletion while the other types at the end', () => {
      const [a, b] = pair();
      // "keep new old link end": offset 10 is inside the deletion "old ".
      type(a, 'keep new o'.length, 'X');
      type(b, 'end', '!');
      merge(a, b);
      const content = contentOf(a);
      expect(strip([{ ...PARAGRAPH, content }])).toEqual(strip([{ ...PARAGRAPH, content: contentOf(b) }]));
      // The deletion keeps its change (id, author, no date) and the typed character is inside it, as a plain edit gives.
      const deleted = content.filter((c) => c.type === 'text' && c.revision?.deletion);
      expect(deleted.map((c) => (c as { text: string }).text).join('')).toBe('oXld ');
      expect(deleted.every((c) => c.type === 'text' && c.revision?.deletion?.id === '2' && c.revision.deletion.author === 'Bo')).toBe(true);
      // The insertion and the link keep both of their changes.
      expect(content.find((c) => c.type === 'text' && c.revision?.insertion)).toMatchObject({ text: 'new ', bold: true, revision: { insertion: ADA } });
      expect(content.find((c) => c.type === 'link')).toMatchObject({ revision: { insertion: ADA, deletion: BO } });
      expect((content[content.length - 1] as { text: string }).text).toBe(' end!');
    });
  });
});

describe('author colours and the list of changes', () => {
  it('give an author the same colour every time, from the palette', () => {
    expect(authorColor('Ada Lovelace')).toBe(authorColor('Ada Lovelace'));
    expect(REVISION_COLORS).toContain(authorColor('Ada Lovelace'));
    expect(REVISION_COLORS).toContain(authorColor(''));
    expect(new Set(['Ada', 'Bo', 'Cy', 'Di', 'Ed', 'Flo', 'Gus', 'Hal'].map(authorColor)).size).toBeGreaterThan(3);
  });

  it('gives every author of a document a colour of their own while the palette lasts', () => {
    const names = ['Alice Archer', 'Bram Keller', 'Chidi Okafor', 'Dana Whitfield'];
    const colors = assignAuthorColors(names);
    expect(new Set(colors.values()).size).toBe(4);
    // The same set in another order, and with a repeated name, gives the same colours.
    expect(assignAuthorColors([...names].reverse().concat('Alice Archer'))).toEqual(colors);
    // A lone author keeps the colour that follows from the name; more authors than colours reuse them.
    expect(assignAuthorColors(['Ada']).get('Ada')).toBe(authorColor('Ada'));
    const many = assignAuthorColors(Array.from({ length: 12 }, (_, i) => `author ${i}`));
    expect(many.size).toBe(12);
    expect(new Set(many.values()).size).toBe(REVISION_COLORS.length);
  });

  it('lists the authors of changes recorded on paragraphs and rows, and of formatting changes, too', () => {
    const pm = blocksToPmDoc([
      { id: 'a', type: 'paragraph', content: [{ type: 'text', text: 'x', revision: { formatChange: { id: '5', author: 'Cy', previous: {} } } }] },
      { id: 'b', type: 'paragraph', content: [{ type: 'text', text: 'y' }], revision: { insertion: { id: '6', author: 'Di' }, formatChange: { id: '7', author: 'Ed', previous: {} } } },
      { id: 'c', type: 'heading', level: 1, content: [{ type: 'text', text: 'z' }], revision: { deletion: { id: '8', author: 'Flo' } } },
      { id: 't', type: 'table', rows: [{ id: 'r', cells: [{ id: 'k', content: [{ type: 'text', text: 'c' }] }], revision: { insertion: { id: '9', author: 'Gus' } } }] },
    ] as Block[]);
    expect(documentRevisionAuthors(pm)).toEqual(['Cy', 'Di', 'Ed', 'Flo', 'Gus']);
  });

  it('lists the authors of changes in a document\'s marks', () => {
    const pm = blocksToPmDoc([{ id: 'p', type: 'paragraph', content: [{ type: 'text', text: 'x', revision: { insertion: ADA } }, { type: 'text', text: 'y', revision: { deletion: BO } }, { type: 'text', text: 'z' }] }]);
    expect(documentRevisionAuthors(pm)).toEqual(['Ada', 'Bo']);
    expect(documentRevisionAuthors(blocksToPmDoc([PARAGRAPH_PLAIN]))).toEqual([]);
  });

  it('collects the changes of a document in order, joining the runs of one change', () => {
    const pm = blocksToPmDoc([
      {
        id: 'p',
        type: 'paragraph',
        content: [
          { type: 'text', text: 'a', revision: { insertion: ADA } },
          { type: 'text', text: 'b', bold: true, revision: { insertion: ADA } },
          { type: 'text', text: 'c' },
          { type: 'text', text: 'd', revision: { deletion: BO } },
          { type: 'text', text: 'e', revision: { insertion: ADA } },
        ],
      },
      { id: 'q', type: 'paragraph', content: [{ type: 'text', text: 'f', revision: { insertion: { id: '3', author: 'Cy' } } }] },
    ]);
    const changes = collectRevisions(pm);
    expect(changes.map((c) => [c.kind, c.id, c.author, c.text])).toEqual([
      ['insertion', '1', 'Ada', 'ab'],
      ['deletion', '2', 'Bo', 'd'],
      ['insertion', '1', 'Ada', 'e'],
      ['insertion', '3', 'Cy', 'f'],
    ]);
    expect(changes[0].date).toBe('2024-05-06T07:08:09Z');
    expect(pm.textBetween(changes[0].from, changes[0].to)).toBe('ab');
    expect(revisionAuthors(changes).map((a) => [a.name, a.count, a.color === assignAuthorColors(['Ada', 'Bo', 'Cy']).get(a.name)])).toEqual([
      ['Ada', 2, true],
      ['Bo', 1, true],
      ['Cy', 1, true],
    ]);
  });
});

describe('one id from two authors', () => {
  it('is two changes, not one, as the DOCX writer treats it', () => {
    const pm = blocksToPmDoc([
      {
        id: 'p',
        type: 'paragraph',
        content: [
          { type: 'text', text: 'a', revision: { insertion: { id: '1', author: 'Ada' } } },
          { type: 'text', text: 'b', revision: { insertion: { id: '1', author: 'Bo' } } },
        ],
      },
    ]);
    expect(collectRevisions(pm).map((c) => [c.id, c.author, c.text])).toEqual([
      ['1', 'Ada', 'a'],
      ['1', 'Bo', 'b'],
    ]);
  });
});

describe('the authors list in state', () => {
  const pm = blocksToPmDoc([{ id: 'p', type: 'paragraph', content: [{ type: 'text', text: 'ab', revision: { insertion: ADA } }, { type: 'text', text: 'c', revision: { deletion: BO } }] }]);

  it('is the same list, by names, counts and colours', () => {
    const authors = revisionAuthors(collectRevisions(pm, false));
    expect(sameRevisionAuthors(authors, revisionAuthors(collectRevisions(pm)))).toBe(true);
    expect(sameRevisionAuthors(authors, authors.slice(1))).toBe(false);
    expect(sameRevisionAuthors(authors, [{ ...authors[0], count: 9 }, authors[1]])).toBe(false);
    expect(sameRevisionAuthors(authors, [{ ...authors[0], color: '#000000' }, authors[1]])).toBe(false);
    expect(sameRevisionAuthors([], [])).toBe(true);
  });

  it('can be collected without the text of each change', () => {
    expect(collectRevisions(pm, false).map((c) => [c.author, c.text, c.to - c.from])).toEqual([['Ada', '', 2], ['Bo', '', 1]]);
  });
});

describe('tracked changes in DOCX', () => {
  it('are written and read back, through the editor model', async () => {
    const doc = new DocumentImpl('doc', 'Doc');
    doc.setData(dataWith([PARAGRAPH]));
    const { bytes } = await exportToDocx(doc);
    const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
    expect(xml).toContain('w:author="Ada"');
    expect(xml).toContain('<w:delText');
    const reopened = new DocumentImpl('doc2', 'Doc2');
    await importFromDocx(bytes, reopened);
    const block = reopened.getSections()[0].blocks[0] as ParagraphBlock;
    expect(strip([block])[0].content).toMatchObject([
      { text: 'keep ' },
      { text: 'new ', bold: true, revision: { insertion: ADA } },
      { text: 'old ', revision: { deletion: BO } },
      { type: 'link', text: 'link', revision: { insertion: ADA, deletion: BO } },
      { text: ' end' },
    ]);
  });

  it('are read from a file with several authors and kept for collaboration', async () => {
    const bytes = new Uint8Array(readFileSync(path.resolve(__dirname, '../../../fixtures/docx/revisions-authors.docx')));
    const doc = new DocumentImpl('doc', 'Doc');
    await importFromDocx(bytes, doc);
    const pm = blocksToPmDoc(doc.getSections().flatMap((s) => s.blocks));
    const changes = collectRevisions(pm);
    expect(changes.length).toBeGreaterThan(10);
    expect(revisionAuthors(changes).length).toBe(4);
    expect(changes.some((c) => c.kind === 'deletion')).toBe(true);
    expect(getCollabBlockers(doc.getData())).toEqual([]);
    expect(JSON.stringify(viaYjs(doc.getSections()[0].blocks))).toContain('"revision"');
  });
});


// Moves, formatting changes, and changes of paragraph marks and table rows.
describe('moves, formatting changes, paragraph marks and rows', () => {
  const MOVE_FROM = { id: '5', author: 'Cy', date: '2024-01-02T00:00:00Z', moveId: 'move-1' };
  const MOVE_TO = { id: '6', author: 'Cy', moveId: 'move-1' };
  const FORMAT = { id: '7', author: 'Di', date: '2024-03-04T00:00:00Z', previous: { bold: true, fontSize: 14, color: '#c00000', namedStyleId: 'Strong' } };
  const KINDS: Block[] = [
    {
      id: 'p',
      type: 'paragraph',
      revision: { insertion: { id: '8', author: 'Ada' }, formatChange: { id: '9', author: 'Bo', previous: { alignment: 'right', leftIndent: 24, headingLevel: 2 } } },
      content: [
        { type: 'text', text: 'from', revision: { deletion: MOVE_FROM } },
        { type: 'text', text: 'to', revision: { insertion: MOVE_TO } },
        { type: 'text', text: 'now', italic: true, revision: { formatChange: FORMAT } },
        { type: 'link', text: 'link', href: 'https://example.com/', revision: { insertion: ADA, formatChange: FORMAT } },
      ],
    },
    { id: 'h', type: 'heading', level: 2, content: [{ type: 'text', text: 'H' }], revision: { deletion: { id: '10', author: 'Bo' } } },
    {
      id: 't',
      type: 'table',
      rows: [
        { id: 'r1', cells: [{ id: 'c1', content: [{ type: 'text', text: 'new row' }] }], revision: { insertion: { id: '11', author: 'Ada' } } },
        { id: 'r2', cells: [{ id: 'c2', content: [{ type: 'text', text: 'old row' }] }], revision: { deletion: { id: '12', author: 'Bo' } } },
        { id: 'r3', cells: [{ id: 'c3', content: [{ type: 'text', text: 'row' }] }] },
      ],
    },
  ];

  it('travel through the ProseMirror document and the Y document unchanged', () => {
    for (const via of [viaPm, viaYjs]) {
      const back = via(KINDS);
      expect(strip(back.slice(0, 2))).toEqual(strip(KINDS.slice(0, 2)));
      expect((back[2] as Extract<Block, { type: 'table' }>).rows.map((r) => r.revision)).toEqual([{ insertion: { id: '11', author: 'Ada' } }, { deletion: { id: '12', author: 'Bo' } }, undefined]);
    }
    expect(getCollabBlockers(dataWith(KINDS))).toEqual([]);
  });

  it('are PM marks and attributes of their own', () => {
    const pm = blocksToPmDoc(KINDS);
    const para = pm.child(0);
    expect(para.attrs.revisionInsertion).toMatchObject({ id: '8' });
    expect(para.attrs.revisionDeletion).toBeNull();
    expect(para.attrs.revisionFormat).toMatchObject({ id: '9', previous: { headingLevel: 2 } });
    expect(readBlockRevision(para.attrs)).toMatchObject({ insertion: { id: '8' }, formatChange: { id: '9', previous: { headingLevel: 2 } } });
    expect(para.child(0).marks.map((m) => [m.type.name, m.attrs.moveId])).toEqual([['deletion', 'move-1']]);
    expect(para.child(2).marks.find((m) => m.type.name === 'formatChange')!.attrs).toMatchObject({ revisionId: '7', previous: { bold: true, fontSize: 14 } });
    expect(pm.child(2).child(0).attrs).toMatchObject({ revisionInsertion: { id: '11', author: 'Ada' }, revisionDeletion: null, revisionFormat: null });
    expect(pm.child(2).child(2).attrs).toMatchObject({ revisionInsertion: null, revisionDeletion: null, revisionFormat: null });
    expect(revisionOf(para.child(2).marks)).toEqual({ formatChange: { id: '7', author: 'Di', date: '2024-03-04T00:00:00Z', previous: FORMAT.previous } });
  });

  it('keep what the model gave them apart from the model (no shared objects)', () => {
    const blocks = JSON.parse(JSON.stringify(KINDS)) as Block[];
    const pm = blocksToPmDoc(blocks);
    (blocks[0] as ParagraphBlock).revision!.insertion!.author = 'changed';
    expect(pm.child(0).attrs.revisionInsertion.author).toBe('Ada');
  });

  it('survive a DOCX export and import through the editor model', async () => {
    const doc = new DocumentImpl('doc', 'Doc');
    doc.setData(dataWith(KINDS));
    const { bytes } = await exportToDocx(doc);
    const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
    for (const element of ['<w:moveFrom ', '<w:moveTo ', '<w:rPrChange ', '<w:pPrChange ', '<w:trPr><w:ins ', '<w:trPr><w:del ']) expect(xml).toContain(element);
    const reopened = new DocumentImpl('doc2', 'Doc2');
    await importFromDocx(bytes, reopened);
    const [para, heading, table] = reopened.getSections()[0].blocks;
    expect((para as ParagraphBlock).revision).toMatchObject({ insertion: { id: '8', author: 'Ada' }, formatChange: { id: '9', author: 'Bo', previous: { alignment: 'right', leftIndent: 24 } } });
    expect((para as ParagraphBlock).content).toMatchObject([
      { revision: { deletion: { id: '5', moveId: 'move-1' } } },
      { revision: { insertion: { id: '6', moveId: 'move-1' } } },
      { italic: true, revision: { formatChange: { id: '7', author: 'Di', previous: { bold: true, fontSize: 14, color: '#c00000' } } } },
      { type: 'link', revision: { insertion: { author: 'Ada' }, formatChange: { id: '7' } } },
    ]);
    expect(heading).toMatchObject({ revision: { deletion: { id: '10', author: 'Bo' } } });
    expect((table as Extract<Block, { type: 'table' }>).rows.map((r) => r.revision)).toMatchObject([{ insertion: { id: '11' } }, { deletion: { id: '12' } }, undefined]);
  });

  it('read a formatting change from pasted HTML with known fields of the right type only', () => {
    const { formatChange } = docsSchema.marks;
    const read = (previous: string) => (formatChange.spec.parseDOM![0].getAttrs as (d: unknown) => { previous: unknown } | false)({
      getAttribute: (n: string) => ({ 'data-format-change': '1', 'data-revision-author': 'A', 'data-format-previous': previous })[n] ?? null,
    });
    expect(read('{"bold":true,"fontSize":12,"color":"#112233","fontFamily":"Arial","namedStyleId":"S","evil":1,"italic":"yes","backgroundColor":"url(x)"}')).toMatchObject({
      previous: { bold: true, fontSize: 12, color: '#112233', fontFamily: 'Arial', namedStyleId: 'S' },
    });
    expect(Object.keys((read('{"bold":true,"evil":1,"italic":"yes","backgroundColor":"url(x)","fontSize":-3}') as { previous: object }).previous)).toEqual(['bold']);
    expect(read('not json')).toMatchObject({ previous: {} });
    expect(read('[1]')).toMatchObject({ previous: {} });
    expect(read('null')).toMatchObject({ previous: {} });
    const mark = formatChange.create({ revisionId: '1', author: 'A', previous: { bold: true } });
    expect(formatChange.spec.toDOM!(mark, true)).toEqual(['span', expect.objectContaining({ 'data-format-change': '1', 'data-format-previous': '{"bold":true}', class: 'revision-format-change' }), 0]);
    expect((formatChange.spec.parseDOM![0].getAttrs as (d: unknown) => unknown)({ getAttribute: () => null })).toBe(false);
  });

  it('keep a tracked paragraph mark in a table cell of one paragraph when the document is saved', () => {
    const cellParagraph: ParagraphBlock = { id: 'c', type: 'paragraph', content: [{ type: 'text', text: 'cell' }], revision: { insertion: { id: '4', author: 'Ada' } } };
    const table: Block = { id: 't', type: 'table', rows: [{ id: 'r', cells: [{ id: 'k', content: [{ type: 'text', text: 'cell' }], blocks: [cellParagraph] }] }] } as never;
    const saved = viaPm([table])[0] as Extract<Block, { type: 'table' }>;
    expect(saved.rows[0].cells[0].blocks).toMatchObject([{ type: 'paragraph', revision: { insertion: { id: '4', author: 'Ada' } } }]);
    // Without a change the cell stays the plain one it was.
    const plainParagraph: ParagraphBlock = { id: 'c', type: 'paragraph', content: [{ type: 'text', text: 'cell' }] };
    const plainTable: Block = { id: 't', type: 'table', rows: [{ id: 'r', cells: [{ id: 'k', content: [{ type: 'text', text: 'cell' }], blocks: [plainParagraph] }] }] } as never;
    const plain = viaPm([plainTable])[0] as Extract<Block, { type: 'table' }>;
    expect(plain.rows[0].cells[0].blocks).toBeUndefined();
  });

  it('keep every field of a run format through copy and paste (HTML)', () => {
    // Typed as a record of every field, so a field added to the run format has to be added here (and to the sanitiser).
    const everything: Required<RunFormat> = {
      bold: true, italic: true, underline: true, strikethrough: true, superscript: true, subscript: true, smallCaps: true, allCaps: true, hidden: true,
      fontFamily: 'Arial', fontFamilyEastAsia: 'MS Mincho', fontFamilyComplex: 'Tahoma', fontSize: 13, color: '#112233', backgroundColor: '#445566', namedStyleId: 'Strong',
    };
    const { formatChange } = docsSchema.marks;
    const mark = formatChange.create({ revisionId: '1', author: 'A', previous: everything });
    const dom = formatChange.spec.toDOM!(mark, true) as unknown as [string, Record<string, string>];
    const read = (formatChange.spec.parseDOM![0].getAttrs as (d: unknown) => { previous: unknown })({ getAttribute: (n: string) => dom[1][n] ?? null });
    expect(read.previous).toEqual(everything);
  });
});

describe('the collaboration schema version', () => {
  it('is 7: the formatChange mark and the revision attributes need a client that knows them', () => {
    // 6 added the insertion and deletion marks; a client at 6 drops formatChange marks and the new attributes.
    expect(COLLAB_SCHEMA_VERSION).toBe(7);
  });

  it('is what a room is written with, and a client that supports less refuses it', () => {
    const ydoc = new Y.Doc();
    hydrateYDocFromData(ydoc, dataWith([PARAGRAPH]));
    expect(ydoc.getMap('meta').get('schemaVersion')).toBe(7);
    expect(() => assertSchemaVersion(ydoc)).not.toThrow();
    // A build from before the formatChange mark (supporting 6) refuses the room.
    let refused: unknown;
    try {
      assertSchemaVersion(ydoc, 6);
    } catch (e) {
      refused = e;
    }
    expect(refused).toBeInstanceOf(CollabSchemaVersionError);
    expect(refused).toMatchObject({ roomVersion: 7, supportedVersion: 6 });
  });
});

describe('change records that come from another client are not trusted', () => {
  it('keeps known fields of the right type in the formatting before a change, for a paragraph and for text', () => {
    const para = docsSchema.nodes.paragraph.create({
      revisionFormat: {
        id: '1',
        author: 'Bo',
        previous: { alignment: 'diagonal', leftIndent: 'wide', rightIndent: Infinity, spaceBefore: 12, keepTogether: 'yes', lineSpacing: { type: 'exactly', value: 'x' }, headingLevel: 9, namedStyleId: 'Callout', evil: 1 },
      },
    }, [docsSchema.text('x')]);
    expect(readBlockRevision(para.attrs)!.formatChange!.previous).toEqual({ spaceBefore: 12, lineSpacing: { type: 'exactly' }, namedStyleId: 'Callout' });
    // Not an object at all.
    expect(readBlockRevision({ revisionFormat: { id: '1', author: 'Bo', previous: 'x' } })!.formatChange!.previous).toEqual({});
    const mark = docsSchema.marks.formatChange.create({ revisionId: '2', author: 'Bo', previous: { bold: true, fontSize: 'big', color: 'url(x)', extra: true } });
    const back = revisionOf([mark])!;
    expect(back.formatChange!.previous).toEqual({ bold: true });
  });

});
