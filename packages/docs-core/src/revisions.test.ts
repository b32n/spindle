// Tracked changes (insertion and deletion marks): the model, the ProseMirror marks, the Y document, collaboration
// blockers and the DOCX mapping.

import * as Y from 'yjs';
import JSZip from 'jszip';
import { readFileSync } from 'node:fs';
import path from 'node:path';
import { docsSchema } from './prosemirror/schema';
import { blocksToPmDoc, proseMirrorToDocument, revisionOf } from './prosemirror/sync';
import { DocumentImpl } from './document';
import { getCollabBlockers } from './collab/blockers';
import { hydrateYDocFromData, serializeYDocToData } from './collab/y-schema';
import { exportToDocx, importFromDocx } from './export/docx';
import { REVISION_COLORS, assignAuthorColors, authorColor, collectRevisions, documentRevisionAuthors, revisionAuthors } from './revisions';
import { DEFAULT_PAGE_CONFIG } from './types';
import type { Block, DocumentData, InlineContent, ParagraphBlock } from './types';

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
    expect(read(insertion, { 'data-insertion': '9', 'data-revision-author': 'Ada', 'data-revision-date': 'd' })).toEqual({ revisionId: '9', author: 'Ada', date: 'd' });
    expect(read(deletion, { 'data-deletion': '9' })).toBe(false);
    expect(read(deletion, { 'data-revision-author': 'x' })).toBe(false);
  });

  it('are not inherited by text typed next to them', () => {
    expect(docsSchema.marks.insertion.spec.inclusive).toBe(false);
    expect(docsSchema.marks.deletion.spec.inclusive).toBe(false);
  });
});

describe('tracked changes in the Y document and in collaboration', () => {
  it('survive hydrate and serialize', () => {
    expect(strip(viaYjs([PARAGRAPH]))).toEqual(strip([PARAGRAPH]));
  });

  it('are not a blocker for collaboration', () => {
    expect(getCollabBlockers(dataWith([PARAGRAPH]))).toEqual([]);
  });

  it('survive a document with two clients editing around them', () => {
    const a = new Y.Doc();
    hydrateYDocFromData(a, dataWith([PARAGRAPH]));
    const b = new Y.Doc();
    Y.applyUpdate(b, Y.encodeStateAsUpdate(a));
    expect(strip(serializeYDocToData(b).sections[0].blocks)).toEqual(strip([PARAGRAPH]));
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
