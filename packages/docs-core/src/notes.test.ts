// Footnote and endnote references: the model, the ProseMirror node, the Y document, collaboration blockers and the DOCX
// mapping. The io package tests the file format itself.

import * as Y from 'yjs';
import { docsSchema } from './prosemirror/schema';
import { blocksToPmDoc, proseMirrorToDocument } from './prosemirror/sync';
import { DocumentImpl } from './document';
import { getCollabBlockers } from './collab/blockers';
import { COLLAB_SCHEMA_VERSION, hydrateYDocFromData, serializeYDocToData } from './collab/y-schema';
import { exportToDocx, importFromDocx } from './export/docx';
import { getPlainText } from './blocks/utils';
import { getBlockTextLength, findRunAtOffset, runOffsetToBlockOffset } from './selection';
import { EditorState, TextSelection } from 'prosemirror-state';
import { findNoteRef, insertNoteReference, nextNoteId, noteOrdinalAt, noteTextFragment } from './prosemirror/notes';
import { DEFAULT_PAGE_CONFIG } from './types';
import type { Block, DocumentData, InlineContent, InlineNoteRef, ParagraphBlock } from './types';

const text = (t: string, extra: Partial<Extract<InlineContent, { type: 'text' }>> = {}): InlineContent => ({ type: 'text', text: t, ...extra });
const para = (...content: InlineContent[]): ParagraphBlock => ({ id: 'p', type: 'paragraph', content });
const note = (kind: 'footnote' | 'endnote', id: string, ...content: InlineContent[]): InlineNoteRef => ({ type: 'noteRef', kind, id, content });

const SAMPLE: Block[] = [
  para(
    text('claim'),
    note('footnote', '1', text('See '), text('page 4', { bold: true }), { type: 'link', text: 'link', href: 'https://example.com/n' }, text('\nand more')),
    text(' and'),
    { ...note('endnote', '1', text('end')), customMark: '*' },
    note('footnote', '2')
  ),
];

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

const viaPm = (blocks: Block[]): Block[] => proseMirrorToDocument(blocksToPmDoc(blocks), dataWith([]) as never).sections[0].blocks;
const viaYjs = (blocks: Block[]): Block[] => {
  const ydoc = new Y.Doc();
  hydrateYDocFromData(ydoc, dataWith(blocks));
  return serializeYDocToData(ydoc).sections[0].blocks;
};
const sampleBlocks = (): Block[] => JSON.parse(JSON.stringify(SAMPLE));
const stripIds = (blocks: Block[]) => JSON.parse(JSON.stringify(blocks, (k, v) => (k === 'id' && typeof v === 'string' && v.startsWith('block_') ? undefined : v)));

describe('note references in the ProseMirror document', () => {
  it('round-trip the kind, id, custom mark and the text of the note, with formatting, a link and a line break', () => {
    const out = viaPm(sampleBlocks()) as ParagraphBlock[];
    const refs = out[0].content.filter((c): c is InlineNoteRef => c.type === 'noteRef');
    expect(refs.map((r) => [r.kind, r.id, r.customMark])).toEqual([
      ['footnote', '1', undefined],
      ['endnote', '1', '*'],
      ['footnote', '2', undefined],
    ]);
    expect(refs[0].content).toMatchObject([{ text: 'See ' }, { text: 'page 4', bold: true }, { type: 'link', text: 'link', href: 'https://example.com/n' }, { text: '\nand more' }]);
    expect(refs[2].content).toEqual([]);
  });

  it('is one atom in the document that takes two positions plus the text of the note, and drops content a note cannot hold', () => {
    const pm = blocksToPmDoc([para(note('footnote', '1', text('abc'), note('footnote', 'x', text('nested')) as never, { type: 'contentControl', control: { type: 'richText' }, content: [text('c')] } as never))]);
    const node = pm.firstChild!.firstChild!;
    expect(node.type.name).toBe('note_ref');
    expect(node.type.spec.atom).toBe(true);
    expect(node.nodeSize).toBe(2 + 'abc'.length);
    expect(node.textContent).toBe('abc');
  });

  it('only holds the inline content a note can hold, so a note cannot hold a note or a content control', () => {
    const { paragraph, note_ref: ref, content_control_inline: control } = docsSchema.nodes;
    expect(() => paragraph.create(null, [ref.create({ kind: 'footnote' }, [docsSchema.text('a')])]).check()).not.toThrow();
    expect(() => ref.create(null, [ref.create()]).check()).toThrow();
    expect(() => ref.create(null, [control.create()]).check()).toThrow();
  });

  it('describes itself in the DOM and reads that description back, refusing anything else', () => {
    const spec = docsSchema.nodes.note_ref.spec;
    const dom = (attrs: Record<string, unknown>) => spec.toDOM!(docsSchema.nodes.note_ref.create(attrs)) as [string, Record<string, string>, number];
    expect(dom({ kind: 'endnote', noteId: '3', customMark: '*' })).toEqual(['sup', { class: 'note-ref', 'data-note-ref': 'endnote', 'data-note-id': '3', 'data-custom-mark': '*' }, 0]);
    expect(dom({ kind: 'weird' })[1]['data-note-ref']).toBe('footnote');
    expect(dom({ kind: 'footnote' })[1]).not.toHaveProperty('data-custom-mark');
    const read = (attrs: Record<string, string | null>) => (spec.parseDOM![0].getAttrs as (d: unknown) => unknown)({ getAttribute: (n: string) => attrs[n] ?? null });
    expect(read({ 'data-note-ref': 'endnote', 'data-note-id': '9', 'data-custom-mark': '†' })).toEqual({ kind: 'endnote', noteId: '9', customMark: '†' });
    expect(read({ 'data-note-ref': 'footnote' })).toEqual({ kind: 'footnote', noteId: '', customMark: null });
    expect(read({ 'data-note-ref': 'javascript:1' })).toBe(false);
  });
});

describe('note references in the Y document', () => {
  it('survive hydrate and serialize', () => {
    expect(stripIds(viaYjs(sampleBlocks()))).toEqual(stripIds(viaPm(sampleBlocks())));
    const refs = (viaYjs(sampleBlocks())[0] as ParagraphBlock).content.filter((c) => c.type === 'noteRef');
    expect(refs).toHaveLength(3);
  });

  it('are in the layout version that introduced them', () => {
    expect(COLLAB_SCHEMA_VERSION).toBeGreaterThanOrEqual(4);
  });
});

describe('collab blockers for note references', () => {
  it('knows the note type, so a document with notes is not blocked', () => {
    expect(getCollabBlockers(dataWith(sampleBlocks()))).toEqual([]);
  });

  it('looks inside the note text for types it does not know', () => {
    const data = dataWith([para(note('footnote', '1', { type: 'smartArt' } as never))]);
    expect(getCollabBlockers(data).map((b) => [b.id, b.count])).toEqual([['collab.unsupportedInline', 1]]);
  });
});

describe('note references in the block helpers', () => {
  it('stay out of the plain text of a paragraph and count as one position in the model selection', () => {
    const block = para(text('ab'), note('footnote', '1', text('hidden')));
    expect(getPlainText(block.content)).toBe('ab');
    expect(getBlockTextLength(block)).toBe(3);
    expect(findRunAtOffset(block.content, 2)).toEqual({ runIndex: 1, offsetInRun: 0 });
    expect(runOffsetToBlockOffset(block.content, 2, 0)).toBe(3);
  });
});

describe('note references through DOCX', () => {
  it('round-trip through export and import', async () => {
    const doc = new DocumentImpl('doc', 'Doc');
    doc.setData(dataWith(sampleBlocks()));
    const { bytes } = await exportToDocx(doc);
    const reopened = new DocumentImpl('doc2', 'Doc2');
    await importFromDocx(bytes, reopened);
    const content = (reopened.getSections()[0].blocks[0] as ParagraphBlock).content;
    expect(content.filter((c): c is InlineNoteRef => c.type === 'noteRef')).toMatchObject([
      { kind: 'footnote', id: '1', content: [{ text: 'See ' }, { text: 'page 4', bold: true }, { type: 'link', href: 'https://example.com/n' }, { text: '\nand more' }] },
      { kind: 'endnote', id: '1', customMark: '*', content: [{ text: 'end' }] },
      { kind: 'footnote', id: '2', content: [] },
    ]);
  });

  it('are left out of a header or footer read from a file', async () => {
    const doc = new DocumentImpl('doc', 'Doc');
    doc.setData({
      ...dataWith([para(text('body'))]),
      sections: [
        { id: 's1', pageConfig: DEFAULT_PAGE_CONFIG, blocks: [para(text('body'))], header: { blocks: [{ type: 'paragraph', content: [{ type: 'text', text: 'Head' }] }] } },
      ],
    });
    const { bytes } = await exportToDocx(doc);
    const reopened = new DocumentImpl('doc2', 'Doc2');
    await importFromDocx(bytes, reopened);
    expect(reopened.getSections()[0].header?.blocks[0].content).toEqual([expect.objectContaining({ text: 'Head' })]);
  });
});

describe('finding, inserting and rewriting note references', () => {
  const { doc: docType, paragraph, note_ref: refType, content_control_inline: control } = docsSchema.nodes;
  const ref = (kind: string, id: string, ...inner: ReturnType<typeof docsSchema.text>[]) => refType.create({ kind, noteId: id }, inner);
  const sample = () =>
    docType.create(null, [
      paragraph.create(null, [docsSchema.text('a'), ref('footnote', '4', docsSchema.text('x')), ref('endnote', '9'), ref('footnote', '2')]),
      paragraph.create(null, [ref('footnote', '7'), docsSchema.text('b')]),
    ]);

  it('finds the nth reference of a kind in document order, and says which one stands at a position', () => {
    const doc = sample();
    expect(findNoteRef(doc, 'footnote', 1)?.node.attrs.noteId).toBe('4');
    expect(findNoteRef(doc, 'footnote', 2)?.node.attrs.noteId).toBe('2');
    expect(findNoteRef(doc, 'footnote', 3)?.node.attrs.noteId).toBe('7');
    expect(findNoteRef(doc, 'endnote', 1)?.node.attrs.noteId).toBe('9');
    expect(findNoteRef(doc, 'footnote', 4)).toBeNull();
    expect(findNoteRef(doc, 'endnote', 2)).toBeNull();
    for (const [kind, ordinal] of [['footnote', 1], ['endnote', 1], ['footnote', 2], ['footnote', 3]] as const) {
      expect(noteOrdinalAt(doc, findNoteRef(doc, kind, ordinal)!.pos)).toEqual({ kind, ordinal });
    }
    expect(noteOrdinalAt(doc, 0)).toBeNull();
    expect(noteOrdinalAt(doc, 1)).toBeNull();
    expect(noteOrdinalAt(doc, 999)).toBeNull();
  });

  it('gives a new note an id one above the largest numeric id of its kind', () => {
    const doc = sample();
    expect(nextNoteId(doc, 'footnote')).toBe('8');
    expect(nextNoteId(doc, 'endnote')).toBe('10');
    expect(nextNoteId(docType.create(null, [paragraph.create()]), 'footnote')).toBe('1');
    expect(nextNoteId(docType.create(null, [paragraph.create(null, [ref('footnote', 'abc')])]), 'footnote')).toBe('1');
  });

  it('inserts an empty reference at the end of the selection and puts the caret after it', () => {
    const state = EditorState.create({ doc: docType.create(null, [paragraph.create(null, [docsSchema.text('hello')])]) });
    const inside = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 3)));
    let next = inside;
    expect(insertNoteReference('footnote')(inside, (tr) => (next = inside.apply(tr)))).toBe(true);
    expect(next.doc.firstChild!.toString()).toBe('paragraph("he", note_ref, "llo")');
    expect(next.doc.nodeAt(3)!.attrs).toMatchObject({ kind: 'footnote', noteId: '1' });
    expect(next.selection.from).toBe(3 + next.doc.nodeAt(3)!.nodeSize);
    // A second one before it gets the next id, and a range selection inserts after the range.
    const range = next.apply(next.tr.setSelection(TextSelection.create(next.doc, 1, 2)));
    let third = range;
    insertNoteReference('endnote')(range, (tr) => (third = range.apply(tr)));
    expect(third.doc.nodeAt(2)!.attrs).toMatchObject({ kind: 'endnote', noteId: '1' });
  });

  it('is only applicable where inline content can go, and does nothing without a dispatch', () => {
    const state = EditorState.create({ doc: docType.create(null, [paragraph.create()]) });
    expect(insertNoteReference('footnote')(state)).toBe(true);
    expect(state.doc.firstChild!.childCount).toBe(0);
    const noteless = EditorState.create({ doc: docType.create(null, [docsSchema.nodes.horizontal_rule.create()]), selection: undefined });
    expect(insertNoteReference('endnote')(noteless)).toBe(false);
  });

  it('makes a note text from the note editor document: paragraphs joined by line breaks, only inline content a note holds', () => {
    const doc = docType.create(null, [
      paragraph.create(null, [docsSchema.text('one ', [docsSchema.marks.bold.create()]), ref('footnote', '1', docsSchema.text('nested')), control.create({ controlType: 'text' }, [docsSchema.text('ctl')])]),
      paragraph.create(null, [docsSchema.text('two')]),
      docsSchema.nodes.horizontal_rule.create(),
    ]);
    const fragment = noteTextFragment(docsSchema, doc);
    expect(fragment.toString()).toBe('<bold("one "), "nestedctl", hard_break, "two">');
    expect(fragment.firstChild!.marks.map((m) => m.type.name)).toEqual(['bold']);
    expect(noteTextFragment(docsSchema, docType.create(null, [paragraph.create()])).childCount).toBe(0);
    const inList = docType.create(null, [docsSchema.nodes.bullet_list.create(null, [docsSchema.nodes.list_item.create(null, [paragraph.create(null, [docsSchema.text('item')])])])]);
    expect(noteTextFragment(docsSchema, inList).toString()).toBe('<"item">');
  });
});
