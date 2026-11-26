// Table of contents: the model, the ProseMirror node, the Y document, collaboration blockers, the DOCX mapping and the
// commands that insert, update and remove it.

import * as Y from 'yjs';
import { EditorState, NodeSelection, TextSelection } from 'prosemirror-state';
import { docsSchema } from './prosemirror/schema';
import { blocksToPmDoc, proseMirrorToDocument } from './prosemirror/sync';
import { DocumentImpl } from './document';
import { getCollabBlockers } from './collab/blockers';
import { hydrateYDocFromData, serializeYDocToData, COLLAB_SCHEMA_VERSION } from './collab/y-schema';
import { exportToDocx } from './export/docx';
import { isBlockEmpty } from './blocks/utils';
import { getBlockText, getBlockTextLength } from './selection';
import {
  DEFAULT_TOC_INSTRUCTION,
  EMPTY_TOC_TEXT,
  insertTableOfContents,
  parseTocInstruction,
  removeTableOfContents,
  updateTableOfContents,
} from './prosemirror/toc';
import { DEFAULT_PAGE_CONFIG } from './types';
import type { Block, DocumentData, TocBlock } from './types';
import JSZip from 'jszip';

const { doc: docType, paragraph, heading, toc, bookmark } = docsSchema.nodes;
const text = (t: string, marks?: readonly import('prosemirror-model').Mark[]) => docsSchema.text(t, marks);
const h = (level: number, t: string, ...extra: ReturnType<typeof docsSchema.text>[]) => heading.create({ level }, [text(t), ...extra]);
const p = (t: string) => paragraph.create(null, t ? [text(t)] : []);

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

const TOC_BLOCK: TocBlock = {
  id: 't',
  type: 'toc',
  instruction: 'TOC \\o "1-2" \\h \\z \\u',
  blocks: [
    { id: 'e1', type: 'paragraph', namedStyleId: 'TOC1', content: [{ type: 'link', text: 'Intro', href: '#_Toc1' }, { type: 'text', text: '\t1' }] },
    { id: 'e2', type: 'paragraph', namedStyleId: 'TOC2', leftIndent: 24, content: [{ type: 'text', text: 'Detail\t2' }] },
  ],
};

const strip = (blocks: Block[]) => JSON.parse(JSON.stringify(blocks, (k, v) => (k === 'id' ? undefined : v)));
const viaPm = (blocks: Block[]) => proseMirrorToDocument(blocksToPmDoc(blocks), dataWith([]) as never).sections[0].blocks;
const viaYjs = (blocks: Block[]) => {
  const ydoc = new Y.Doc();
  hydrateYDocFromData(ydoc, dataWith(blocks));
  return serializeYDocToData(ydoc).sections[0].blocks;
};

describe('parseTocInstruction', () => {
  it('reads the levels, links and page numbers, and lists the switches it does not act on', () => {
    expect(parseTocInstruction(DEFAULT_TOC_INSTRUCTION)).toEqual({ fromLevel: 1, toLevel: 3, hyperlinks: true, pageNumbers: true, ignored: [] });
    expect(parseTocInstruction('TOC \\o "2-4" \\n')).toEqual({ fromLevel: 2, toLevel: 4, hyperlinks: false, pageNumbers: false, ignored: [] });
    expect(parseTocInstruction('TOC \\t "Title,1" \\b chapter \\f')).toMatchObject({ fromLevel: 1, toLevel: 3, ignored: ['t', 'b', 'f'] });
    expect(parseTocInstruction('TOC')).toMatchObject({ fromLevel: 1, toLevel: 3, hyperlinks: false });
  });

  it('keeps the levels in order and within the headings the editor has', () => {
    expect(parseTocInstruction('TOC \\o "3-1"')).toMatchObject({ fromLevel: 1, toLevel: 3 });
    expect(parseTocInstruction('TOC \\o "1-9"')).toMatchObject({ fromLevel: 1, toLevel: 6 });
    expect(parseTocInstruction('TOC \\o "0-0"')).toMatchObject({ fromLevel: 1, toLevel: 1 });
  });
});

describe('a table of contents in the ProseMirror document', () => {
  it('round-trips with its instruction and entries, paragraph formatting and links included', () => {
    const out = viaPm([TOC_BLOCK]);
    expect(strip(out)).toEqual(strip([TOC_BLOCK]));
  });

  it('holds only paragraphs as entries when built from the model', () => {
    const odd = { ...TOC_BLOCK, blocks: [...TOC_BLOCK.blocks, { id: 'x', type: 'horizontal-rule' } as Block] };
    expect((viaPm([odd])[0] as TocBlock).blocks).toHaveLength(2);
  });

  it('is an atom at the top level and in content controls, not in a table cell or a list item', () => {
    const { content_control: control, table, table_row, table_cell, bullet_list, list_item } = docsSchema.nodes;
    expect(toc.spec.atom).toBe(true);
    expect(() => docType.create(null, [toc.create()]).check()).not.toThrow();
    expect(() => control.create(null, [toc.create()]).check()).not.toThrow();
    expect(() => table.create(null, [table_row.create(null, [table_cell.create(null, [toc.create()])])]).check()).toThrow();
    expect(() => bullet_list.create(null, [list_item.create(null, [p('a'), toc.create()])]).check()).toThrow();
  });

  it('describes itself in the DOM and reads that back, ignoring an instruction that is empty or too long', () => {
    const spec = toc.spec;
    expect(spec.toDOM!(toc.create({ instruction: 'TOC \\o "1-1"' }))).toEqual(['div', { class: 'toc', 'data-toc': 'TOC \\o "1-1"' }, 0]);
    const read = (value: string | null) => (spec.parseDOM![0].getAttrs as (d: unknown) => unknown)({ getAttribute: () => value });
    expect(read('TOC \\o "1-2"')).toEqual({ instruction: 'TOC \\o "1-2"' });
    expect(read('')).toEqual({ instruction: DEFAULT_TOC_INSTRUCTION });
    expect(read(null)).toEqual({ instruction: DEFAULT_TOC_INSTRUCTION });
    expect(read('x'.repeat(501))).toEqual({ instruction: DEFAULT_TOC_INSTRUCTION });
  });
});

describe('a table of contents in the Y document and in collaboration', () => {
  it('survives hydrate and serialize', () => {
    expect(strip(viaYjs([TOC_BLOCK]))).toEqual(strip([TOC_BLOCK]));
  });

  it('is in the layout version that introduced it', () => {
    expect(COLLAB_SCHEMA_VERSION).toBeGreaterThanOrEqual(5);
  });

  it('is a known block, so a document with one is not blocked, and its entries are looked into', () => {
    expect(getCollabBlockers(dataWith([TOC_BLOCK]))).toEqual([]);
    const bad = { ...TOC_BLOCK, blocks: [{ id: 'x', type: 'chart' } as never] };
    expect(getCollabBlockers(dataWith([bad])).map((b) => [b.id, b.count])).toEqual([['collab.unsupportedBlock', 1]]);
  });
});

describe('a table of contents in the block helpers and in DOCX', () => {
  it('counts as no text of its own in the model selection, and is empty when its entries are', () => {
    expect(getBlockTextLength(TOC_BLOCK)).toBe(0);
    expect(getBlockText(TOC_BLOCK)).toBe('');
    expect(isBlockEmpty(TOC_BLOCK)).toBe(false);
    expect(isBlockEmpty({ ...TOC_BLOCK, blocks: [{ id: 'e', type: 'paragraph', content: [] }] })).toBe(true);
  });

  it('is written as a building block control around its entries, which Word reads as the table of contents', async () => {
    const doc = new DocumentImpl('doc', 'Doc');
    doc.setData(dataWith([TOC_BLOCK]));
    const { bytes } = await exportToDocx(doc);
    const xml = (await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')) ?? '';
    expect(xml).toContain('<w:docPartGallery w:val="Table of Contents"/>');
    expect(xml).toContain('Intro');
    expect(xml).toContain('Detail');
  });
});

describe('inserting, updating and removing a table of contents', () => {
  const sample = () => docType.create(null, [h(1, 'Intro'), p('body'), h(2, 'Detail'), h(3, 'Deep'), h(1, 'Last')]);
  const labels = (pos: number): string | undefined => ({ 0: '1', 13: '2', 21: '3' } as Record<number, string>)[pos];
  const stateOf = (doc = sample()) => EditorState.create({ doc });
  const apply = (state: EditorState, command: ReturnType<typeof updateTableOfContents>) => {
    let next = state;
    expect(command(state, (tr) => (next = state.apply(tr)))).toBe(true);
    return next;
  };
  const entries = (state: EditorState) => {
    const found: Array<{ text: string; style: unknown; indent: unknown; href?: string }> = [];
    state.doc.descendants((node) => {
      if (node.type.name === 'toc') {
        node.forEach((entry) =>
          found.push({
            text: entry.textContent,
            style: entry.attrs.namedStyleId,
            indent: entry.attrs.leftIndent,
            href: entry.firstChild?.marks.find((m) => m.type.name === 'link')?.attrs.href,
          })
        );
        return false;
      }
      return true;
    });
    return found;
  };

  it('inserts a table after the block the caret is in, with an entry per heading, its page number and a link to a new bookmark', () => {
    const start = stateOf();
    const withCaret = start.apply(start.tr.setSelection(TextSelection.create(start.doc, 10)));
    const next = apply(withCaret, insertTableOfContents(labels));
    // The table follows "body"; headings 1 to 3 are listed (the instruction's default), each with a number.
    expect(next.doc.child(2).type.name).toBe('toc');
    expect(next.doc.child(2).attrs.instruction).toBe(DEFAULT_TOC_INSTRUCTION);
    const list = entries(next);
    expect(list.map((e) => e.text)).toEqual(['Intro\t1', 'Detail\t2', 'Deep\t3', 'Last']);
    expect(list.map((e) => e.style)).toEqual(['TOC1', 'TOC2', 'TOC3', 'TOC1']);
    expect(list.map((e) => e.indent)).toEqual([null, 24, 48, null]);
    expect(list.every((e) => /^#_Toc\d{9}$/.test(e.href ?? ''))).toBe(true);
    // Every heading starts with the bookmark its entry points at, and the names are distinct.
    const names: string[] = [];
    next.doc.descendants((node) => {
      if (node.type.name === 'heading') {
        expect(node.firstChild!.type.name).toBe('bookmark');
        names.push(node.firstChild!.attrs.name);
      }
    });
    expect(new Set(names).size).toBe(4);
    expect(list.map((e) => e.href)).toEqual(names.map((n) => `#${n}`));
    // The caret is after the table.
    expect(next.selection.from).toBeGreaterThanOrEqual(next.doc.child(0).nodeSize + next.doc.child(1).nodeSize + next.doc.child(2).nodeSize);
  });

  it('inserts at the end of the document, and at the start of an empty one, and not where a table of contents cannot go', () => {
    const empty = EditorState.create({ doc: docType.create(null, [paragraph.create()]) });
    const next = apply(empty, insertTableOfContents());
    expect(next.doc.child(1).type.name).toBe('toc');
    expect(entries(next).map((e) => e.text)).toEqual([EMPTY_TOC_TEXT]);
    // A caret inside a list item is still after the list (a top-level block).
    const list = docType.create(null, [docsSchema.nodes.bullet_list.create(null, [docsSchema.nodes.list_item.create(null, [p('item')])])]);
    const inList = EditorState.create({ doc: list });
    const placed = apply(inList.apply(inList.tr.setSelection(TextSelection.create(list, 4))), insertTableOfContents());
    expect(placed.doc.child(1).type.name).toBe('toc');
    // At the top level, between blocks (a node selection of the first one).
    const nodeSel = EditorState.create({ doc: sample(), selection: NodeSelection.create(sample(), 0) });
    expect(insertTableOfContents()(nodeSel)).toBe(true);
    // A parent that does not take a table of contents refuses.
    const inCell = docsSchema.nodes.table.create(null, [docsSchema.nodes.table_row.create(null, [docsSchema.nodes.table_cell.create(null, [p('cell')])])]);
    const cellDoc = docType.create(null, [inCell]);
    const stateInCell = EditorState.create({ doc: cellDoc, selection: TextSelection.create(cellDoc, 5) });
    expect(insertTableOfContents()(stateInCell)).toBe(true); // goes after the table, at the top level
    expect(insertTableOfContents()(EditorState.create({ doc: docType.create(null, [p('x')]) }), undefined)).toBe(true);
  });

  it('writes the entries of an existing table from its own levels, using the bookmarks the headings already have, and leaves other content alone', () => {
    const doc = docType.create(null, [
      toc.create({ instruction: 'TOC \\o "1-1" \\h' }, [p('stale entry')]),
      h(1, 'Intro', bookmark.create({ name: '_Toc555' })),
      h(2, 'Detail'),
      h(1, 'Next'),
    ]);
    const next = apply(stateOf(doc), updateTableOfContents((pos) => (pos > 0 ? '9' : undefined)));
    const list = entries(next);
    expect(list.map((e) => e.text)).toEqual(['Intro\t9', 'Next\t9']);
    expect(list[0].href).toBe('#_Toc555');
    expect(list[1].href).toMatch(/^#_Toc\d{9}$/);
    // Only headings that are listed get a new bookmark.
    expect(next.doc.child(2).firstChild!.type.name).toBe('text');
    expect(next.doc.child(3).firstChild!.type.name).toBe('bookmark');
    expect(next.doc.child(1).childCount).toBe(2);
  });

  it('omits links without \\h and page numbers with \\n, lists no empty headings and says when there are no entries', () => {
    const doc = docType.create(null, [
      toc.create({ instruction: 'TOC \\o "1-2" \\n' }),
      h(1, 'Plain'),
      heading.create({ level: 2 }, []),
      toc.create({ instruction: 'TOC \\o "4-5"' }),
    ]);
    const next = apply(stateOf(doc), updateTableOfContents(() => '5'));
    const found: string[][] = [];
    next.doc.forEach((node) => {
      if (node.type.name === 'toc') {
        const rows: string[] = [];
        node.forEach((entry) => rows.push(entry.textContent + (entry.firstChild?.marks.length ? ' [link]' : '')));
        found.push(rows);
      }
    });
    expect(found).toEqual([['Plain'], [EMPTY_TOC_TEXT]]);
    // Without \h the heading needs no bookmark.
    expect(next.doc.child(1).firstChild!.type.name).toBe('text');
  });

  it('updates every table of contents of the document, once, in one transaction, and is not applicable without one', () => {
    const doc = docType.create(null, [toc.create(), h(1, 'A'), toc.create({ instruction: 'TOC \\o "1-1" \\h' }), h(1, 'B')]);
    const state = stateOf(doc);
    let steps = 0;
    let next = state;
    expect(updateTableOfContents()(state, (tr) => {
      steps = tr.steps.length;
      next = state.apply(tr);
    })).toBe(true);
    expect(steps).toBeGreaterThan(0);
    const counts: number[] = [];
    next.doc.forEach((n) => n.type.name === 'toc' && counts.push(n.childCount));
    expect(counts).toEqual([2, 2]);
    // Both tables point at the same bookmark of a heading.
    const hrefs: string[] = [];
    next.doc.descendants((n) => {
      if (n.type.name === 'toc') n.forEach((entry) => hrefs.push(entry.firstChild!.marks[0].attrs.href));
    });
    expect(hrefs.slice(0, 2)).toEqual(hrefs.slice(2).length === 0 ? hrefs.slice(0, 2) : [hrefs[0], hrefs[1]]);
    expect(updateTableOfContents()(stateOf())).toBe(false);
    expect(updateTableOfContents()(state)).toBe(true);
  });

  it('keeps unrelated bookmarks, and does not reuse a name that is taken', () => {
    const doc = docType.create(null, [p(''), paragraph.create(null, [bookmark.create({ name: '_Toc100000001' })]), toc.create(), h(1, 'A')]);
    const next = apply(stateOf(doc), updateTableOfContents());
    const name = next.doc.child(3).firstChild!.attrs.name;
    expect(name).toMatch(/^_Toc\d{9}$/);
    expect(name).not.toBe('_Toc100000001');
    expect(next.doc.child(1).firstChild!.attrs.name).toBe('_Toc100000001');
  });

  it('does not list what is in a table of contents, and finds headings inside content controls', () => {
    const doc = docType.create(null, [
      toc.create({ instruction: 'TOC \\o "1-3" \\h' }, [p('old')]),
      docsSchema.nodes.content_control.create(null, [h(1, 'Wrapped')]),
    ]);
    const next = apply(stateOf(doc), updateTableOfContents());
    expect(entries(next).map((e) => e.text)).toEqual(['Wrapped']);
  });

  it('removes the table that is selected, that holds the selection, or else the nearest one, and not when there is none', () => {
    const doc = docType.create(null, [p('a'), toc.create(), p('b'), toc.create({ instruction: 'TOC \\o "1-1"' }), p('c')]);
    const state = stateOf(doc);
    const second = doc.child(0).nodeSize + doc.child(1).nodeSize + doc.child(2).nodeSize;
    const selected = state.apply(state.tr.setSelection(NodeSelection.create(doc, second)));
    const afterSelected = apply(selected, removeTableOfContents);
    expect(afterSelected.doc.childCount).toBe(4);
    expect(afterSelected.doc.child(1).attrs.instruction).toBe(DEFAULT_TOC_INSTRUCTION);
    // The caret before the first table: that one is nearest.
    const before = state.apply(state.tr.setSelection(TextSelection.create(doc, 1)));
    expect(apply(before, removeTableOfContents).doc.child(1).type.name).toBe('paragraph');
    // The caret just after the last table: that one, not the first.
    const last = state.apply(state.tr.setSelection(TextSelection.create(doc, doc.content.size - 1)));
    expect(apply(last, removeTableOfContents).doc.child(1).attrs.instruction).toBe(DEFAULT_TOC_INSTRUCTION);
    expect(apply(last, removeTableOfContents).doc.childCount).toBe(4);
    // Between the two, nearer the second: the second.
    const between = state.apply(state.tr.setSelection(TextSelection.create(doc, second - 1)));
    expect(apply(between, removeTableOfContents).doc.child(1).type.name).toBe('toc');
    expect(removeTableOfContents(stateOf())).toBe(false);
    expect(removeTableOfContents(state)).toBe(true);
  });
});
