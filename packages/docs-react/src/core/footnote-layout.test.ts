import { applyPageNotes, createFootnoteHooks, footnoteRefs, type FootnotePlan } from './footnote-layout';
import { computeTrueLayout, type PageConfig } from './true-layout-engine';
import { computeSectionedLayout } from './section-layout';
import type { FlowBlock, ParagraphBlock, Run, TableBlock } from './flow-blocks';
import type { LineMeasure, Measure, ParagraphMeasure } from './measurer';

// A page 700 tall with 50px margins: the content area is 600px.
const PAGE: PageConfig = { width: 800, height: 700, margins: { top: 50, right: 50, bottom: 50, left: 50 } };
const CONTENT = 600;

const lineOf = (height: number, segments?: LineMeasure['segments']): LineMeasure => ({ height, width: 100, ascent: height * 0.8, descent: height * 0.2, yOffset: 0, ...(segments && { segments }) });
const measureOf = (id: string, heights: number[], segments: Array<LineMeasure['segments']> = [], spaceAfter = 0): ParagraphMeasure => ({
  kind: 'paragraph',
  blockId: id,
  lines: heights.map((h, i) => lineOf(h, segments[i])),
  totalHeight: heights.reduce((a, b) => a + b, 0) + spaceAfter,
  spaceBefore: 0,
  spaceAfter,
});
const seg = (runIndex: number) => ({ runIndex, startOffset: 0, endOffset: 1, text: 'x' });
const ref = (ordinal: number, kind: 'footnote' | 'endnote' = 'footnote'): Run => ({ kind: 'text', text: String(ordinal), superscript: true, noteRef: { kind, ordinal } });
const para = (id: string, runs: Run[] = [{ kind: 'text', text: 'text' }]): ParagraphBlock => ({ kind: 'paragraph', id, runs });
const noteBlock = (ordinal: number): ParagraphBlock => ({ kind: 'paragraph', id: `note${ordinal}`, synthetic: true, noteTarget: { kind: 'footnote', ordinal }, runs: [ref(ordinal), { kind: 'text', text: ' note text' }] });

interface Fixture {
  blocks: FlowBlock[];
  measures: Measure[];
  plan: FootnotePlan;
}

/**
 * A body of `lines` 20px lines per block, `refLines` giving, per block, the lines that hold a footnote reference (numbered in
 * order), and one note per reference with the given line counts (20px lines, nothing else).
 */
function fixture(body: Array<{ lines: number; refLines?: number[] }>, noteLines: number[], over: Partial<FootnotePlan> = {}): Fixture {
  const blocks: FlowBlock[] = [];
  const measures: Measure[] = [];
  let ordinal = 0;
  body.forEach((b, i) => {
    const runs: Run[] = [{ kind: 'text', text: 'text' }];
    const segments: Array<LineMeasure['segments']> = [];
    (b.refLines ?? []).forEach((line) => {
      runs.push(ref(++ordinal));
      segments[line] = [...(segments[line] ?? []), seg(runs.length - 1)];
    });
    blocks.push(para(`b${i}`, runs));
    measures.push(measureOf(`b${i}`, Array(b.lines).fill(20), segments));
  });
  const bodyCount = blocks.length;
  const noteBlocks = new Map<string, { index: number; id: string }>();
  noteLines.forEach((n, k) => {
    blocks.push(noteBlock(k + 1));
    measures.push(measureOf(`note${k + 1}`, Array(n).fill(20)));
    noteBlocks.set(`footnote:${k + 1}`, { index: bodyCount + k, id: `note${k + 1}` });
  });
  return { blocks, measures, plan: { refs: footnoteRefs(blocks, measures, bodyCount), noteBlocks, separatorHeight: 10, ...over } };
}

function lay(f: Fixture, bodyCount: number, scale = 1) {
  const hooks = createFootnoteHooks(f.plan, f.measures, scale);
  const base = computeTrueLayout(f.blocks.slice(0, bodyCount), f.measures.slice(0, bodyCount), { pageConfig: PAGE, scale, pageGap: 24, minLinesAtBreak: 1, footnotes: hooks });
  return applyPageNotes(base, f.plan);
}

const noteFragments = (layout: ReturnType<typeof lay>, page: number) => layout.pages[page].fragments.filter((fr) => fr.blockId.startsWith('note'));
const bodyFragments = (layout: ReturnType<typeof lay>, page: number) => layout.pages[page].fragments.filter((fr) => !fr.blockId.startsWith('note'));

describe('footnoteRefs', () => {
  it('finds the line each reference is on, in order, with several on one line', () => {
    const f = fixture([{ lines: 5, refLines: [1, 1, 3] }], []);
    expect(f.plan.refs.get(0)).toEqual([
      { line: 1, key: 'footnote:1' },
      { line: 1, key: 'footnote:2' },
      { line: 3, key: 'footnote:3' },
    ]);
  });

  it('puts every reference of a block with no line data on line 0, and one the lines do not show on line 0 too', () => {
    const block = para('p', [{ kind: 'text', text: 'a' }, ref(1)]);
    const noLines = { kind: 'image', blockId: 'p', totalHeight: 10, width: 10 } as Measure;
    expect(footnoteRefs([block], [noLines], 1).get(0)).toEqual([{ line: 0, key: 'footnote:1' }]);
    expect(footnoteRefs([block], [measureOf('p', [20, 20])], 1).get(0)).toEqual([{ line: 0, key: 'footnote:1' }]);
  });

  it('gives a table, which is not split, the footnotes of all its cells, nested ones included, on its first line', () => {
    const cell = (blocks: FlowBlock[]) => ({ id: 'c', blocks });
    const inner: TableBlock = { kind: 'table', id: 'inner', rows: [{ id: 'r', cells: [cell([para('x', [ref(2)])])] }] };
    const table: TableBlock = { kind: 'table', id: 't', rows: [{ id: 'r', cells: [cell([para('y', [ref(1)]), inner]), cell([para('z', [ref(3)])])] }] };
    const measure = { kind: 'table', blockId: 't', totalHeight: 50, rowHeights: [50] } as Measure;
    expect(footnoteRefs([table], [measure], 1).get(0)).toEqual([
      { line: 0, key: 'footnote:1' },
      { line: 0, key: 'footnote:2' },
      { line: 0, key: 'footnote:3' },
    ]);
  });

  it('ignores endnotes, blocks without references, and the blocks after the body', () => {
    const blocks: FlowBlock[] = [para('a', [ref(1, 'endnote')]), para('b'), { kind: 'heading', id: 'h', level: 1, runs: [ref(1)] }, { kind: 'horizontalRule', id: 'hr' }, para('notes', [ref(9)])];
    const measures: Measure[] = [measureOf('a', [20]), measureOf('b', [20]), measureOf('h', [20], [[seg(0)]]), { kind: 'horizontalRule', blockId: 'hr', totalHeight: 5 } as Measure, measureOf('notes', [20])];
    const refs = footnoteRefs(blocks, measures, 4);
    expect([...refs.keys()]).toEqual([2]);
    const list: FlowBlock = { kind: 'listItem', id: 'li', listType: 'bullet', level: 0, runs: [ref(4)] };
    expect(footnoteRefs([list], [measureOf('li', [20], [[seg(0)]])], 1).get(0)).toEqual([{ line: 0, key: 'footnote:4' }]);
  });
});

describe('footnotes at the bottom of their pages', () => {
  it('lays the body out as it would without notes when no line refers to one', () => {
    const f = fixture([{ lines: 40 }], []);
    const layout = lay(f, 1);
    expect(layout.pages.every((p) => p.footnotes === undefined && p.notes === undefined)).toBe(true);
    const plain = computeTrueLayout(f.blocks.slice(0, 1), f.measures.slice(0, 1), { pageConfig: PAGE, scale: 1, pageGap: 24, minLinesAtBreak: 1 });
    expect(layout.pages.map((p) => p.fragments)).toEqual(plain.pages.map((p) => p.fragments));
  });

  it('puts the note of a reference at the bottom of its page, under a separator', () => {
    // 10 body lines (200px), the reference on line 2, one 2-line note (40px) and a 10px separator.
    const f = fixture([{ lines: 10, refLines: [2] }], [2]);
    const layout = lay(f, 1);
    expect(layout.pages).toHaveLength(1);
    const [note] = noteFragments(layout, 0);
    expect(note).toMatchObject({ blockId: 'note1', fromLine: 0, toLine: 2, height: 40, isFirstFragment: true, isLastFragment: true });
    // The area is 50px tall at the bottom of the 600px content area: separator at 550, the note at 560.
    expect(layout.pages[0].footnotes).toEqual({ top: CONTENT - 50, separatorHeight: 10, continued: false });
    expect(note.y).toBe(CONTENT - 40);
  });

  it('moves a line whose reference would not leave room for its note to the next page, with the note', () => {
    // 29 lines of body fill 580px; the reference is on the last one: with a 2-line note (+10px separator) the page would
    // need 630px, so the last line goes to page 2 and the note follows it.
    const f = fixture([{ lines: 29, refLines: [28] }], [2]);
    const layout = lay(f, 1);
    expect(layout.pages).toHaveLength(2);
    expect(bodyFragments(layout, 0)[0]).toMatchObject({ fromLine: 0, toLine: 28 });
    expect(bodyFragments(layout, 1)[0]).toMatchObject({ fromLine: 28, toLine: 29 });
    expect(noteFragments(layout, 0)).toHaveLength(0);
    expect(noteFragments(layout, 1).map((fr) => fr.blockId)).toEqual(['note1']);
    expect(layout.pages[0].footnotes).toBeUndefined();
  });

  it('keeps a reference on its page when the note still fits', () => {
    const f = fixture([{ lines: 27, refLines: [26] }], [1]);
    const layout = lay(f, 1);
    expect(layout.pages).toHaveLength(1);
    expect(noteFragments(layout, 0)).toHaveLength(1);
  });

  it('places several notes in the order of their references, and each note once', () => {
    const f = fixture([{ lines: 6, refLines: [1, 1, 4] }, { lines: 3, refLines: [] }], [1, 2, 1]);
    // A second reference to the first note (same key) is not placed twice.
    f.plan.refs.set(1, [{ line: 0, key: 'footnote:1' }]);
    const layout = lay(f, 2);
    expect(noteFragments(layout, 0).map((fr) => fr.blockId)).toEqual(['note1', 'note2', 'note3']);
    expect(noteFragments(layout, 0).map((fr) => fr.y)).toEqual([600 - 80, 600 - 60, 600 - 20]);
    expect(layout.pages[0].footnotes?.top).toBe(600 - 90);
  });

  it('gives every page room for the notes of its own lines, where moving a line changes what a later page needs', () => {
    // Many references spread over several pages: every page ends with room for its own notes and nothing overlaps.
    const body = Array.from({ length: 8 }, () => ({ lines: 12, refLines: [11] }));
    const f = fixture(body, Array(8).fill(3));
    const layout = lay(f, body.length);
    for (const page of layout.pages) {
      const bodyEnd = Math.max(0, ...bodyFragmentsOf(page).map((fr) => fr.y + fr.height));
      const top = page.footnotes?.top ?? CONTENT;
      expect(bodyEnd).toBeLessThanOrEqual(top + 0.5);
    }
    // Every note is somewhere, once.
    const placed = layout.pages.flatMap((p) => p.fragments.filter((fr) => fr.blockId.startsWith('note')).map((fr) => fr.blockId));
    expect(placed.sort()).toEqual(Array.from({ length: 8 }, (_, k) => `note${k + 1}`).sort());
    // And each page's notes are the ones its own lines refer to.
    layout.pages.forEach((page) => {
      const keys = new Set<string>();
      bodyFragmentsOf(page).forEach((fr) => f.plan.refs.get(fr.blockIndex)?.forEach((r) => r.line >= fr.fromLine && r.line < fr.toLine && keys.add(r.key.replace('footnote:', 'note'))));
      const onPage = page.fragments.filter((fr) => fr.blockId.startsWith('note')).map((fr) => fr.blockId);
      expect(onPage.sort()).toEqual([...keys].sort());
    });
  });

  it('continues a long note on the next page, with a full-width separator there, when it is more than a page may give', () => {
    // A 40-line note (800px) does not fit in half a page: 14 lines (the most 300px - 10px allows) go on page 1, the rest on 2.
    const f = fixture([{ lines: 5, refLines: [0] }], [40]);
    const layout = lay(f, 1);
    expect(layout.pages.length).toBeGreaterThanOrEqual(2);
    const first = noteFragments(layout, 0)[0];
    expect(first).toMatchObject({ fromLine: 0, toLine: 14, isFirstFragment: true, isLastFragment: false });
    expect(layout.pages[0].footnotes?.continued).toBe(false);
    const second = noteFragments(layout, 1)[0];
    expect(second).toMatchObject({ fromLine: 14, isFirstFragment: false });
    expect(layout.pages[1].footnotes?.continued).toBe(true);
    const covered = layout.pages.flatMap((p) => p.fragments).filter((fr) => fr.blockId === 'note1').reduce((n, fr) => n + fr.toLine - fr.fromLine, 0);
    expect(covered).toBe(40);
  });

  it('gives notes that continue past the last page pages of their own, set up like the last page', () => {
    const f = fixture([{ lines: 3, refLines: [0] }], [60]);
    const layout = lay(f, 1);
    expect(layout.pages.length).toBeGreaterThanOrEqual(3);
    const last = layout.pages[layout.pages.length - 1];
    expect(last.fragments.every((fr) => fr.blockId === 'note1')).toBe(true);
    expect(last.pageConfig ?? layout.pageConfig).toEqual(PAGE);
    expect(layout.totalHeight).toBe(layout.pages.length * 700 + (layout.pages.length - 1) * 24);
    expect(layout.pages.map((p) => p.pageIndex)).toEqual(layout.pages.map((_, i) => i));
  });

  it('places the first line of a note even when it is taller than the share a page gives to notes', () => {
    const f = fixture([{ lines: 3, refLines: [0] }], [1]);
    f.measures[1] = measureOf('note1', [400]);
    const layout = lay(f, 1);
    expect(noteFragments(layout, 0)).toHaveLength(1);
    expect(noteFragments(layout, 0)[0]).toMatchObject({ fromLine: 0, toLine: 1, height: 400 });
  });

  it('defers a note to the next page when the notes before it already use the page’s share', () => {
    // Two notes of 14 lines each (280px): the first fits in 300px, the second does not and goes to the next page.
    const f = fixture([{ lines: 3, refLines: [0, 1] }], [14, 14]);
    const layout = lay(f, 1);
    expect(noteFragments(layout, 0).map((fr) => fr.blockId)).toEqual(['note1']);
    expect(noteFragments(layout, 1).map((fr) => fr.blockId)).toEqual(['note2']);
    expect(layout.pages[1].footnotes?.continued).toBe(true);
  });

  it('scales the separator and the notes with the zoom', () => {
    const f = fixture([{ lines: 10, refLines: [2] }], [2]);
    const layout = lay(f, 1, 2);
    expect(layout.pages[0].footnotes).toEqual({ top: 1200 - 100, separatorHeight: 20, continued: false });
    expect(noteFragments(layout, 0)[0].height).toBe(80);
  });

  it('works through a layout of several sections, each page getting the notes of its own lines', () => {
    const f = fixture([{ lines: 32, refLines: [31] }, { lines: 5, refLines: [4] }], [2, 2]);
    const sections = [
      { pageConfig: PAGE, blockStart: 0, blockEnd: 1 },
      { pageConfig: PAGE, blockStart: 1, blockEnd: 2 },
    ];
    const hooks = createFootnoteHooks(f.plan, f.measures, 1);
    const layout = applyPageNotes(computeSectionedLayout(f.blocks, f.measures, sections, { scale: 1, pageGap: 24, minLinesAtBreak: 1, footnotes: hooks }), f.plan);
    expect(layout.pages.map((p) => p.sectionIndex)).toEqual([0, 0, 1]);
    expect(noteFragments(layout, 1).map((fr) => fr.blockId)).toEqual(['note1']);
    expect(noteFragments(layout, 2).map((fr) => fr.blockId)).toEqual(['note2']);
  });

  it('shares the page of a continuous section: the notes of both sections are on it, and the room they take counts for both', () => {
    // Section 1: 5 lines with a reference; section 2 continues on the same page with 5 more lines and another reference.
    const f = fixture([{ lines: 5, refLines: [4] }, { lines: 5, refLines: [4] }], [2, 2]);
    const sections = [
      { pageConfig: PAGE, blockStart: 0, blockEnd: 1 },
      { pageConfig: PAGE, blockStart: 1, blockEnd: 2, breakType: 'continuous' as const },
    ];
    const hooks = createFootnoteHooks(f.plan, f.measures, 1);
    const layout = applyPageNotes(computeSectionedLayout(f.blocks, f.measures, sections, { scale: 1, pageGap: 24, minLinesAtBreak: 1, footnotes: hooks }), f.plan);
    expect(layout.pages).toHaveLength(1);
    expect(noteFragments(layout, 0).map((fr) => fr.blockId)).toEqual(['note1', 'note2']);
    expect(layout.pages[0].notes?.keys).toEqual(['footnote:1', 'footnote:2']);
  });

  it('does not share a page that only holds the rest of a note with the next, continuous section', () => {
    const f = fixture([{ lines: 3, refLines: [0] }, { lines: 3 }], [60]);
    const sections = [
      { pageConfig: PAGE, blockStart: 0, blockEnd: 1 },
      { pageConfig: PAGE, blockStart: 1, blockEnd: 2, breakType: 'continuous' as const },
    ];
    const hooks = createFootnoteHooks(f.plan, f.measures, 1);
    const layout = computeSectionedLayout(f.blocks, f.measures, sections, { scale: 1, pageGap: 24, minLinesAtBreak: 1, footnotes: hooks });
    const last = layout.pages[layout.pages.length - 1];
    expect(last.sectionIndex).toBe(1);
    expect(last.notesOnly).toBeUndefined();
    expect(layout.pages.some((p) => p.notesOnly)).toBe(true);
  });
});

function bodyFragmentsOf(page: { fragments: Array<{ blockId: string; y: number; height: number; blockIndex: number; fromLine: number; toLine: number }> }) {
  return page.fragments.filter((fr) => !fr.blockId.startsWith('note'));
}

describe('createFootnoteHooks', () => {
  it('lists the footnotes referenced by a range of lines of a block, once each, in order', () => {
    const f = fixture([{ lines: 6, refLines: [1, 1, 4] }], [1, 1, 1]);
    const hooks = createFootnoteHooks(f.plan, f.measures, 1);
    expect(hooks.refsIn(0, 0, 6)).toEqual(['footnote:1', 'footnote:2', 'footnote:3']);
    expect(hooks.refsIn(0, 0, 1)).toEqual([]);
    expect(hooks.refsIn(0, 1, 2)).toEqual(['footnote:1', 'footnote:2']);
    expect(hooks.refsIn(0, 2, 6)).toEqual(['footnote:3']);
    expect(hooks.refsIn(5, 0, 6)).toEqual([]);
    // A section laid out on its own sees its blocks from 0.
    const shifted = createFootnoteHooks(f.plan, f.measures, 1, 0);
    expect(shifted.refsIn(0, 0, 6)).toHaveLength(3);
    expect(createFootnoteHooks(f.plan, f.measures, 1, 7).refsIn(-7, 0, 6)).toEqual(['footnote:1', 'footnote:2', 'footnote:3']);
  });

  it('places whole notes while they fit in the share of the page, continuing the first one that does not', () => {
    const f = fixture([{ lines: 3, refLines: [0, 0, 0] }], [14, 14, 14]);
    const hooks = createFootnoteHooks(f.plan, f.measures, 1);
    const page = hooks.place(CONTENT, [], ['footnote:1', 'footnote:2', 'footnote:3']);
    // 10px separator + 280px for the first leaves 10px: the second note's first line (20px) does not fit.
    expect(page.items.map((i) => [i.key, i.fromLine, i.toLine])).toEqual([['footnote:1', 0, 14]]);
    expect(page.height).toBe(290);
    expect(page.carry).toEqual([{ key: 'footnote:2', fromLine: 0 }, { key: 'footnote:3', fromLine: 0 }]);
    expect(page.continued).toBe(false);
    const next = hooks.place(CONTENT, page.carry, []);
    expect(next.items.map((i) => i.key)).toEqual(['footnote:2']);
    expect(next.continued).toBe(true);
    expect(next.carry).toEqual([{ key: 'footnote:3', fromLine: 0 }]);
  });

  it('splits a note at a line boundary, and finishes it on the next page', () => {
    const f = fixture([{ lines: 1, refLines: [0] }], [20]);
    const hooks = createFootnoteHooks(f.plan, f.measures, 1);
    const first = hooks.place(CONTENT, [], ['footnote:1']);
    expect(first.items).toEqual([expect.objectContaining({ fromLine: 0, toLine: 14, totalLines: 20, height: 280 })]);
    expect(first.carry).toEqual([{ key: 'footnote:1', fromLine: 14 }]);
    const rest = hooks.place(CONTENT, first.carry, []);
    expect(rest.items).toEqual([expect.objectContaining({ fromLine: 14, toLine: 20, height: 120 })]);
    expect(rest.carry).toEqual([]);
    expect(rest.continued).toBe(true);
    expect(rest.height).toBe(130);
  });

  it('places nothing, and carries nothing, for a page with no notes; ignores a note without a block and one already carried', () => {
    const f = fixture([{ lines: 1, refLines: [0] }], [1]);
    const hooks = createFootnoteHooks(f.plan, f.measures, 1);
    expect(hooks.place(CONTENT, [], [])).toEqual({ items: [], height: 0, continued: false, carry: [] });
    expect(hooks.place(CONTENT, [], ['footnote:42'])).toEqual({ items: [], height: 0, continued: false, carry: [] });
    const both = hooks.place(CONTENT, [{ key: 'footnote:1', fromLine: 0 }], ['footnote:1']);
    expect(both.items).toHaveLength(1);
  });

  it('puts the first line of a note on the page even when it is taller than the share, and the whole note when it is one line', () => {
    const f = fixture([{ lines: 1, refLines: [0] }], [1]);
    f.measures[1] = measureOf('note1', [400]);
    const hooks = createFootnoteHooks(f.plan, f.measures, 1);
    expect(hooks.place(CONTENT, [], ['footnote:1']).items).toEqual([expect.objectContaining({ fromLine: 0, toLine: 1, height: 400 })]);
    f.measures[1] = measureOf('note1', [400, 400]);
    const tall = hooks.place(CONTENT, [], ['footnote:1']);
    expect(tall.items).toEqual([expect.objectContaining({ fromLine: 0, toLine: 1, height: 400, totalLines: 2 })]);
    expect(tall.carry).toEqual([{ key: 'footnote:1', fromLine: 1 }]);
  });

  it('treats a note whose block has no lines as one line', () => {
    const f = fixture([{ lines: 1, refLines: [0] }], [1]);
    f.measures[1] = { kind: 'image', blockId: 'note1', totalHeight: 30, width: 10 } as Measure;
    const placed = createFootnoteHooks(f.plan, f.measures, 1).place(CONTENT, [], ['footnote:1']);
    expect(placed.items).toEqual([expect.objectContaining({ fromLine: 0, toLine: 1, totalLines: 1, height: 30 })]);
  });
});

describe('applyPageNotes', () => {
  it('draws a copy of the layout with the notes of every page, and leaves pages without notes as they are', () => {
    const f = fixture([{ lines: 10, refLines: [2] }, { lines: 40 }], [2]);
    const hooks = createFootnoteHooks(f.plan, f.measures, 1);
    const base = computeTrueLayout(f.blocks.slice(0, 2), f.measures.slice(0, 2), { pageConfig: PAGE, scale: 1, pageGap: 24, minLinesAtBreak: 1, footnotes: hooks });
    const applied = applyPageNotes(base, f.plan);
    expect(applied).not.toBe(base);
    expect(base.pages[0].fragments.some((fr) => fr.blockId === 'note1')).toBe(false);
    expect(applied.pages[0].fragments.some((fr) => fr.blockId === 'note1')).toBe(true);
    expect(applied.pages[1]).toBe(base.pages[1]);
    expect(applied.pages[1].footnotes).toBeUndefined();
  });
});
