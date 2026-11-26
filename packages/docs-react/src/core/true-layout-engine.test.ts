import {
  computeTrueLayout,
  findBlockPage,
  findBlockFragments,
  getPageY,
  findPageAtY,
  type FootnoteHooks,
  type LayoutOptions,
} from './true-layout-engine';
import type { FlowBlock } from './flow-blocks';
import type { Measure, LineMeasure, ParagraphMeasure, PageBreakMeasure, ImageMeasure } from './measurer';

// --- fixtures ---------------------------------------------------------------

// A US-Letter-ish page: content area is 600px tall (height 700, 50px margins).
const CONTENT_HEIGHT = 600;
const opts = (over: Partial<LayoutOptions> = {}): LayoutOptions => ({
  pageConfig: { width: 800, height: 700, margins: { top: 50, right: 50, bottom: 50, left: 50 } },
  scale: 1,
  pageGap: 24,
  minLinesAtBreak: 2,
  ...over,
});

const block = (id: string): FlowBlock => ({ id, type: 'paragraph', content: [] } as unknown as FlowBlock);

const lineOf = (height: number): LineMeasure => ({ height, width: 100, ascent: height * 0.8, descent: height * 0.2, yOffset: 0 });

// Paragraph with `n` lines each `lineHeight` tall (no space before/after by default).
const para = (id: string, lineHeight: number, n: number, spaceBefore = 0, spaceAfter = 0): ParagraphMeasure => ({
  kind: 'paragraph',
  blockId: id,
  lines: Array.from({ length: n }, () => lineOf(lineHeight)),
  totalHeight: spaceBefore + lineHeight * n + spaceAfter,
  spaceBefore,
  spaceAfter,
});

const pageBreak = (id: string): PageBreakMeasure => ({ kind: 'pageBreak', blockId: id, totalHeight: 0, forceBreak: true });
const image = (id: string, height: number): ImageMeasure => ({ kind: 'image', blockId: id, totalHeight: height, width: 200 });

function layout(pairs: Array<{ b: FlowBlock; m: Measure }>, over?: Partial<LayoutOptions>) {
  return computeTrueLayout(pairs.map((p) => p.b), pairs.map((p) => p.m), opts(over));
}

// --- tests ------------------------------------------------------------------

describe('computeTrueLayout — single page', () => {
  it('places a block that fits as one full fragment', () => {
    const l = layout([{ b: block('a'), m: para('a', 20, 3) }]);
    expect(l.pages).toHaveLength(1);
    expect(l.pages[0].fragments).toHaveLength(1);
    expect(l.pages[0].fragments[0]).toMatchObject({ blockId: 'a', fromLine: 0, toLine: 3, y: 0, isFirstFragment: true, isLastFragment: true });
    expect(l.pages[0].contentHeight).toBe(60);
    expect(l.pages[0].remainingHeight).toBe(CONTENT_HEIGHT - 60);
  });

  it('stacks multiple blocks on one page with increasing y', () => {
    const l = layout([
      { b: block('a'), m: para('a', 20, 2) },
      { b: block('b'), m: para('b', 20, 2) },
    ]);
    expect(l.pages).toHaveLength(1);
    expect(l.pages[0].fragments.map((f) => f.y)).toEqual([0, 40]);
  });

  it('returns a single empty page for empty input', () => {
    const l = computeTrueLayout([], [], opts());
    expect(l.pages).toHaveLength(1);
    expect(l.pages[0].fragments).toHaveLength(0);
  });
});

describe('computeTrueLayout — pagination', () => {
  it('moves a whole block that does not fit to the next page', () => {
    // First block fills 580 of 600; a 3-line (60px) unsplittable-sized block...
    // is splittable, so use a single-line block that can't split.
    const l = layout([
      { b: block('a'), m: para('a', 580, 1) },
      { b: block('b'), m: para('b', 40, 1) },
    ]);
    expect(l.pages).toHaveLength(2);
    expect(l.pages[0].fragments.map((f) => f.blockId)).toEqual(['a']);
    expect(l.pages[1].fragments.map((f) => f.blockId)).toEqual(['b']);
  });

  it('splits a multi-line block across two pages at a line boundary', () => {
    // 40 lines × 20px = 800px total, content height 600 → ~30 lines fit, rest overflow.
    const l = layout([{ b: block('a'), m: para('a', 20, 40) }]);
    expect(l.pages.length).toBeGreaterThanOrEqual(2);
    const frags = l.pages.flatMap((p) => p.fragments).filter((f) => f.blockId === 'a');
    // Fragments must tile [0..40] with no gap or overlap and no lost lines.
    expect(frags[0].fromLine).toBe(0);
    for (let i = 1; i < frags.length; i++) expect(frags[i].fromLine).toBe(frags[i - 1].toLine);
    expect(frags[frags.length - 1].toLine).toBe(40);
    expect(frags[0].isFirstFragment).toBe(true);
    expect(frags[frags.length - 1].isLastFragment).toBe(true);
  });

  it('forces a new page on a pageBreak measure', () => {
    const l = layout([
      { b: block('a'), m: para('a', 20, 2) },
      { b: block('pb'), m: pageBreak('pb') },
      { b: block('b'), m: para('b', 20, 2) },
    ]);
    expect(l.pages).toHaveLength(2);
    expect(l.pages[0].fragments.map((f) => f.blockId)).toEqual(['a']);
    expect(l.pages[1].fragments.map((f) => f.blockId)).toEqual(['b']);
  });

  it('applies widow control: a lone trailing line is pushed to the next page', () => {
    // Page fits 30 lines (600/20). A 31-line block would leave 1 line alone on
    // page 2 — with minLinesAtBreak=2 the break moves up so ≥2 lines carry over.
    const l = layout([{ b: block('a'), m: para('a', 20, 31) }]);
    const lastPageFrag = l.pages[l.pages.length - 1].fragments.find((f) => f.blockId === 'a')!;
    expect(lastPageFrag.toLine - lastPageFrag.fromLine).toBeGreaterThanOrEqual(2);
  });
});

describe('computeTrueLayout — oversized content never vanishes', () => {
  it('lets an unsplittable block taller than the page overflow onto its own page', () => {
    const l = layout([{ b: block('img'), m: image('img', 900) }]);
    const frag = l.pages.flatMap((p) => p.fragments).find((f) => f.blockId === 'img');
    expect(frag).toBeDefined();
    // Regression: the fragment height was clamped to the page content height
    // (600), silently truncating the bottom 300px of the block. It must cover
    // the block's full measured height so no content is lost.
    expect(frag!.height).toBe(900);
  });

  it('preserves the full height of an unsplittable block preceded by content', () => {
    // A block that fills part of page 1, then a tall non-line block that cannot
    // split: the tall block moves to its own page and keeps its whole height.
    const l = layout([
      { b: block('a'), m: para('a', 20, 5) },
      { b: block('img'), m: image('img', 1000) },
    ]);
    const frag = l.pages.flatMap((p) => p.fragments).find((f) => f.blockId === 'img')!;
    expect(frag.height).toBe(1000);
  });

  it('keeps every line of a multi-line block whose first line is taller than the page', () => {
    // Regression: a line taller than a full page used to make the widow guard
    // return "0 lines fit", and the fragment loop then broke, silently dropping
    // the block. Every line must still be emitted across pages.
    const l = layout([{ b: block('a'), m: para('a', 700, 2) }]);
    const frags = l.pages.flatMap((p) => p.fragments).filter((f) => f.blockId === 'a');
    const linesCovered = frags.reduce((n, f) => n + (f.toLine - f.fromLine), 0);
    expect(linesCovered).toBe(2);
    expect(frags[0].fromLine).toBe(0);
    expect(frags[frags.length - 1].toLine).toBe(2);
  });
});

describe('layout query helpers', () => {
  const l = layout([
    { b: block('a'), m: para('a', 20, 40) }, // spans pages
    { b: block('b'), m: para('b', 20, 2) },
  ]);

  it('findBlockPage returns the first page a block appears on', () => {
    expect(findBlockPage(l, 'a')).toBe(0);
    expect(findBlockPage(l, 'missing')).toBeNull();
  });

  it('findBlockFragments returns every fragment of a block in order', () => {
    const frags = findBlockFragments(l, 'a');
    expect(frags.length).toBeGreaterThanOrEqual(2);
    expect(frags[0].fromLine).toBe(0);
  });

  it('getPageY increases by page height + gap per page', () => {
    expect(getPageY(l, 0)).toBe(0);
    expect(getPageY(l, 1)).toBe(700 + 24);
  });

  it('findPageAtY maps a y coordinate back to its page index', () => {
    expect(findPageAtY(l, 10)).toBe(0);
    expect(findPageAtY(l, 700 + 24 + 10)).toBe(1);
  });
});

describe('computeTrueLayout — footnotes', () => {
  // Hooks for a body of blocks where block `a` has a reference to note `n1` on line 0, and so on. A page's notes take
  // 10px of separator and 30px for each note (so a page with one note gives up 40px).
  const hooks = (refs: Record<string, string[]>, indexOf: (i: number) => string, overrides: Partial<FootnoteHooks> = {}): FootnoteHooks => ({
    refsIn: (i, from, to) => (from <= 0 && to > 0 ? refs[indexOf(i)] ?? [] : []),
    place: (_h, carry, keys) => {
      const all = [...carry.map((c) => c.key), ...keys];
      const items = all.map((key) => ({ key, blockIndex: 99, blockId: key, fromLine: 0, toLine: 1, totalLines: 1, height: 30 }));
      return { items, height: items.length > 0 ? 10 + 30 * items.length : 0, continued: carry.length > 0, carry: [] };
    },
    ...overrides,
  });

  it('is the same layout as without footnotes when no line refers to one', () => {
    const pairs = [{ b: block('a'), m: para('a', 20, 40) }, { b: block('b'), m: image('b', 100) }];
    const plain = layout(pairs);
    const withHooks = layout(pairs, { footnotes: hooks({}, () => '') });
    expect(withHooks.pages.map((p) => p.fragments)).toEqual(plain.pages.map((p) => p.fragments));
    expect(withHooks.pages.every((p) => p.notes === undefined)).toBe(true);
  });

  it('counts the notes a block brings against the space the page has for it', () => {
    // Three 190px blocks fit a 600px page (570px); a note of 40px for the third leaves 560px: it goes to the next page.
    const ids = ['a', 'b', 'c'];
    const l = layout(ids.map((id) => ({ b: block(id), m: image(id, 190) })), { footnotes: hooks({ c: ['n1'] }, (i) => ids[i]) });
    expect(l.pages.map((p) => p.fragments.map((f) => f.blockId))).toEqual([['a', 'b'], ['c']]);
    expect(l.pages[0].notes).toBeUndefined();
    expect(l.pages[1].notes).toMatchObject({ keys: ['n1'], height: 40 });
    expect(l.pages[1].remainingHeight).toBe(CONTENT_HEIGHT - 190 - 40);
  });

  it('fits a block with its notes when there is room, and keeps what the page carries', () => {
    const ids = ['a', 'b'];
    const l = layout(ids.map((id) => ({ b: block(id), m: image(id, 200) })), { footnotes: hooks({ a: ['n1'], b: ['n2'] }, (i) => ids[i]) });
    expect(l.pages).toHaveLength(1);
    expect(l.pages[0].notes).toMatchObject({ keys: ['n1', 'n2'], height: 70, carryIn: [], carry: [] });
    expect(l.pages[0].remainingHeight).toBe(CONTENT_HEIGHT - 400 - 70);
  });

  it('places the lines of a split block only while they fit with their notes', () => {
    // 40 lines of 20px; the reference is on line 27 and the note area takes 40px: 27 lines (540px) + the line (20px) + 40px = 600px fits.
    const refs = (i: number, from: number, to: number) => (from <= 27 && to > 27 ? ['n1'] : []);
    const fits = layout([{ b: block('a'), m: para('a', 20, 40) }], { footnotes: hooks({}, () => '', { refsIn: refs }), minLinesAtBreak: 1 });
    expect(fits.pages[0].fragments[0]).toMatchObject({ fromLine: 0, toLine: 28 });
    // With a 41px note area (30px per note + 11 separator) the line with the reference no longer fits.
    const tight = layout([{ b: block('a'), m: para('a', 20, 40) }], {
      footnotes: hooks({}, () => '', { refsIn: refs, place: (_h, _c, keys) => ({ items: keys.length ? [{ key: 'n1', blockIndex: 9, blockId: 'n1', fromLine: 0, toLine: 1, totalLines: 1, height: 31 }] : [], height: keys.length ? 41 : 0, continued: false, carry: [] }) }),
      minLinesAtBreak: 1,
    });
    expect(tight.pages[0].fragments[0]).toMatchObject({ fromLine: 0, toLine: 27 });
    expect(tight.pages[1].fragments[0]).toMatchObject({ fromLine: 27 });
    expect(tight.pages[0].notes).toBeUndefined();
    expect(tight.pages[1].notes).toMatchObject({ keys: ['n1'] });
  });

  it('starts the next page with the notes that did not fit, and gives left-over notes pages of their own', () => {
    let calls = 0;
    const carrying = hooks({ a: ['n1'] }, () => 'a', {
      place: (_h, carry, keys) => {
        calls++;
        // The first note needs two pages: it places 30px and carries the rest, then a second piece, then nothing.
        const piece = carry.length > 0 ? carry[0].fromLine : 0;
        const items = carry.length + keys.length > 0 ? [{ key: 'n1', blockIndex: 9, blockId: 'n1', fromLine: piece, toLine: piece + 1, totalLines: 3, height: 30 }] : [];
        return { items, height: items.length ? 40 : 0, continued: carry.length > 0, carry: piece + 1 < 3 && items.length ? [{ key: 'n1', fromLine: piece + 1 }] : [] };
      },
    });
    const l = layout([{ b: block('a'), m: image('a', 100) }], { footnotes: carrying });
    expect(l.pages.map((p) => p.notes?.items.map((i) => [i.fromLine, i.toLine]))).toEqual([[[0, 1]], [[1, 2]], [[2, 3]]]);
    expect(l.pages.map((p) => p.notesOnly ?? false)).toEqual([false, true, true]);
    expect(l.pages.map((p) => p.notes?.continued)).toEqual([false, true, true]);
    expect(l.pages.map((p) => p.pageIndex)).toEqual([0, 1, 2]);
    expect(l.totalHeight).toBe(3 * 700 + 2 * 24);
    expect(calls).toBeGreaterThan(0);
  });

  it('accounts for notes carried onto a page when placing its first lines, and lets a continuing page start with notes already on it', () => {
    const carryHooks = hooks({}, () => '', { refsIn: () => [] });
    const l = layout([{ b: block('a'), m: image('a', 570) }], { footnotes: carryHooks, initialNotes: { keys: ['n0'], carry: [{ key: 'n0', fromLine: 1 }] }, initialY: 10 });
    // The page already carries a 40px area (10 + 30): 570px more does not fit in the 590px left, the block goes to page 2.
    expect(l.pages.map((p) => p.fragments.length)).toEqual([0, 1]);
    expect(l.pages[0].notes).toMatchObject({ keys: ['n0'], carryIn: [{ key: 'n0', fromLine: 1 }] });
  });

  it('forces the first line of a page even when its notes do not fit with it, and keeps oversize blocks and their notes', () => {
    const l = layout([{ b: block('a'), m: para('a', 700, 2) }], { footnotes: hooks({ a: ['n1'] }, () => 'a'), minLinesAtBreak: 1 });
    const covered = l.pages.flatMap((p) => p.fragments).reduce((n, f) => n + f.toLine - f.fromLine, 0);
    expect(covered).toBe(2);
    const big = layout([{ b: block('img'), m: image('img', 900) }], { footnotes: hooks({ img: ['n1'] }, () => 'img') });
    expect(big.pages[0].fragments[0].height).toBe(900);
    expect(big.pages[0].notes?.keys).toEqual(['n1']);
  });

  it('puts a page break between pages and carries notes across it', () => {
    const ids = ['a', 'pb', 'b'];
    const l = layout([{ b: block('a'), m: image('a', 50) }, { b: block('pb'), m: pageBreak('pb') }, { b: block('b'), m: image('b', 50) }], { footnotes: hooks({ a: ['n1'], b: ['n2'] }, (i) => ids[i]) });
    expect(l.pages.map((p) => p.notes?.keys)).toEqual([['n1'], ['n2']]);
  });
});
