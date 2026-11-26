import { computeTrueLayout, findPageAtY, getPageConfig, getPageY, type PageConfig } from './true-layout-engine';
import {
  canContinueOnPage,
  computeSectionedLayout,
  formatPageNumber,
  pickHeaderFooterVariant,
  resolveInheritedContent,
  type LayoutSection,
} from './section-layout';
import type { FlowBlock } from './flow-blocks';
import type { LineMeasure, Measure, ParagraphMeasure } from './measurer';

const portrait: PageConfig = { width: 800, height: 700, margins: { top: 50, right: 50, bottom: 50, left: 50 } }; // content 600
const landscape: PageConfig = { width: 1000, height: 500, margins: { top: 50, right: 80, bottom: 50, left: 80 } }; // content 400
const a4: PageConfig = { width: 794, height: 1123, margins: { top: 96, right: 72, bottom: 96, left: 72 } };

const lineOf = (height: number): LineMeasure => ({ height, width: 100, ascent: height * 0.8, descent: height * 0.2, yOffset: 0 });
const para = (id: string, lineHeight: number, n: number): ParagraphMeasure => ({
  kind: 'paragraph',
  blockId: id,
  lines: Array.from({ length: n }, () => lineOf(lineHeight)),
  totalHeight: lineHeight * n,
  spaceBefore: 0,
  spaceAfter: 0,
});
const block = (id: string): FlowBlock => ({ id, type: 'paragraph', content: [] } as unknown as FlowBlock);

/** Blocks and measures for sections described as lists of block heights (one 10px-line per 10px of height). */
function build(sizes: number[][]): { blocks: FlowBlock[]; measures: Measure[]; ranges: Array<[number, number]> } {
  const blocks: FlowBlock[] = [];
  const measures: Measure[] = [];
  const ranges: Array<[number, number]> = [];
  let n = 0;
  for (const heights of sizes) {
    const start = blocks.length;
    for (const h of heights) {
      const id = `b${n++}`;
      blocks.push(block(id));
      measures.push(para(id, 10, h / 10));
    }
    ranges.push([start, blocks.length]);
  }
  return { blocks, measures, ranges };
}

function sectionsOf(ranges: Array<[number, number]>, configs: Array<Partial<LayoutSection> & { pageConfig: PageConfig }>): LayoutSection[] {
  return ranges.map(([blockStart, blockEnd], i) => ({ blockStart, blockEnd, ...configs[i] }));
}

describe('computeSectionedLayout — pages of different sections', () => {
  it('gives the same pages as computeTrueLayout for one section', () => {
    const { blocks, measures, ranges } = build([[200, 300, 250]]);
    const sectioned = computeSectionedLayout(blocks, measures, sectionsOf(ranges, [{ pageConfig: portrait }]), { scale: 1, pageGap: 24 });
    const plain = computeTrueLayout(blocks, measures, { pageConfig: portrait, scale: 1, pageGap: 24 });
    expect(sectioned.pages.map((p) => p.fragments)).toEqual(plain.pages.map((p) => p.fragments));
    expect(sectioned.totalHeight).toBe(plain.totalHeight);
    expect(sectioned.pages.map((p) => p.pageNumber)).toEqual([1, 2]);
  });

  it('starts each section on a new page with its own page setup, size and position', () => {
    const { blocks, measures, ranges } = build([[100], [100], [100]]);
    const layout = computeSectionedLayout(blocks, measures, sectionsOf(ranges, [{ pageConfig: portrait }, { pageConfig: landscape }, { pageConfig: a4 }]), {
      scale: 1,
      pageGap: 24,
    });
    expect(layout.pages.map((p) => p.sectionIndex)).toEqual([0, 1, 2]);
    expect(layout.pages.map((p) => getPageConfig(layout, p.pageIndex))).toEqual([portrait, landscape, a4]);
    expect(layout.pages.map((p) => p.top)).toEqual([0, 724, 1248]);
    expect(layout.totalHeight).toBe(700 + 24 + 500 + 24 + 1123);
    expect(layout.pageConfig).toBe(portrait);
    expect(layout.pages[1].fragments[0]).toMatchObject({ blockId: 'b1', blockIndex: 1, y: 0 });
  });

  it('paginates a long section by its own content height', () => {
    // 750px of text: two pages of 600, but three of 400 (paragraphs split at line boundaries).
    const { blocks, measures, ranges } = build([[150, 150, 150, 150, 150], [150, 150, 150, 150, 150]]);
    const layout = computeSectionedLayout(blocks, measures, sectionsOf(ranges, [{ pageConfig: portrait }, { pageConfig: landscape }]));
    expect(layout.pages.map((p) => [p.sectionIndex, p.contentHeight])).toEqual([[0, 600], [0, 150], [1, 400], [1, 350]]);
  });

  it('scales page positions with the zoom', () => {
    const { blocks, measures, ranges } = build([[100], [100]]);
    const layout = computeSectionedLayout(blocks, measures, sectionsOf(ranges, [{ pageConfig: portrait }, { pageConfig: landscape }]), { scale: 0.5, pageGap: 24 });
    expect(layout.pages.map((p) => p.top)).toEqual([0, 350 + 12]);
    expect(layout.totalHeight).toBe(350 + 12 + 250);
  });

  it('gives an empty section its own empty page and keeps the order', () => {
    const { blocks, measures, ranges } = build([[100], [], [100]]);
    const layout = computeSectionedLayout(blocks, measures, sectionsOf(ranges, [{ pageConfig: portrait }, { pageConfig: portrait }, { pageConfig: portrait }]));
    expect(layout.pages.map((p) => [p.sectionIndex, p.fragments.length])).toEqual([[0, 1], [1, 0], [2, 1]]);
  });

  it('rejects an empty list of sections', () => {
    expect(() => computeSectionedLayout([], [], [])).toThrow('at least one section');
  });
});

describe('computeSectionedLayout — continuous breaks', () => {
  it('continues on the same page below the previous section', () => {
    const { blocks, measures, ranges } = build([[200], [100, 100]]);
    const layout = computeSectionedLayout(blocks, measures, sectionsOf(ranges, [{ pageConfig: portrait }, { pageConfig: portrait, breakType: 'continuous' }]));
    expect(layout.pages).toHaveLength(1);
    expect(layout.pages[0].fragments.map((f) => [f.blockIndex, f.y])).toEqual([[0, 0], [1, 200], [2, 300]]);
    expect(layout.pages[0].contentHeight).toBe(400);
    expect(layout.pages[0].remainingHeight).toBe(200);
    expect(layout.pages[0].sectionIndex).toBe(0);
  });

  it('flows onto the next page when it does not fit, which belongs to the new section', () => {
    const { blocks, measures, ranges } = build([[500], [100, 100]]);
    const layout = computeSectionedLayout(blocks, measures, sectionsOf(ranges, [{ pageConfig: portrait }, { pageConfig: portrait, breakType: 'continuous' }]));
    expect(layout.pages.map((p) => [p.sectionIndex, p.fragments.map((f) => f.blockIndex)])).toEqual([[0, [0, 1]], [1, [2]]]);
    expect(layout.pages[1].fragments[0].y).toBe(0);
  });

  it('starts a new page when the page setup differs', () => {
    const { blocks, measures, ranges } = build([[200], [100]]);
    const layout = computeSectionedLayout(blocks, measures, sectionsOf(ranges, [{ pageConfig: portrait }, { pageConfig: landscape, breakType: 'continuous' }]));
    expect(layout.pages.map((p) => p.sectionIndex)).toEqual([0, 1]);
    expect(canContinueOnPage(portrait, landscape)).toBe(false);
    expect(canContinueOnPage(portrait, { ...portrait, margins: { ...portrait.margins, left: 60 } })).toBe(false);
    expect(canContinueOnPage(portrait, { ...portrait })).toBe(true);
  });

  it('honours a page break at the start of a continuous section and an empty continuous section', () => {
    const blocks: FlowBlock[] = [block('a'), block('pb'), block('c'), block('d')];
    const measures: Measure[] = [para('a', 10, 5), { kind: 'pageBreak', blockId: 'pb', totalHeight: 0, forceBreak: true }, para('c', 10, 5), para('d', 10, 5)];
    const layout = computeSectionedLayout(
      blocks,
      measures,
      [
        { pageConfig: portrait, blockStart: 0, blockEnd: 1 },
        { pageConfig: portrait, blockStart: 1, blockEnd: 3, breakType: 'continuous' },
        { pageConfig: portrait, blockStart: 3, blockEnd: 3, breakType: 'continuous' },
        { pageConfig: portrait, blockStart: 3, blockEnd: 4, breakType: 'continuous' },
      ],
    );
    expect(layout.pages.map((p) => p.fragments.map((f) => f.blockIndex))).toEqual([[0], [2, 3]]);
  });
});

describe('computeSectionedLayout — even and odd page starts', () => {
  it('adds a blank page so an odd-page section starts on an odd page', () => {
    const { blocks, measures, ranges } = build([[100], [100], [100]]);
    const layout = computeSectionedLayout(blocks, measures, sectionsOf(ranges, [{ pageConfig: portrait }, { pageConfig: portrait, breakType: 'oddPage' }, { pageConfig: portrait, breakType: 'evenPage' }]));
    // page 1 (section 0), page 2 blank, page 3 (section 1, odd), page 4 (section 2, even; no blank needed)
    expect(layout.pages.map((p) => [p.pageNumber, p.isBlank === true, p.sectionIndex])).toEqual([
      [1, false, 0],
      [2, true, 0],
      [3, false, 1],
      [4, false, 2],
    ]);
    expect(layout.pages[1].fragments).toEqual([]);
  });

  it('adds a blank page of the previous setup before an even-page section', () => {
    const { blocks, measures, ranges } = build([[100], [100]]);
    const layout = computeSectionedLayout(blocks, measures, sectionsOf(ranges, [{ pageConfig: landscape }, { pageConfig: portrait, breakType: 'oddPage' }]));
    // section 1 wants an odd page and would be page 2: a blank landscape page goes first.
    expect(layout.pages.map((p) => [p.pageNumber, p.isBlank === true])).toEqual([[1, false], [2, true], [3, false]]);
    expect(layout.pages[1].pageConfig).toBe(landscape);
    expect(layout.pages[1].remainingHeight).toBe(400);
  });

  it('does not add a filler page when the section restarts its numbering', () => {
    const { blocks, measures, ranges } = build([[100], [100]]);
    const layout = computeSectionedLayout(blocks, measures, sectionsOf(ranges, [{ pageConfig: portrait }, { pageConfig: portrait, breakType: 'evenPage', pageNumbering: { start: 1 } }]));
    expect(layout.pages.map((p) => [p.pageNumber, p.isBlank === true])).toEqual([[1, false], [1, false]]);
  });
});

describe('computeSectionedLayout — page numbers and header/footer variants', () => {
  it('restarts and continues numbering per section and carries the format', () => {
    const { blocks, measures, ranges } = build([[500, 500], [100], [100, 600]]);
    const layout = computeSectionedLayout(
      blocks,
      measures,
      sectionsOf(ranges, [
        { pageConfig: portrait },
        { pageConfig: portrait, pageNumbering: { start: 10, format: 'lowerRoman' } },
        { pageConfig: portrait, pageNumbering: { format: 'upperLetter' } },
      ]),
    );
    expect(layout.pages.map((p) => [p.sectionIndex, p.pageNumber, p.pageNumberFormat])).toEqual([
      [0, 1, undefined],
      [0, 2, undefined],
      [1, 10, 'lowerRoman'],
      [2, 11, 'upperLetter'],
      [2, 12, 'upperLetter'],
    ]);
  });

  it('picks first, even and default variants per section', () => {
    const { blocks, measures, ranges } = build([[500, 500, 500], [500, 500, 500]]);
    const layout = computeSectionedLayout(
      blocks,
      measures,
      sectionsOf(ranges, [
        { pageConfig: portrait, differentFirstPage: true },
        { pageConfig: portrait, differentFirstPage: true, differentOddEven: true },
      ]),
    );
    expect(layout.pages.map((p) => p.headerFooterVariant)).toEqual(['first', 'default', 'default', 'first', 'default', 'even']);
  });

  it('a continuous section does not make its first page a first page', () => {
    const { blocks, measures, ranges } = build([[200], [100, 600]]);
    const layout = computeSectionedLayout(blocks, measures, sectionsOf(ranges, [{ pageConfig: portrait }, { pageConfig: portrait, breakType: 'continuous', differentFirstPage: true }]));
    expect(layout.pages.map((p) => p.headerFooterVariant)).toEqual(['default', 'default']);
  });

  it('pickHeaderFooterVariant uses the page number, not the position', () => {
    const s = { differentFirstPage: false, differentOddEven: true };
    expect(pickHeaderFooterVariant(s, false, 7)).toBe('default');
    expect(pickHeaderFooterVariant(s, false, 8)).toBe('even');
    expect(pickHeaderFooterVariant({ differentFirstPage: true }, true, 1)).toBe('first');
    expect(pickHeaderFooterVariant({}, true, 2)).toBe('default');
  });
});

describe('formatPageNumber', () => {
  it('formats roman, letter and decimal numbers', () => {
    expect(formatPageNumber(4, 'lowerRoman')).toBe('iv');
    expect(formatPageNumber(1994, 'upperRoman')).toBe('MCMXCIV');
    expect(formatPageNumber(4000, 'upperRoman')).toBe('4000');
    expect(formatPageNumber(0, 'lowerRoman')).toBe('0');
    expect(formatPageNumber(1, 'upperLetter')).toBe('A');
    expect(formatPageNumber(26, 'lowerLetter')).toBe('z');
    expect(formatPageNumber(27, 'upperLetter')).toBe('AA');
    expect(formatPageNumber(0, 'upperLetter')).toBe('0');
    expect(formatPageNumber(12)).toBe('12');
    expect(formatPageNumber(12, 'decimalZero')).toBe('12');
  });
});

describe('resolveInheritedContent', () => {
  it('lets a section with none show the previous one’s', () => {
    expect(resolveInheritedContent(['a', undefined, 'b', undefined, undefined])).toEqual(['a', 'a', 'b', 'b', 'b']);
    expect(resolveInheritedContent([undefined, 'x'])).toEqual([undefined, 'x']);
    expect(resolveInheritedContent([])).toEqual([]);
  });
});

describe('page lookups on a layout with pages of different heights', () => {
  it('finds the page tops and the page at a y', () => {
    const { blocks, measures, ranges } = build([[100], [100], [100]]);
    const layout = computeSectionedLayout(blocks, measures, sectionsOf(ranges, [{ pageConfig: portrait }, { pageConfig: landscape }, { pageConfig: a4 }]));
    expect([0, 1, 2].map((i) => getPageY(layout, i))).toEqual([0, 724, 1248]);
    expect([0, 699, 700, 723, 724, 1000, 1247, 1248, 9999, -5].map((y) => findPageAtY(layout, y))).toEqual([0, 0, 0, 0, 1, 1, 1, 2, 2, 0]);
    expect(getPageConfig(layout, 2)).toBe(a4);
  });

  it('keeps the uniform lookups for a layout without sections', () => {
    const { blocks, measures } = build([[500, 500]]);
    const layout = computeTrueLayout(blocks, measures, { pageConfig: portrait, scale: 1, pageGap: 24 });
    expect(getPageY(layout, 1)).toBe(724);
    expect(findPageAtY(layout, 800)).toBe(1);
    expect(getPageConfig(layout, 0)).toBe(portrait);
    expect(getPageConfig(layout, 99)).toBe(portrait);
  });
});
