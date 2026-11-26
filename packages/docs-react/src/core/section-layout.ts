/**
 * Section layout - paginates a document that has several sections
 *
 * Each section brings its own page setup (size, orientation, margins), page numbering, break type and header/footer
 * switches. This module lays the sections out one after another with `computeTrueLayout` and stitches the pages into one
 * `DocumentLayout` whose pages know their own setup, position, number and header/footer variant. Columns are not laid
 * out (one column), and a document with a single section gives the same pages as `computeTrueLayout`.
 */

import type { FlowBlock } from './flow-blocks';
import type { Measure } from './measurer';
import {
  computeTrueLayout,
  type DocumentLayout,
  type HeaderFooterVariant,
  type PageConfig,
  type PageLayout,
} from './true-layout-engine';

/** How a section starts relative to the previous one. */
export type LayoutSectionBreak = 'nextPage' | 'continuous' | 'evenPage' | 'oddPage';

/** One section to lay out: its page setup and the slice of the document's blocks it owns. */
export interface LayoutSection {
  /** Page setup as rendered (width and height already swapped for landscape). */
  pageConfig: PageConfig;
  /** First block (index into the blocks and measures arrays) of the section. */
  blockStart: number;
  /** One past the last block of the section. */
  blockEnd: number;
  /** How the section starts; the first section always starts on the first page. Default: a new page. */
  breakType?: LayoutSectionBreak;
  /** Page numbering: `start` restarts the count at that number, otherwise it continues from the previous page. */
  pageNumbering?: { start?: number; format?: string };
  /** The first page of the section has its own header/footer. */
  differentFirstPage?: boolean;
  /** Even and odd pages have different headers/footers. */
  differentOddEven?: boolean;
}

export interface SectionedLayoutOptions {
  /** Scale factor (zoom level) */
  scale?: number;
  /** Gap between pages in pixels */
  pageGap?: number;
  /** Minimum number of lines to keep together at page break (widow/orphan control) */
  minLinesAtBreak?: number;
}

/** True when a continuous section can share the page of the previous one: same size and same margins. */
export function canContinueOnPage(previous: PageConfig, next: PageConfig): boolean {
  const a = previous.margins;
  const b = next.margins;
  return (
    previous.width === next.width &&
    previous.height === next.height &&
    a.top === b.top &&
    a.bottom === b.bottom &&
    a.left === b.left &&
    a.right === b.right
  );
}

/** Roman numerals for 1 to 3999; anything outside falls back to decimal. */
function toRoman(n: number): string {
  if (n < 1 || n > 3999) return String(n);
  const table: Array<[number, string]> = [
    [1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'],
    [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I'],
  ];
  let rest = n;
  let out = '';
  for (const [value, glyph] of table) {
    while (rest >= value) {
      out += glyph;
      rest -= value;
    }
  }
  return out;
}

/** Letters for 1 to 26 repeat as A..Z, AA..ZZ, AAA..: the sequence Word uses for `upperLetter` and `lowerLetter`. */
function toLetters(n: number): string {
  if (n < 1) return String(n);
  const letter = String.fromCharCode(65 + ((n - 1) % 26));
  return letter.repeat(Math.floor((n - 1) / 26) + 1);
}

/**
 * A page number as text in a DOCX number format (`w:pgNumType@w:fmt`): `lowerRoman`, `upperRoman`, `lowerLetter` and
 * `upperLetter` are drawn as such; every other format (and no format) is drawn as a decimal number.
 */
export function formatPageNumber(n: number, format?: string): string {
  switch (format) {
    case 'lowerRoman':
      return toRoman(n).toLowerCase();
    case 'upperRoman':
      return toRoman(n);
    case 'lowerLetter':
      return toLetters(n).toLowerCase();
    case 'upperLetter':
      return toLetters(n);
    default:
      return String(n);
  }
}

/** The header/footer of a section is its own, or the one of the previous section when it has none ("link to previous"). */
export function resolveInheritedContent<T>(contents: ReadonlyArray<T | undefined>): Array<T | undefined> {
  const out: Array<T | undefined> = [];
  let previous: T | undefined;
  for (const content of contents) {
    previous = content ?? previous;
    out.push(previous);
  }
  return out;
}

/** The header/footer variant of a page: the section's first page, an even page, or the default. */
export function pickHeaderFooterVariant(
  section: Pick<LayoutSection, 'differentFirstPage' | 'differentOddEven'>,
  isFirstPageOfSection: boolean,
  pageNumber: number,
): HeaderFooterVariant {
  if (section.differentFirstPage && isFirstPageOfSection) return 'first';
  if (section.differentOddEven && pageNumber % 2 === 0) return 'even';
  return 'default';
}

/**
 * Lay out a document of several sections.
 *
 * `blocks` and `measures` cover all sections; each section owns the slice `blockStart..blockEnd`, and blocks must be
 * measured at the content width of their own section. Fragment `blockIndex` values refer to the whole array.
 */
export function computeSectionedLayout(
  blocks: FlowBlock[],
  measures: Measure[],
  sections: LayoutSection[],
  options: SectionedLayoutOptions = {},
): DocumentLayout {
  const { scale = 1, pageGap = 24, minLinesAtBreak = 2 } = options;
  const scaledGap = pageGap * scale;
  if (sections.length === 0) throw new Error('computeSectionedLayout needs at least one section');
  const first = sections[0];

  const pages: PageLayout[] = [];
  let lastPageNumber = 0;

  sections.forEach((section, sectionIndex) => {
    const previous = sectionIndex > 0 ? sections[sectionIndex - 1] : undefined;
    const lastPage = pages[pages.length - 1];
    const continues =
      !!previous &&
      !!lastPage &&
      section.breakType === 'continuous' &&
      canContinueOnPage(lastPage.pageConfig ?? previous.pageConfig, section.pageConfig);

    const laid = computeTrueLayout(blocks.slice(section.blockStart, section.blockEnd), measures.slice(section.blockStart, section.blockEnd), {
      pageConfig: section.pageConfig,
      scale,
      pageGap,
      minLinesAtBreak,
      initialY: continues ? lastPage.contentHeight : 0,
    });
    const own: PageLayout[] = laid.pages.map((page) => ({
      ...page,
      fragments: page.fragments.map((f) => ({ ...f, blockIndex: f.blockIndex + section.blockStart })),
    }));

    let newPages = own;
    if (continues) {
      // The first laid page is the rest of the previous section's last page: its fragments go below the earlier ones.
      const [head, ...rest] = own;
      lastPage.fragments.push(...head.fragments);
      lastPage.contentHeight = Math.max(lastPage.contentHeight, head.contentHeight);
      lastPage.remainingHeight = head.remainingHeight;
      newPages = rest;
    }

    // Starting on an even or odd page may need a blank page first. A section that restarts its numbering takes the
    // number it asks for, so it gets no filler page.
    const restartsAt = section.pageNumbering?.start;
    const wantsParity = section.breakType === 'evenPage' || section.breakType === 'oddPage';
    if (previous && newPages.length > 0 && wantsParity && restartsAt === undefined) {
      const wantEven = section.breakType === 'evenPage';
      const nextIsEven = (lastPageNumber + 1) % 2 === 0;
      if (nextIsEven !== wantEven) {
        const m = previous.pageConfig.margins;
        const blank: PageLayout = {
          pageIndex: 0,
          fragments: [],
          contentHeight: 0,
          remainingHeight: (previous.pageConfig.height - m.top - m.bottom) * scale,
          isBlank: true,
        };
        lastPageNumber += 1;
        pages.push(stamp(blank, sectionIndex - 1, previous, false, lastPageNumber));
      }
    }

    newPages.forEach((page, i) => {
      const startsSection = i === 0 && !continues;
      lastPageNumber = startsSection && restartsAt !== undefined && Number.isInteger(restartsAt) ? restartsAt : lastPageNumber + 1;
      pages.push(stamp(page, sectionIndex, section, startsSection, lastPageNumber));
    });
  });

  // Stack the pages: each page's top follows the previous page's height and a gap.
  let top = 0;
  pages.forEach((page, index) => {
    page.pageIndex = index;
    page.top = top;
    top += (page.pageConfig ?? first.pageConfig).height * scale + scaledGap;
  });

  return {
    pages,
    totalHeight: pages.length > 0 ? top - scaledGap : 0,
    pageConfig: first.pageConfig,
    scale,
    pageGap: scaledGap,
  };
}

/** Stamp a page with the setup, number and header/footer variant of the section it belongs to. */
function stamp(page: PageLayout, sectionIndex: number, section: LayoutSection, isFirstPageOfSection: boolean, pageNumber: number): PageLayout {
  return {
    ...page,
    sectionIndex,
    pageConfig: section.pageConfig,
    pageNumber,
    ...(section.pageNumbering?.format && { pageNumberFormat: section.pageNumbering.format }),
    headerFooterVariant: pickHeaderFooterVariant(section, isFirstPageOfSection, pageNumber),
  };
}
