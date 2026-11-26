/**
 * True Layout Engine - Computes page layout with line-level pagination
 * 
 * This is the core of the true layout engine. It takes measured blocks and
 * assigns them to pages, splitting blocks at line boundaries when necessary.
 * 
 * Key concepts:
 * - PageFragment: A portion of a block that appears on a page
 * - PageLayout: All fragments on a single page
 * - DocumentLayout: The complete layout of all pages
 */

import { FlowBlock } from './flow-blocks';
import { Measure, hasLineData, getLineCount, getLineRangeHeight } from './measurer';

// ============================================================================
// Layout Types
// ============================================================================

/**
 * Page configuration for layout
 */
export interface PageConfig {
  /** Page width in pixels */
  width: number;
  /** Page height in pixels */
  height: number;
  /** Page margins */
  margins: {
    top: number;
    bottom: number;
    left: number;
    right: number;
    /** Distance from top edge to header area (default: 48px / 0.5 inch) */
    header?: number;
    /** Distance from bottom edge to footer area (default: 48px / 0.5 inch) */
    footer?: number;
  };
}

/**
 * A fragment of a block that appears on a page
 * 
 * When a block spans multiple pages, it's split into fragments.
 * Each fragment knows which lines of the original block it contains.
 */
export interface PageFragment {
  /** ID of the original block */
  blockId: string;
  /** Index of the original block in the blocks array */
  blockIndex: number;
  /** Starting line index (0-based) */
  fromLine: number;
  /** Ending line index (exclusive) */
  toLine: number;
  /** X position within the page content area */
  x: number;
  /** Y position within the page content area */
  y: number;
  /** Height of this fragment */
  height: number;
  /** Whether this is the first fragment of the block */
  isFirstFragment: boolean;
  /** Whether this is the last fragment of the block */
  isLastFragment: boolean;
}

/** Which header/footer content a page shows when its section has first-page and even-page variants. */
export type HeaderFooterVariant = 'default' | 'first' | 'even';

/**
 * Layout for a single page
 *
 * The optional fields below are set by `computeSectionedLayout` (documents with sections); a layout from
 * `computeTrueLayout` leaves them unset and every page then uses `DocumentLayout.pageConfig`, is numbered by its index
 * and shows the default header and footer.
 */
export interface PageLayout {
  /** Page index (0-based) */
  pageIndex: number;
  /** Fragments on this page */
  fragments: PageFragment[];
  /** Total content height used on this page */
  contentHeight: number;
  /** Available content height remaining */
  remainingHeight: number;
  /** The section whose setup the page uses: the section that starts it (a continuous section shares the page). */
  sectionIndex?: number;
  /** Page setup of this page (unscaled, rendered size), when it can differ from `DocumentLayout.pageConfig`. */
  pageConfig?: PageConfig;
  /** Distance from the top of the first page to the top of this page, scaled, gaps included. */
  top?: number;
  /** The number the page shows (restarts and continues per section), as opposed to its position `pageIndex + 1`. */
  pageNumber?: number;
  /** DOCX number format name of `pageNumber` (e.g. `lowerRoman`); absent for decimal. */
  pageNumberFormat?: string;
  /** True for the first page that starts its section (a page a continuous section merely shares does not count). */
  startsSection?: boolean;
  /** Which header/footer variant the section's switches give the page (a header and a footer may switch differently: the painter picks per content). */
  headerFooterVariant?: HeaderFooterVariant;
  /** True for a page added only so that the next section starts on an even or odd page. */
  isBlank?: boolean;
  /** The footnotes the page carries (set by a layout made with `LayoutOptions.footnotes`; see `applyPageNotes` for drawing them). */
  notes?: PageNotes;
  /** True for a page added only for footnotes that continue past the last page of the body. */
  notesOnly?: boolean;
  /** Where the page's footnotes start (set by `applyPageNotes`): the separator is drawn at `top` (scaled, in the content area), the notes follow it. */
  footnotes?: { top: number; separatorHeight: number; continued: boolean };
}

/** A footnote, or what is left of one, waiting for room on a page: it starts at `fromLine`. */
export interface PendingNote {
  key: string;
  fromLine: number;
}

/** A piece of a footnote on a page. */
export interface PageNote {
  key: string;
  blockIndex: number;
  blockId: string;
  fromLine: number;
  toLine: number;
  /** Lines the note has in all: a piece that ends before this continues on the next page. */
  totalLines: number;
  /** Height (scaled). */
  height: number;
}

/** What a page carries at its bottom: its footnotes (or parts of them), and what the layout needs to go on from it. */
export interface PageNotes {
  items: PageNote[];
  /** Height of the whole area, separator included (scaled); 0 when there are no notes. */
  height: number;
  /** The first note continues from the previous page. */
  continued: boolean;
  /** Notes (or their remaining lines) that did not fit and go on to the next page. */
  carry: PendingNote[];
  /** The footnotes referenced from the lines on the page. */
  keys: string[];
  /** Notes that continued onto this page from the one before. */
  carryIn: PendingNote[];
}

/** How the layout learns about footnotes: which lines refer to which, and what a page's notes take. */
export interface FootnoteHooks {
  /** The footnotes referenced by lines `fromLine`..`toLine` of a block (an index into the array given to the layout), in order. */
  refsIn(blockIndex: number, fromLine: number, toLine: number): string[];
  /** The notes a page of this content height (scaled) carries, with `carry` continuing onto it from the page before and `keys` referenced on it. */
  place(contentHeight: number, carry: PendingNote[], keys: string[]): Pick<PageNotes, 'items' | 'height' | 'continued' | 'carry'>;
}

/**
 * Complete document layout
 */
export interface DocumentLayout {
  /** All pages in the document */
  pages: PageLayout[];
  /** Total height of all pages stacked (including gaps) */
  totalHeight: number;
  /** Page configuration used */
  pageConfig: PageConfig;
  /** Scale factor applied */
  scale: number;
  /** Gap between pages in pixels */
  pageGap: number;
}

/**
 * Options for layout computation
 */
export interface LayoutOptions {
  /** Page configuration */
  pageConfig: PageConfig;
  /** Scale factor (zoom level) */
  scale?: number;
  /** Gap between pages in pixels */
  pageGap?: number;
  /** Minimum number of lines to keep together at page break (widow/orphan control) */
  minLinesAtBreak?: number;
  /**
   * Content height (scaled) already used on the first page by earlier content, for a section that continues on the
   * page of the previous one. The first page of the result then holds only the new fragments, placed below it.
   */
  initialY?: number;
  /**
   * Footnotes at the bottom of their pages. A line (or block) is placed on a page only if it fits together with the
   * notes the page then carries, and a note that does not fit goes on to the next page. Without it there are no notes.
   */
  footnotes?: FootnoteHooks;
  /** For a first page that continues earlier content: the footnotes already on it, and those continuing onto it. */
  initialNotes?: { keys: string[]; carry: PendingNote[] };
}

// ============================================================================
// Layout Engine
// ============================================================================

/** The keys of `base` followed by those of `more` that are not in it. */
function unionKeys(base: readonly string[], more: readonly string[]): string[] {
  if (more.length === 0) return base as string[];
  const out = [...base];
  for (const key of more) if (!out.includes(key)) out.push(key);
  return out;
}

/** Notes still waiting to be placed are at most this many pages past the last page of the body (a guard against a runaway). */
const MAX_NOTES_ONLY_PAGES = 200;

/**
 * Compute the document layout
 * 
 * This function takes blocks and their measurements and assigns them to pages,
 * splitting blocks at line boundaries when necessary.
 * 
 * @param blocks - The FlowBlocks to lay out
 * @param measures - Measurements for each block (must be same length as blocks)
 * @param options - Layout options
 * @returns The computed document layout
 */
export function computeTrueLayout(
  blocks: FlowBlock[],
  measures: Measure[],
  options: LayoutOptions
): DocumentLayout {
  const { pageConfig, scale = 1, pageGap = 24, minLinesAtBreak = 2, initialY = 0, footnotes, initialNotes } = options;
  
  // Calculate content area dimensions (scaled)
  const contentHeight = (pageConfig.height - pageConfig.margins.top - pageConfig.margins.bottom) * scale;
  
  // Initialize pages
  const pages: PageLayout[] = [];
  let currentPage: PageLayout = createEmptyPage(0, contentHeight);
  let currentY = Math.min(Math.max(initialY, 0), contentHeight);
  // A first page that continues earlier content counts as occupied even before it gets a fragment of its own.
  const carried = currentY > 0;
  const pageOccupied = (): boolean => currentPage.fragments.length > 0 || (carried && pages.length === 0);

  // The footnotes of the page being filled: those its lines refer to, and those continuing onto it from the page before.
  let pageKeys: string[] = initialNotes?.keys ?? [];
  let carryIn: PendingNote[] = initialNotes?.carry ?? [];
  /** The height the notes of the current page take, given the footnotes its lines (so far, and `more`) refer to. */
  const notesHeight = (more: readonly string[]): number => (footnotes ? footnotes.place(contentHeight, carryIn, unionKeys(pageKeys, more)).height : 0);

  /** Finish the current page (it goes into `pages`) and start the next one. */
  const closePage = () => {
    const placed = footnotes ? footnotes.place(contentHeight, carryIn, pageKeys) : undefined;
    currentPage.contentHeight = currentY;
    currentPage.remainingHeight = contentHeight - currentY - (placed?.height ?? 0);
    if (placed && (placed.items.length > 0 || carryIn.length > 0 || pageKeys.length > 0)) {
      currentPage.notes = { ...placed, keys: pageKeys, carryIn };
    }
    pages.push(currentPage);
    carryIn = placed?.carry ?? [];
    pageKeys = [];
    currentPage = createEmptyPage(pages.length, contentHeight);
    currentY = 0;
  };
  
  // Process each block
  for (let blockIndex = 0; blockIndex < blocks.length; blockIndex++) {
    const block = blocks[blockIndex];
    const measure = measures[blockIndex];
    
    // Handle page breaks
    if (measure.kind === 'pageBreak') {
      // Force a new page
      if (pageOccupied()) closePage();
      continue;
    }
    
    // Scale the measurement height
    const blockHeight = measure.totalHeight * scale;
    const lineCount = getLineCount(measure);
    const wholeKeys = footnotes ? footnotes.refsIn(blockIndex, 0, lineCount) : [];
    
    // Check if the entire block fits on the current page (with the notes it brings)
    if (currentY + blockHeight + notesHeight(wholeKeys) <= contentHeight) {
      // Block fits entirely - add as single fragment
      currentPage.fragments.push({
        blockId: block.id,
        blockIndex,
        fromLine: 0,
        toLine: lineCount,
        x: 0,
        y: currentY,
        height: blockHeight,
        isFirstFragment: true,
        isLastFragment: true,
      });
      currentY += blockHeight;
      pageKeys = unionKeys(pageKeys, wholeKeys);
    } else if (hasLineData(measure) && measure.lines.length > 1) {
      // Block doesn't fit and has multiple lines - split it at line boundaries
      const totalLines = measure.lines.length;
      let line = 0;
      let first = true;
      while (line < totalLines) {
        const refsOfLine = (l: number) => (footnotes ? footnotes.refsIn(blockIndex, l, l + 1) : []);
        const fit = findLinesToFit(measure, line, contentHeight - (first ? currentY : 0), scale, minLinesAtBreak, totalLines, footnotes && {
          held: pageKeys,
          refsOfLine: refsOfLine,
          heightFor: (keys) => footnotes.place(contentHeight, carryIn, keys).height,
        });
        let take = fit.linesToFit;
        let height = fit.heightUsed;
        if (take === 0 && !first) {
          // findLinesToFit can return 0 even on a full page — e.g. a single line taller than the whole page, where the
          // widow guard refuses to place fewer than minLinesAtBreak lines. Force one line so the loop always advances;
          // otherwise the remaining lines are silently dropped and the block vanishes.
          take = 1;
          height = getLineRangeHeight(measure, line, line + 1) * scale;
        }
        if (take > 0) {
          const toLine = line + take;
          currentPage.fragments.push({
            blockId: block.id,
            blockIndex,
            fromLine: line,
            toLine,
            x: 0,
            y: first ? currentY : 0,
            height,
            isFirstFragment: line === 0,
            isLastFragment: toLine >= totalLines,
          });
          currentY = (first ? currentY : 0) + height;
          pageKeys = unionKeys(pageKeys, footnotes ? footnotes.refsIn(blockIndex, line, toLine) : []);
          line = toLine;
        }
        // Whatever is left goes on a new page.
        if (line < totalLines) closePage();
        first = false;
      }
    } else {
      // Block doesn't fit and can't be split - start new page
      if (pageOccupied()) closePage();
      
      // Add block to new page (even if it overflows). Use the block's full
      // height — never clamp to contentHeight, or the overflowing remainder of
      // a non-splittable block (a tall table or image) would be silently lost.
      currentPage.fragments.push({
        blockId: block.id,
        blockIndex,
        fromLine: 0,
        toLine: lineCount,
        x: 0,
        y: 0,
        height: blockHeight,
        isFirstFragment: true,
        isLastFragment: true,
      });
      currentY = blockHeight;
      pageKeys = unionKeys(pageKeys, wholeKeys);
      
      // If block overflows single page, we need to handle that
      // For now, we just let it overflow (tables/images might do this)
      if (blockHeight > contentHeight) {
        currentY = contentHeight;
        closePage();
      }
    }
  }
  
  // Don't forget the last page
  if (pageOccupied()) closePage();
  
  // Ensure at least one page exists
  if (pages.length === 0) {
    pages.push(createEmptyPage(0, contentHeight));
  }

  // Footnotes still waiting for room get pages of their own after the last one.
  for (let extra = 0; footnotes && carryIn.length > 0 && extra < MAX_NOTES_ONLY_PAGES; extra++) {
    const placed = footnotes.place(contentHeight, carryIn, []);
    const page = createEmptyPage(pages.length, contentHeight);
    page.remainingHeight = contentHeight - placed.height;
    page.notes = { ...placed, keys: [], carryIn };
    page.notesOnly = true;
    pages.push(page);
    carryIn = placed.carry;
  }
  
  // Calculate total height
  const pageHeight = pageConfig.height * scale;
  const scaledPageGap = pageGap * scale;
  const totalHeight = pages.length * pageHeight + (pages.length - 1) * scaledPageGap;
  
  return {
    pages,
    totalHeight,
    pageConfig,
    scale,
    pageGap: scaledPageGap,
  };
}

/**
 * Create an empty page layout
 */
function createEmptyPage(pageIndex: number, contentHeight: number): PageLayout {
  return {
    pageIndex,
    fragments: [],
    contentHeight: 0,
    remainingHeight: contentHeight,
  };
}

/** What `findLinesToFit` needs to know about footnotes: those already on the page, the ones a line refers to, and what notes take. */
interface NotesInFit {
  held: readonly string[];
  refsOfLine: (line: number) => string[];
  heightFor: (keys: string[]) => number;
}

/**
 * Find how many lines fit in available height. With `notes`, a line fits only together with the footnotes the page then
 * carries (their height counts against the available height).
 */
function findLinesToFit(
  measure: Measure,
  startLine: number,
  availableHeight: number,
  scale: number,
  minLinesAtBreak: number,
  totalLines: number,
  notes?: NotesInFit
): { linesToFit: number; heightUsed: number } {
  if (!hasLineData(measure)) {
    return { linesToFit: 1, heightUsed: measure.totalHeight * scale };
  }
  
  let height = 0;
  let lineIndex = startLine;
  let keys: string[] = notes ? [...notes.held] : [];
  
  // Add space before if this is the start of the block
  if (startLine === 0) {
    height += measure.spaceBefore * scale;
  }
  
  // Add lines until we run out of space
  while (lineIndex < totalLines) {
    const lineHeight = measure.lines[lineIndex].height * scale;
    const withLine = notes ? unionKeys(keys, notes.refsOfLine(lineIndex)) : keys;
    const notesHeight = notes ? notes.heightFor(withLine) : 0;
    
    // Check if adding this line would exceed available height
    if (height > 0 && height + lineHeight + notesHeight > availableHeight) {
      break;
    }
    
    height += lineHeight;
    keys = withLine;
    lineIndex++;
  }
  
  // Apply widow/orphan control
  const linesOnThisPage = lineIndex - startLine;
  const linesRemaining = totalLines - lineIndex;
  
  // If we'd leave too few lines on this page, move them all to next page
  if (linesOnThisPage > 0 && linesOnThisPage < minLinesAtBreak && startLine === 0) {
    return { linesToFit: 0, heightUsed: 0 };
  }
  
  // If we'd leave too few lines on the next page, keep more on this page
  if (linesRemaining > 0 && linesRemaining < minLinesAtBreak) {
    // Try to move some lines to the next page
    const linesToMove = minLinesAtBreak - linesRemaining;
    const newLineIndex = Math.max(startLine + minLinesAtBreak, lineIndex - linesToMove);
    
    if (newLineIndex > startLine) {
      // Recalculate height
      height = startLine === 0 ? measure.spaceBefore * scale : 0;
      for (let i = startLine; i < newLineIndex; i++) {
        height += measure.lines[i].height * scale;
      }
      return { linesToFit: newLineIndex - startLine, heightUsed: height };
    }
  }
  
  // Add space after if this is the end of the block
  if (lineIndex >= totalLines) {
    height += measure.spaceAfter * scale;
  }
  
  return { linesToFit: lineIndex - startLine, heightUsed: height };
}

// ============================================================================
// Layout Query Functions
// ============================================================================

/**
 * Find which page contains a given block
 */
export function findBlockPage(layout: DocumentLayout, blockId: string): number | null {
  for (const page of layout.pages) {
    for (const fragment of page.fragments) {
      if (fragment.blockId === blockId) {
        return page.pageIndex;
      }
    }
  }
  return null;
}

/**
 * Find all fragments for a given block
 */
export function findBlockFragments(layout: DocumentLayout, blockId: string): PageFragment[] {
  const fragments: PageFragment[] = [];
  for (const page of layout.pages) {
    for (const fragment of page.fragments) {
      if (fragment.blockId === blockId) {
        fragments.push(fragment);
      }
    }
  }
  return fragments;
}

/**
 * Get the Y coordinate for a page (accounting for gaps)
 */
export function getPageY(layout: DocumentLayout, pageIndex: number): number {
  const page = layout.pages[pageIndex];
  if (page?.top !== undefined) return page.top;
  const pageHeight = layout.pageConfig.height * layout.scale;
  return pageIndex * (pageHeight + layout.pageGap);
}

/** The page setup of a page: its own when the layout has sections, else the document's. */
export function getPageConfig(layout: DocumentLayout, pageIndex: number): PageConfig {
  return layout.pages[pageIndex]?.pageConfig ?? layout.pageConfig;
}

/** The widest page of the layout, unscaled: pages of a narrower section are centred within it. */
export function getMaxPageWidth(layout: DocumentLayout): number {
  let max = layout.pageConfig.width;
  for (const page of layout.pages) if (page.pageConfig && page.pageConfig.width > max) max = page.pageConfig.width;
  return max;
}

/** Distance from the left edge of the page stack to the left edge of a page, scaled (0 when all pages are as wide). */
export function getPageX(layout: DocumentLayout, pageIndex: number): number {
  return ((getMaxPageWidth(layout) - getPageConfig(layout, pageIndex).width) / 2) * layout.scale;
}

/** Where the content area of a page starts in the page stack (scaled): the page's position plus its margins. */
export function getPageContentOrigin(layout: DocumentLayout, pageIndex: number): { x: number; y: number } {
  const margins = getPageConfig(layout, pageIndex).margins;
  return {
    x: getPageX(layout, pageIndex) + margins.left * layout.scale,
    y: getPageY(layout, pageIndex) + margins.top * layout.scale,
  };
}

/**
 * Find which page a Y coordinate falls on
 */
export function findPageAtY(layout: DocumentLayout, y: number): number {
  const last = layout.pages.length - 1;
  if (layout.pages.some((p) => p.top !== undefined)) {
    // Pages of different heights: the last page that starts at or above y (a gap belongs to the page above it).
    let found = 0;
    for (let i = 0; i <= last; i++) {
      if (getPageY(layout, i) <= y) found = i;
      else break;
    }
    return found;
  }
  const pageHeight = layout.pageConfig.height * layout.scale;
  const pageWithGap = pageHeight + layout.pageGap;
  
  const pageIndex = Math.floor(y / pageWithGap);
  return Math.min(Math.max(0, pageIndex), last);
}

/**
 * Convert a position within a block to page coordinates
 */
export function blockPositionToPageCoords(
  layout: DocumentLayout,
  blockId: string,
  lineIndex: number,
  xOffset: number
): { pageIndex: number; x: number; y: number } | null {
  // Find the fragment containing this line
  for (const page of layout.pages) {
    for (const fragment of page.fragments) {
      if (fragment.blockId === blockId) {
        if (lineIndex >= fragment.fromLine && lineIndex < fragment.toLine) {
          // This fragment contains the line
          // Calculate Y offset within the fragment
          // (This is simplified - a full implementation would use line measurements)
          const lineWithinFragment = lineIndex - fragment.fromLine;
          const estimatedLineHeight = fragment.height / (fragment.toLine - fragment.fromLine);
          const yWithinFragment = lineWithinFragment * estimatedLineHeight;
          
          return {
            pageIndex: page.pageIndex,
            x: fragment.x + xOffset,
            y: fragment.y + yWithinFragment,
          };
        }
      }
    }
  }
  return null;
}

