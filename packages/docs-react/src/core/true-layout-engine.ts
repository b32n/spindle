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
  /** Which header/footer variant of the section the page shows. */
  headerFooterVariant?: HeaderFooterVariant;
  /** True for a page added only so that the next section starts on an even or odd page. */
  isBlank?: boolean;
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
}

// ============================================================================
// Layout Engine
// ============================================================================

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
  const { pageConfig, scale = 1, pageGap = 24, minLinesAtBreak = 2, initialY = 0 } = options;
  
  // Calculate content area dimensions (scaled)
  const contentHeight = (pageConfig.height - pageConfig.margins.top - pageConfig.margins.bottom) * scale;
  
  // Initialize pages
  const pages: PageLayout[] = [];
  let currentPage: PageLayout = createEmptyPage(0, contentHeight);
  let currentY = Math.min(Math.max(initialY, 0), contentHeight);
  // A first page that continues earlier content counts as occupied even before it gets a fragment of its own.
  const carried = currentY > 0;
  const pageOccupied = (): boolean => currentPage.fragments.length > 0 || (carried && pages.length === 0);
  
  // Process each block
  for (let blockIndex = 0; blockIndex < blocks.length; blockIndex++) {
    const block = blocks[blockIndex];
    const measure = measures[blockIndex];
    
    // Handle page breaks
    if (measure.kind === 'pageBreak') {
      // Force a new page
      if (pageOccupied()) {
        currentPage.contentHeight = currentY;
        currentPage.remainingHeight = contentHeight - currentY;
        pages.push(currentPage);
        currentPage = createEmptyPage(pages.length, contentHeight);
        currentY = 0;
      }
      continue;
    }
    
    // Scale the measurement height
    const blockHeight = measure.totalHeight * scale;
    
    // Check if the entire block fits on the current page
    if (currentY + blockHeight <= contentHeight) {
      // Block fits entirely - add as single fragment
      currentPage.fragments.push({
        blockId: block.id,
        blockIndex,
        fromLine: 0,
        toLine: getLineCount(measure),
        x: 0,
        y: currentY,
        height: blockHeight,
        isFirstFragment: true,
        isLastFragment: true,
      });
      currentY += blockHeight;
    } else if (hasLineData(measure) && measure.lines.length > 1) {
      // Block doesn't fit and has multiple lines - try to split
      const result = splitBlockAcrossPages(
        block,
        measure,
        blockIndex,
        currentY,
        contentHeight,
        scale,
        minLinesAtBreak
      );
      
      // Add first fragment to current page (if any)
      if (result.firstFragment) {
        currentPage.fragments.push(result.firstFragment);
        currentY += result.firstFragment.height;
      }
      
      // Push current page and create new pages for remaining fragments
      if (result.remainingFragments.length > 0) {
        currentPage.contentHeight = currentY;
        currentPage.remainingHeight = contentHeight - currentY;
        pages.push(currentPage);
        
        // Add fragments to new pages
        for (let i = 0; i < result.remainingFragments.length; i++) {
          const fragment = result.remainingFragments[i];
          const isLastFragment = i === result.remainingFragments.length - 1;
          
          currentPage = createEmptyPage(pages.length, contentHeight);
          
          // Update fragment position for new page
          fragment.y = 0;
          currentPage.fragments.push(fragment);
          currentY = fragment.height;
          
          // If not the last fragment, push this page too
          if (!isLastFragment) {
            currentPage.contentHeight = currentY;
            currentPage.remainingHeight = contentHeight - currentY;
            pages.push(currentPage);
          }
        }
      }
    } else {
      // Block doesn't fit and can't be split - start new page
      if (pageOccupied()) {
        currentPage.contentHeight = currentY;
        currentPage.remainingHeight = contentHeight - currentY;
        pages.push(currentPage);
        currentPage = createEmptyPage(pages.length, contentHeight);
        currentY = 0;
      }
      
      // Add block to new page (even if it overflows). Use the block's full
      // height — never clamp to contentHeight, or the overflowing remainder of
      // a non-splittable block (a tall table or image) would be silently lost.
      currentPage.fragments.push({
        blockId: block.id,
        blockIndex,
        fromLine: 0,
        toLine: getLineCount(measure),
        x: 0,
        y: 0,
        height: blockHeight,
        isFirstFragment: true,
        isLastFragment: true,
      });
      currentY = blockHeight;
      
      // If block overflows single page, we need to handle that
      // For now, we just let it overflow (tables/images might do this)
      if (blockHeight > contentHeight) {
        currentPage.contentHeight = contentHeight;
        currentPage.remainingHeight = 0;
        pages.push(currentPage);
        currentPage = createEmptyPage(pages.length, contentHeight);
        currentY = 0;
      }
    }
  }
  
  // Don't forget the last page
  if (pageOccupied()) {
    currentPage.contentHeight = currentY;
    currentPage.remainingHeight = contentHeight - currentY;
    pages.push(currentPage);
  }
  
  // Ensure at least one page exists
  if (pages.length === 0) {
    pages.push(createEmptyPage(0, contentHeight));
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

/**
 * Split a block across pages at line boundaries
 */
function splitBlockAcrossPages(
  block: FlowBlock,
  measure: Measure,
  blockIndex: number,
  startY: number,
  pageContentHeight: number,
  scale: number,
  minLinesAtBreak: number
): { firstFragment: PageFragment | null; remainingFragments: PageFragment[] } {
  if (!hasLineData(measure)) {
    // Can't split - return as single fragment
    return {
      firstFragment: {
        blockId: block.id,
        blockIndex,
        fromLine: 0,
        toLine: 1,
        x: 0,
        y: startY,
        height: measure.totalHeight * scale,
        isFirstFragment: true,
        isLastFragment: true,
      },
      remainingFragments: [],
    };
  }
  
  const lines = measure.lines;
  const totalLines = lines.length;
  
  // Find how many lines fit on the current page
  const availableHeight = pageContentHeight - startY;
  const { linesToFit, heightUsed } = findLinesToFit(
    measure,
    0,
    availableHeight,
    scale,
    minLinesAtBreak,
    totalLines
  );
  
  // If no lines fit, don't add a first fragment
  if (linesToFit === 0) {
    // All lines go to next page(s)
    const remainingFragments = createFragmentsForLines(
      block,
      measure,
      blockIndex,
      0,
      totalLines,
      pageContentHeight,
      scale,
      minLinesAtBreak
    );
    return { firstFragment: null, remainingFragments };
  }
  
  // Create first fragment
  const firstFragment: PageFragment = {
    blockId: block.id,
    blockIndex,
    fromLine: 0,
    toLine: linesToFit,
    x: 0,
    y: startY,
    height: heightUsed,
    isFirstFragment: true,
    isLastFragment: linesToFit >= totalLines,
  };
  
  // If all lines fit, we're done
  if (linesToFit >= totalLines) {
    return { firstFragment, remainingFragments: [] };
  }
  
  // Create fragments for remaining lines
  const remainingFragments = createFragmentsForLines(
    block,
    measure,
    blockIndex,
    linesToFit,
    totalLines,
    pageContentHeight,
    scale,
    minLinesAtBreak
  );
  
  return { firstFragment, remainingFragments };
}

/**
 * Find how many lines fit in available height
 */
function findLinesToFit(
  measure: Measure,
  startLine: number,
  availableHeight: number,
  scale: number,
  minLinesAtBreak: number,
  totalLines: number
): { linesToFit: number; heightUsed: number } {
  if (!hasLineData(measure)) {
    return { linesToFit: 1, heightUsed: measure.totalHeight * scale };
  }
  
  let height = 0;
  let lineIndex = startLine;
  
  // Add space before if this is the start of the block
  if (startLine === 0) {
    height += measure.spaceBefore * scale;
  }
  
  // Add lines until we run out of space
  while (lineIndex < totalLines) {
    const lineHeight = measure.lines[lineIndex].height * scale;
    
    // Check if adding this line would exceed available height
    if (height > 0 && height + lineHeight > availableHeight) {
      break;
    }
    
    height += lineHeight;
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

/**
 * Create fragments for a range of lines that span multiple pages
 */
function createFragmentsForLines(
  block: FlowBlock,
  measure: Measure,
  blockIndex: number,
  startLine: number,
  endLine: number,
  pageContentHeight: number,
  scale: number,
  minLinesAtBreak: number
): PageFragment[] {
  if (!hasLineData(measure)) {
    return [];
  }
  
  const fragments: PageFragment[] = [];
  let currentLine = startLine;

  while (currentLine < endLine) {
    const { linesToFit, heightUsed } = findLinesToFit(
      measure,
      currentLine,
      pageContentHeight,
      scale,
      minLinesAtBreak,
      endLine
    );

    // findLinesToFit can return 0 even on a full page — e.g. a single line
    // taller than the whole page, where the widow guard refuses to place fewer
    // than minLinesAtBreak lines. Force one line so the loop always advances;
    // otherwise the remaining lines are silently dropped and the block vanishes.
    const take = linesToFit > 0 ? linesToFit : 1;
    const height =
      linesToFit > 0 ? heightUsed : getLineRangeHeight(measure, currentLine, currentLine + 1) * scale;

    const toLine = currentLine + take;
    
    fragments.push({
      blockId: block.id,
      blockIndex,
      fromLine: currentLine,
      toLine,
      x: 0,
      y: 0, // Will be set by caller
      height,
      isFirstFragment: currentLine === 0,
      isLastFragment: toLine >= endLine,
    });

    currentLine = toLine;
  }
  
  return fragments;
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

