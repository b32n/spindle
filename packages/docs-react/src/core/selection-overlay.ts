/**
 * Selection Overlay - Renders cursor and selection highlights
 * 
 * Since the visible content is rendered separately from the ProseMirror editor,
 * we need to render the cursor and selection as an overlay on top of the pages.
 * 
 * This component:
 * 1. Maps ProseMirror selection to page coordinates
 * 2. Renders a blinking caret for collapsed selection
 * 3. Renders highlight rectangles for range selection
 */

import { EditorView } from 'prosemirror-view';
import { EditorState } from 'prosemirror-state';
import { FlowBlock } from './flow-blocks';
import { Measure, hasLineData } from './measurer';
import { DocumentLayout, getPageConfig, getPageContentOrigin } from './true-layout-engine';
import { createBlockPositionMap } from './pm-to-blocks';
import { Node as PmNode } from 'prosemirror-model';

// ============================================================================
// Types
// ============================================================================

/**
 * Caret position information
 */
export interface CaretPosition {
  /** Page index */
  pageIndex: number;
  /** X coordinate within page content area */
  x: number;
  /** Y coordinate within page content area */
  y: number;
  /** Height of the caret */
  height: number;
}

/**
 * Selection rectangle
 */
export interface SelectionRect {
  /** Page index */
  pageIndex: number;
  /** X coordinate within page content area */
  x: number;
  /** Y coordinate within page content area */
  y: number;
  /** Width of the rectangle */
  width: number;
  /** Height of the rectangle */
  height: number;
}

/**
 * Complete selection state
 */
export interface SelectionState {
  /** Caret position (if selection is collapsed) */
  caret: CaretPosition | null;
  /** Selection rectangles (if selection is a range) */
  rects: SelectionRect[];
  /** Whether the editor is focused */
  isFocused: boolean;
}

// ============================================================================
// Selection Overlay Class
// ============================================================================

/**
 * Cached span position entry for fast lookup
 */
interface SpanPositionEntry {
  element: HTMLElement;
  pmStart: number;
  pmEnd: number;
  pageIndex: number;
}

/**
 * SelectionOverlay - Manages selection rendering
 */
export class SelectionOverlayManager {
  private layout: DocumentLayout | null = null;
  private blocks: FlowBlock[] = [];
  private measures: Measure[] = [];
  private blockPositionMap: Map<string, { start: number; end: number }> = new Map();
  private overlayContainer: HTMLElement | null = null;
  private mount: HTMLElement | null = null;  // Reference to pages container for DOM queries
  private caretElement: HTMLElement | null = null;
  private selectionElements: HTMLElement[] = [];
  private selectionPool: HTMLElement[] = [];  // Pool of reusable selection rect elements
  private isFocused: boolean = false;
  private animationFrame: number | null = null;
  private cellSelectionActive: boolean = false;  // When true, hide caret (cell is selected)
  
  // Performance: cache last selection to skip redundant updates
  private lastSelectionFrom: number = -1;
  private lastSelectionTo: number = -1;
  
  // Performance: cache span positions for fast lookup (sorted by pmStart)
  private spanPositionCache: SpanPositionEntry[] = [];
  private spanCacheValid: boolean = false;
  
  // Performance: cache contentArea rects per page
  // Invalidated on scroll (tracked via scrollVersion) or layout changes
  private contentRectCache: Map<number, DOMRect> = new Map();
  private lastScrollVersion = 0;
  private currentScrollVersion = 0;
  
  // Track layout version to force selection re-render when layout changes
  private layoutVersion = 0;
  private lastRenderedLayoutVersion = 0;
  
  // Performance: cache last selection rects to avoid redundant renders
  private lastRectsHash: string = '';
  
  // Scroll container for invalidating contentRectCache on scroll
  private scrollContainer: HTMLElement | null = null;
  private scrollHandler: (() => void) | null = null;
  
  constructor() {}
  
  /**
   * Set the scroll container to track scroll events.
   * ContentRectCache is invalidated on scroll since getBoundingClientRect() 
   * returns viewport-relative coordinates that change when scrolling.
   */
  setScrollContainer(container: HTMLElement | null): void {
    // Remove old listener
    if (this.scrollContainer && this.scrollHandler) {
      this.scrollContainer.removeEventListener('scroll', this.scrollHandler);
    }
    
    this.scrollContainer = container;
    
    if (container) {
      this.scrollHandler = () => {
        // Increment scroll version - cache will be invalidated on next selection
        this.currentScrollVersion++;
      };
      container.addEventListener('scroll', this.scrollHandler, { passive: true });
    }
  }
  
  /**
   * Initialize the overlay with a container element
   * @param container - The overlay container element
   * @param pagesMount - Optional reference to pages container for DOM-based position mapping
   */
  initialize(container: HTMLElement, pagesMount?: HTMLElement): void {
    this.overlayContainer = container;
    this.mount = pagesMount || null;
    
    // Create caret element - use transform for GPU-accelerated positioning
    this.caretElement = document.createElement('div');
    this.caretElement.className = 'selection-caret';
    this.caretElement.style.cssText = `
      position: absolute;
      top: 0;
      left: 0;
      width: 2px;
      background-color: #000;
      pointer-events: none;
      z-index: 100;
      display: none;
    `;
    container.appendChild(this.caretElement);
    
    // Add caret blink animation
    const style = document.createElement('style');
    style.textContent = `
      @keyframes caret-blink {
        0%, 100% { opacity: 1; }
        50% { opacity: 0; }
      }
      .selection-caret.blinking {
        animation: caret-blink 1s step-end infinite;
      }
      .selection-caret {
        will-change: transform, left, top;
        contain: layout style;
      }
      .selection-rect {
        position: absolute;
        top: 0;
        left: 0;
        background-color: rgba(66, 133, 244, 0.3);
        pointer-events: none;
        will-change: transform;
        contain: layout style;
        z-index: 99;
        will-change: transform;
      }
    `;
    container.appendChild(style);
  }
  
  /**
   * Update layout data
   */
  updateLayout(
    layout: DocumentLayout,
    blocks: FlowBlock[],
    measures: Measure[],
    doc: PmNode
  ): void {
    this.layout = layout;
    this.blocks = blocks;
    this.measures = measures;
    this.blockPositionMap = createBlockPositionMap(doc, blocks);
    
    // Invalidate caches when layout changes
    this.spanCacheValid = false;
    this.contentRectCache.clear();
    this.lastRectsHash = '';
    
    // Increment layout version to force selection re-render
    // (selection positions may not change but visual rects do when text is resized)
    this.layoutVersion++;
  }
  
  /**
   * Build the span position cache from DOM
   * This is done once per layout update, not per selection change
   */
  private buildSpanPositionCache(): void {
    if (!this.mount) {
      this.spanPositionCache = [];
      this.spanCacheValid = true;
      return;
    }
    
    const entries: SpanPositionEntry[] = [];
    const spans = this.mount.querySelectorAll('span[data-pm-start], a[data-pm-start]');
    
    for (const span of spans) {
      const htmlSpan = span as HTMLElement;
      const pmStart = parseInt(htmlSpan.dataset.pmStart || '', 10);
      const pmEnd = parseInt(htmlSpan.dataset.pmEnd || '', 10);
      
      if (!Number.isFinite(pmStart) || !Number.isFinite(pmEnd)) continue;
      
      const pageEl = htmlSpan.closest('.page') as HTMLElement | null;
      const pageIndex = pageEl ? parseInt(pageEl.dataset.pageIndex || '0', 10) : 0;
      
      entries.push({ element: htmlSpan, pmStart, pmEnd, pageIndex });
    }
    
    // Sort by pmStart for binary search
    entries.sort((a, b) => a.pmStart - b.pmStart);
    
    this.spanPositionCache = entries;
    this.spanCacheValid = true;
  }
  
  /**
   * Find spans that overlap with a position range using binary search
   */
  private findOverlappingSpans(minPos: number, maxPos: number): SpanPositionEntry[] {
    if (!this.spanCacheValid) {
      this.buildSpanPositionCache();
    }
    
    const cache = this.spanPositionCache;
    if (cache.length === 0) return [];
    
    // Binary search to find starting point
    let left = 0;
    let right = cache.length;
    while (left < right) {
      const mid = (left + right) >> 1;
      if (cache[mid].pmEnd <= minPos) {
        left = mid + 1;
      } else {
        right = mid;
      }
    }
    
    // Collect all overlapping spans
    const result: SpanPositionEntry[] = [];
    for (let i = left; i < cache.length; i++) {
      const entry = cache[i];
      if (entry.pmStart >= maxPos) break;  // No more overlapping spans
      if (entry.pmEnd > minPos && entry.pmStart < maxPos) {
        result.push(entry);
      }
    }
    
    return result;
  }
  
  /**
   * Set the pages container reference for DOM-based position mapping
   */
  setPagesMount(mount: HTMLElement | null): void {
    if (this.mount !== mount) {
      this.mount = mount;
      this.spanCacheValid = false;  // Invalidate cache when mount changes
    }
  }
  
  /**
   * Invalidate the span position cache (call after DOM changes)
   */
  invalidateSpanCache(): void {
    this.spanCacheValid = false;
  }
  
  /**
   * Set whether a cell is currently selected (hides caret when true)
   */
  setCellSelectionActive(active: boolean): void {
    this.cellSelectionActive = active;
    
    // Hide caret when cell is selected
    if (this.caretElement) {
      if (active) {
        this.caretElement.style.display = 'none';
      }
    }
  }
  
  /**
   * Update focus state
   */
  setFocused(focused: boolean): void {
    this.isFocused = focused;
    
    if (this.caretElement) {
      // Always keep caret visible, just toggle blinking
      if (focused) {
        this.caretElement.classList.add('blinking');
      } else {
        this.caretElement.classList.remove('blinking');
      }
    }
  }
  
  /**
   * Update selection from editor state
   * Selection rendering is always batched via RAF for performance.
   * Multiple calls per frame are coalesced into a single render.
   * 
   * @param immediate - If true, render synchronously (bypasses RAF for drag responsiveness)
   */
  updateSelection(state: EditorState, view: EditorView | null, immediate = false): void {
    if (!this.layout || !this.overlayContainer) return;
    
    const { selection } = state;
    
    // Check if layout changed (text resize, etc.) - force re-render even if positions same
    const layoutChanged = this.layoutVersion !== this.lastRenderedLayoutVersion;
    
    // Early skip: if selection hasn't changed AND layout hasn't changed, skip
    if (!layoutChanged && selection.from === this.lastSelectionFrom && selection.to === this.lastSelectionTo) {
      return;
    }
    
    // Update the cached values immediately to prevent duplicate calls
    this.lastSelectionFrom = selection.from;
    this.lastSelectionTo = selection.to;
    
    // If immediate, render now and cancel any pending RAF
    if (immediate) {
      // Cancel pending RAF since we're rendering now
      if (this.animationFrame !== null) {
        cancelAnimationFrame(this.animationFrame);
        this.animationFrame = null;
      }
      this.pendingState = null;
      this.pendingView = null;
      
      this.renderSelection(state, view);
      return;
    }
    
    // Store the latest state for batched rendering
    this.pendingState = state;
    this.pendingView = view;
    
    // Schedule update for next frame if not already scheduled
    if (this.animationFrame === null) {
      this.animationFrame = requestAnimationFrame(() => {
        this.animationFrame = null;
        if (this.pendingState) {
          this.renderSelection(this.pendingState, this.pendingView);
          this.pendingState = null;
          this.pendingView = null;
        }
      });
    }
  }
  
  // Pending state for batched rendering
  private pendingState: EditorState | null = null;
  private pendingView: EditorView | null = null;
  
  /**
   * Render the selection
   */
  private renderSelection(state: EditorState, view: EditorView | null): void {
    if (!this.layout || !this.overlayContainer || !this.caretElement) return;
    
    const { selection } = state;
    
    // Track that we've rendered for this layout version
    this.lastRenderedLayoutVersion = this.layoutVersion;
    
    // Return unused elements to pool instead of destroying
    this.returnRectsToPool();
    
    if (selection.empty) {
      // Collapsed selection - show caret
      this.lastRectsHash = '';  // Reset rects hash for next range selection
      const caretPos = this.positionToCaret(selection.from, view);
      
      if (caretPos) {
        this.renderCaret(caretPos);
      } else {
        this.hideCaret();
      }
    } else {
      // Range selection - show rectangles
      this.hideCaret();
      
      const rects = this.selectionToRects(selection.from, selection.to, view);
      
      // Quick hash to skip redundant renders (common during fast drag)
      const rectsHash = rects.length > 0 
        ? `${rects.length}:${rects[0].x.toFixed(0)},${rects[0].y.toFixed(0)}:${rects[rects.length-1].x.toFixed(0)},${rects[rects.length-1].width.toFixed(0)}`
        : '';
      
      if (rectsHash === this.lastRectsHash) {
        return;  // Skip render - rects haven't meaningfully changed
      }
      this.lastRectsHash = rectsHash;
      
      this.renderSelectionRectsPooled(rects);
    }
  }
  
  /**
   * Get the text length of a block
   */
  private getBlockTextLength(block: FlowBlock): number {
    if (block.kind === 'paragraph' || block.kind === 'heading') {
      return block.runs.reduce((len, r) => len + (r.kind === 'text' ? r.text.length : 0), 0);
    } else if (block.kind === 'listItem') {
      return block.runs.reduce((len, r) => len + (r.kind === 'text' ? r.text.length : 0), 0);
    }
    return 0;
  }
  
  /**
   * Convert a ProseMirror position to caret coordinates
   */
  /**
   * Translate a PM position to a CaretPosition on the visible pages.
   * Public surface for collaborator-cursor overlays, which feed in PM
   * positions decoded from awareness state.
   */
  getCaretForPos(pos: number): CaretPosition | null {
    return this.positionToCaret(pos, null);
  }

  /**
   * Translate a PM range to SelectionRects on the visible pages. Public
   * surface for collaborator-selection overlays.
   */
  getRectsForRange(from: number, to: number): SelectionRect[] {
    return this.selectionToRects(from, to, null);
  }

  /** Public read-only access to the current layout (for collaborator overlays). */
  getLayout(): DocumentLayout | null {
    return this.layout;
  }

  private positionToCaret(pos: number, _view: EditorView | null): CaretPosition | null {
    if (!this.layout) return null;
    
    // Try DOM-based mapping first (more accurate when PM positions are in the DOM)
    const domPos = this.positionToCaretDom(pos);
    if (domPos) {
      return domPos;
    }
    
    // Fallback to estimation from block data and layout
    return this.estimateCaretPosition(pos);
  }
  
  /**
   * DOM-based position-to-caret mapping.
   * Finds the span with matching PM position and calculates exact coordinates.
   */
  private positionToCaretDom(pos: number): CaretPosition | null {
    if (!this.mount) return null;
    
    // Find a span that contains this position
    const spans = this.mount.querySelectorAll('span[data-pm-start], a[data-pm-start]');
    
    for (const span of spans) {
      const htmlSpan = span as HTMLElement;
      const pmStart = parseInt(htmlSpan.dataset.pmStart || '', 10);
      const pmEnd = parseInt(htmlSpan.dataset.pmEnd || '', 10);
      
      if (!Number.isFinite(pmStart) || !Number.isFinite(pmEnd)) continue;
      
      // Check if position is within this span
      if (pos >= pmStart && pos <= pmEnd) {
        // Found the span! Now we need to find its page and calculate coordinates
        const pageEl = htmlSpan.closest('.page') as HTMLElement | null;
        const contentArea = pageEl?.querySelector('.page-content') as HTMLElement | null;
        
        // If we can't find the page structure, fall back to estimation
        if (!pageEl || !contentArea) {
          continue; // Try next span or fall back to estimateCaretPosition
        }
        
        const rect = htmlSpan.getBoundingClientRect();
        const contentRect = contentArea.getBoundingClientRect();
        const pageIndex = parseInt(pageEl.dataset.pageIndex || '0', 10);
        
        // Calculate X position using character offset
        let xPos = 0;
        const charOffset = pos - pmStart;
        const textNode = htmlSpan.firstChild;
        
        if (textNode && textNode.nodeType === Node.TEXT_NODE && textNode.textContent) {
          const text = textNode.textContent;
          if (charOffset <= 0) {
            xPos = 0;
          } else if (charOffset >= text.length) {
            xPos = rect.width;
          } else {
            // Use Range API to measure exact X position
            const range = document.createRange();
            range.setStart(textNode, 0);
            range.setEnd(textNode, Math.min(charOffset, text.length));
            const charRect = range.getBoundingClientRect();
            xPos = charRect.width;
          }
        } else {
          // No text node - use simple ratio
          const pmRange = pmEnd - pmStart;
          if (pmRange > 0) {
            xPos = (charOffset / pmRange) * rect.width;
          }
        }
        
        // Calculate positions relative to page content area
        const baseX = rect.left - contentRect.left + xPos;
        const baseY = rect.top - contentRect.top;
        
        return {
          pageIndex,
          x: baseX,
          y: baseY,
          height: rect.height,
        };
      }
    }
    
    // Check for table cell paragraphs (including empty cells)
    const cellParagraphs = this.mount.querySelectorAll('.cell-paragraph[data-pm-start]');
    for (const para of cellParagraphs) {
      const htmlPara = para as HTMLElement;
      const pmStart = parseInt(htmlPara.dataset.pmStart || '', 10);
      const pmEnd = parseInt(htmlPara.dataset.pmEnd || '', 10);
      
      if (!Number.isFinite(pmStart) || !Number.isFinite(pmEnd)) continue;
      
      // Check if position is within this paragraph
      if (pos >= pmStart && pos <= pmEnd) {
        const pageEl = htmlPara.closest('.page') as HTMLElement | null;
        const contentArea = pageEl?.querySelector('.page-content') as HTMLElement | null;
        
        if (!pageEl || !contentArea) continue;
        
        const rect = htmlPara.getBoundingClientRect();
        const contentRect = contentArea.getBoundingClientRect();
        const pageIndex = parseInt(pageEl.dataset.pageIndex || '0', 10);
        
        // For cell paragraphs, try to find the exact span position first
        const childSpans = htmlPara.querySelectorAll('span[data-pm-start]');
        for (const childSpan of childSpans) {
          const spanEl = childSpan as HTMLElement;
          const spanStart = parseInt(spanEl.dataset.pmStart || '', 10);
          const spanEnd = parseInt(spanEl.dataset.pmEnd || '', 10);
          
          if (Number.isFinite(spanStart) && Number.isFinite(spanEnd) && 
              pos >= spanStart && pos <= spanEnd) {
            const spanRect = spanEl.getBoundingClientRect();
            const charOffset = pos - spanStart;
            const textNode = spanEl.firstChild;
            let xPos = 0;
            
            if (textNode && textNode.nodeType === Node.TEXT_NODE && textNode.textContent) {
              const text = textNode.textContent;
              if (charOffset <= 0) {
                xPos = 0;
              } else if (charOffset >= text.length) {
                xPos = spanRect.width;
              } else {
                const range = document.createRange();
                range.setStart(textNode, 0);
                range.setEnd(textNode, Math.min(charOffset, text.length));
                xPos = range.getBoundingClientRect().width;
              }
            }
            
            return {
              pageIndex,
              x: spanRect.left - contentRect.left + xPos,
              y: spanRect.top - contentRect.top,
              height: spanRect.height || rect.height,
            };
          }
        }
        
        // No span found (empty paragraph) - position cursor at start of paragraph
        // Use the paragraph's position and a reasonable default height
        const paraHeight = Math.max(rect.height, 16 * (this.layout?.scale || 1));
        
        return {
          pageIndex,
          x: rect.left - contentRect.left,
          y: rect.top - contentRect.top,
          height: paraHeight,
        };
      }
    }
    
    // Check for table cells directly (for positions at cell boundaries)
    const tableCells = this.mount.querySelectorAll('.table-cell[data-pm-start]');
    for (const cell of tableCells) {
      const htmlCell = cell as HTMLElement;
      const pmStart = parseInt(htmlCell.dataset.pmStart || '', 10);
      const pmEnd = parseInt(htmlCell.dataset.pmEnd || '', 10);
      
      if (!Number.isFinite(pmStart) || !Number.isFinite(pmEnd)) continue;
      
      // Check if position is within this cell
      if (pos >= pmStart && pos <= pmEnd) {
        const pageEl = htmlCell.closest('.page') as HTMLElement | null;
        const contentArea = pageEl?.querySelector('.page-content') as HTMLElement | null;
        
        if (!pageEl || !contentArea) continue;
        
        // Find the first paragraph in the cell
        const firstPara = htmlCell.querySelector('.cell-paragraph') as HTMLElement | null;
        if (firstPara) {
          const paraRect = firstPara.getBoundingClientRect();
          const contentRect = contentArea.getBoundingClientRect();
          const pageIndex = parseInt(pageEl.dataset.pageIndex || '0', 10);
          
          return {
            pageIndex,
            x: paraRect.left - contentRect.left,
            y: paraRect.top - contentRect.top,
            height: Math.max(paraRect.height, 16 * (this.layout?.scale || 1)),
          };
        }
        
        // Fallback to cell position
        const cellRect = htmlCell.getBoundingClientRect();
        const contentRect = contentArea.getBoundingClientRect();
        const pageIndex = parseInt(pageEl.dataset.pageIndex || '0', 10);
        
        return {
          pageIndex,
          x: cellRect.left - contentRect.left + 8, // Add padding offset
          y: cellRect.top - contentRect.top + 4,   // Add padding offset
          height: 16 * (this.layout?.scale || 1),
        };
      }
    }
    
    // Also check for position between spans (at block boundaries)
    const lines = this.mount.querySelectorAll('.line[data-pm-start]');
    for (const line of lines) {
      const htmlLine = line as HTMLElement;
      const pmStart = parseInt(htmlLine.dataset.pmStart || '', 10);
      const pmEnd = parseInt(htmlLine.dataset.pmEnd || '', 10);
      
      if (!Number.isFinite(pmStart) || !Number.isFinite(pmEnd)) continue;
      
      // Check if position is at the start or end of this line
      if (pos === pmStart || pos === pmEnd) {
        const pageEl = htmlLine.closest('.page') as HTMLElement | null;
        const contentArea = pageEl?.querySelector('.page-content') as HTMLElement | null;
        
        // If we can't find the page structure, skip
        if (!pageEl || !contentArea) continue;
        
        const rect = htmlLine.getBoundingClientRect();
        const contentRect = contentArea.getBoundingClientRect();
        const pageIndex = parseInt(pageEl.dataset.pageIndex || '0', 10);
        
        const baseX = pos === pmStart ? 0 : rect.width;
        const baseY = rect.top - contentRect.top;
        
        return {
          pageIndex,
          x: rect.left - contentRect.left + baseX,
          y: baseY,
          height: rect.height,
        };
      }
    }
    
    // Check for fragment elements (paragraph/heading blocks)
    // This handles positions at the very start of a paragraph (before any span)
    const fragments = this.mount.querySelectorAll('.fragment[data-pm-start]');
    for (const fragment of fragments) {
      const htmlFragment = fragment as HTMLElement;
      const pmStart = parseInt(htmlFragment.dataset.pmStart || '', 10);
      const pmEnd = parseInt(htmlFragment.dataset.pmEnd || '', 10);
      
      if (!Number.isFinite(pmStart) || !Number.isFinite(pmEnd)) continue;
      
      // Skip table fragments - tables should not have cursor positioned via fragment fallback
      if (htmlFragment.querySelector('.table-block')) {
        continue;
      }
      
      // Check if position is within this fragment (covers start of paragraph case)
      // Position pmStart + 1 is the "inside paragraph but before text" position
      if (pos >= pmStart && pos <= pmEnd) {
        const pageEl = htmlFragment.closest('.page') as HTMLElement | null;
        const contentArea = pageEl?.querySelector('.page-content') as HTMLElement | null;
        
        if (!pageEl || !contentArea) continue;
        
        // Try to find the first line or span in this fragment
        const firstLine = htmlFragment.querySelector('.line') as HTMLElement | null;
        const firstSpan = htmlFragment.querySelector('span[data-pm-start]') as HTMLElement | null;
        
        const contentRect = contentArea.getBoundingClientRect();
        const pageIndex = parseInt(pageEl.dataset.pageIndex || '0', 10);
        
        if (firstLine) {
          const lineRect = firstLine.getBoundingClientRect();
          return {
            pageIndex,
            x: lineRect.left - contentRect.left,
            y: lineRect.top - contentRect.top,
            height: lineRect.height,
          };
        }
        
        if (firstSpan) {
          const spanRect = firstSpan.getBoundingClientRect();
          return {
            pageIndex,
            x: spanRect.left - contentRect.left,
            y: spanRect.top - contentRect.top,
            height: spanRect.height,
          };
        }
        
        // Fallback to fragment position - cap height to reasonable line height
        const fragRect = htmlFragment.getBoundingClientRect();
        const defaultLineHeight = 20 * (this.layout?.scale || 1);
        return {
          pageIndex,
          x: fragRect.left - contentRect.left,
          y: fragRect.top - contentRect.top,
          // Cap height to prevent huge cursors for large blocks
          height: Math.min(fragRect.height, defaultLineHeight * 2),
        };
      }
    }
    
    return null;
  }
  
  /**
   * Estimate caret position from block data
   */
  private estimateCaretPosition(pos: number): CaretPosition | null {
    if (!this.layout || this.layout.pages.length === 0) return null;
    
    const defaultHeight = 20 * this.layout.scale;
    
    // If no blocks or position map, return default position at first fragment
    if (this.blockPositionMap.size === 0 || this.blocks.length === 0) {
      const firstPage = this.layout.pages[0];
      if (firstPage && firstPage.fragments.length > 0) {
        return {
          pageIndex: 0,
          x: 0,
          y: firstPage.fragments[0].y,
          height: defaultHeight,
        };
      }
      return {
        pageIndex: 0,
        x: 0,
        y: 0,
        height: defaultHeight,
      };
    }
    
    // Find which block contains this position by iterating blocks in order
    for (let blockIndex = 0; blockIndex < this.blocks.length; blockIndex++) {
      const block = this.blocks[blockIndex];
      const posInfo = this.blockPositionMap.get(block.id);
      if (!posInfo) continue;
      
      // Skip table blocks - cursor positioning inside tables is handled by DOM-based mapping
      // (tables have their own cell-level positioning via positionToCaretDom)
      if (block.kind === 'table') {
        continue;
      }
      
      // Check if position is within this block (inclusive on both ends)
      if (pos >= posInfo.start && pos <= posInfo.end) {
        const measure = this.measures[blockIndex];
        
        // Calculate character offset within block
        // ProseMirror position: posInfo.start is before the node, +1 enters the node
        // Clamp charOffset to valid range
        const maxCharOffset = this.getBlockTextLength(block);
        const charOffset = Math.max(0, Math.min(pos - posInfo.start - 1, maxCharOffset));
        
        // Get line data with fallback
        const hasLines = measure && hasLineData(measure) && measure.lines.length > 0;
        const lines = hasLines ? measure.lines : [];
        
        // Calculate total characters from segments, with fallback to block text length
        let totalCharsFromSegments = 0;
        for (const line of lines) {
          if (line.segments) {
            for (const seg of line.segments) {
              totalCharsFromSegments += seg.text.length;
            }
          }
        }
        
        // If segments are empty, use block text length as fallback
        const useSegments = totalCharsFromSegments > 0;
        const totalChars = useSegments ? totalCharsFromSegments : maxCharOffset;
        
        // Find which line contains this character offset
        let targetLineIndex = 0;
        let charsBeforeLine = 0;
        let lineHeight = defaultHeight;
        
        if (lines.length > 0) {
          if (useSegments) {
            // Use segment data for accurate positioning
            let charCount = 0;
            for (let i = 0; i < lines.length; i++) {
              const line = lines[i];
              let lineChars = 0;
              if (line.segments) {
                for (const seg of line.segments) {
                  lineChars += seg.text.length;
                }
              }
              
              if (charOffset < charCount + lineChars) {
                targetLineIndex = i;
                charsBeforeLine = charCount;
                lineHeight = line.height * this.layout.scale;
                break;
              }
              
              charCount += lineChars;
              targetLineIndex = i;
              charsBeforeLine = charCount;
              lineHeight = line.height * this.layout.scale;
            }
          } else {
            // Fallback: distribute characters evenly across lines
            const charsPerLine = Math.max(1, Math.ceil(totalChars / lines.length));
            targetLineIndex = Math.min(Math.floor(charOffset / charsPerLine), lines.length - 1);
            charsBeforeLine = targetLineIndex * charsPerLine;
            lineHeight = lines[targetLineIndex].height * this.layout.scale;
          }
        }
        
        // Find the fragment that contains this line
        for (const page of this.layout.pages) {
          for (const fragment of page.fragments) {
            if (fragment.blockId === block.id) {
              // Check if this fragment contains the target line
              if (targetLineIndex >= fragment.fromLine && targetLineIndex < fragment.toLine) {
                // Calculate Y position within fragment
                let yWithinFragment = 0;
                for (let i = fragment.fromLine; i < targetLineIndex && i < lines.length; i++) {
                  yWithinFragment += lines[i].height * this.layout.scale;
                }
                
                // Estimate X position based on character offset within line
                let xPos = 0;
                const charsInLine = Math.max(0, charOffset - charsBeforeLine);
                
                if (useSegments && lines[targetLineIndex]) {
                  const line = lines[targetLineIndex];
                  let totalLineChars = 0;
                  if (line.segments) {
                    for (const seg of line.segments) {
                      totalLineChars += seg.text.length;
                    }
                  }
                  if (totalLineChars > 0 && line.width > 0) {
                    const ratio = Math.min(charsInLine / totalLineChars, 1);
                    xPos = ratio * line.width * this.layout.scale;
                  }
                } else if (lines[targetLineIndex]) {
                  // Fallback: use character ratio * line width
                  const line = lines[targetLineIndex];
                  const charsPerLine = Math.max(1, Math.ceil(totalChars / lines.length));
                  if (line.width > 0 && charsPerLine > 0) {
                    const ratio = Math.min(charsInLine / charsPerLine, 1);
                    xPos = ratio * line.width * this.layout.scale;
                  }
                }
                
                return {
                  pageIndex: page.pageIndex,
                  x: xPos,
                  y: fragment.y + yWithinFragment,
                  height: lineHeight,
                };
              }
            }
          }
        }
        
        // If we couldn't find the exact fragment, return first fragment of block
        for (const page of this.layout.pages) {
          for (const fragment of page.fragments) {
            if (fragment.blockId === block.id) {
              return {
                pageIndex: page.pageIndex,
                x: 0,
                y: fragment.y,
                height: lineHeight,
              };
            }
          }
        }
      }
    }
    
    // Position is past the end of the document - return end of last block
    if (this.blocks.length > 0 && this.layout.pages.length > 0) {
      const lastPage = this.layout.pages[this.layout.pages.length - 1];
      if (lastPage.fragments.length > 0) {
        const lastFragment = lastPage.fragments[lastPage.fragments.length - 1];
        return {
          pageIndex: lastPage.pageIndex,
          x: 0,
          y: lastFragment.y + lastFragment.height - defaultHeight,
          height: defaultHeight,
        };
      }
    }
    
    // Fallback: position at start of first page
    return {
      pageIndex: 0,
      x: 0,
      y: 0,
      height: defaultHeight,
    };
  }
  
  /**
   * Convert a selection range to rectangles using DOM-based mapping
   */
  private selectionToRects(from: number, to: number, _view: EditorView | null): SelectionRect[] {
    if (!this.layout || !this.mount) return [];
    
    const rects: SelectionRect[] = [];
    
    // Use DOM-based mapping for accurate selection rectangles
    const domRects = this.selectionToRectsDom(from, to);
    if (domRects.length > 0) {
      return domRects;
    }
    
    // Fallback: estimate from block data
    const startCaret = this.estimateCaretPosition(from);
    const endCaret = this.estimateCaretPosition(to);
    
    if (startCaret && endCaret) {
      const startPageConfig = getPageConfig(this.layout, startCaret.pageIndex);
      const contentWidth = (startPageConfig.width - 
        startPageConfig.margins.left - 
        startPageConfig.margins.right) * this.layout.scale;
      
      if (startCaret.pageIndex === endCaret.pageIndex) {
        // Same page - use actual x positions for single-line, or full width for multi-line
        const isSingleLine = Math.abs(startCaret.y - endCaret.y) < startCaret.height;
        
        if (isSingleLine) {
          rects.push({
            pageIndex: startCaret.pageIndex,
            x: startCaret.x,
            y: startCaret.y,
            width: Math.max(endCaret.x - startCaret.x, 4),  // Min width 4px
            height: startCaret.height,
          });
        } else {
          // Multi-line on same page
          rects.push({
            pageIndex: startCaret.pageIndex,
            x: startCaret.x,
            y: startCaret.y,
            width: contentWidth - startCaret.x,
            height: startCaret.height,
          });
          
          // Middle lines (full width)
          const middleHeight = endCaret.y - (startCaret.y + startCaret.height);
          if (middleHeight > 0) {
            rects.push({
              pageIndex: startCaret.pageIndex,
              x: 0,
              y: startCaret.y + startCaret.height,
              width: contentWidth,
              height: middleHeight,
            });
          }
          
          // Last line
          rects.push({
            pageIndex: endCaret.pageIndex,
            x: 0,
            y: endCaret.y,
            width: endCaret.x,
            height: endCaret.height,
          });
        }
      } else {
        // Handle multi-page selection
        for (let i = startCaret.pageIndex; i <= endCaret.pageIndex; i++) {
          const page = this.layout.pages[i];
          if (!page) continue;
          
          rects.push({
            pageIndex: i,
            x: 0,
            y: i === startCaret.pageIndex ? startCaret.y : 0,
            width: contentWidth,
            height: i === endCaret.pageIndex 
              ? endCaret.y + endCaret.height 
              : page.contentHeight,
          });
        }
      }
    }
    
    return rects;
  }
  
  /**
   * DOM-based selection rectangle calculation
   * Uses cached span positions for fast lookup
   * 
   * Performance optimizations:
   * - Uses span position cache with binary search
   * - Caches contentRects until scroll happens (scroll version tracking)
   * - Batches DOM reads to minimize layout thrashing
   */
  private selectionToRectsDom(from: number, to: number): SelectionRect[] {
    if (!this.mount || !this.layout) return [];
    
    const rects: SelectionRect[] = [];
    const minPos = Math.min(from, to);
    const maxPos = Math.max(from, to);
    
    // Invalidate cache if scroll has happened since last selection
    if (this.currentScrollVersion !== this.lastScrollVersion) {
      this.contentRectCache.clear();
      this.lastScrollVersion = this.currentScrollVersion;
    }
    
    // Use cached spans with binary search for fast lookup
    const overlappingSpans = this.findOverlappingSpans(minPos, maxPos);
    
    // Early exit for empty selection
    if (overlappingSpans.length === 0) {
      return [];
    }
    
    // Batch read all contentArea rects first (read phase)
    // Uses persistent cache (invalidated on scroll) for performance
    for (const entry of overlappingSpans) {
      const { element: htmlSpan, pageIndex } = entry;
      
      if (!this.contentRectCache.has(pageIndex)) {
        const pageEl = htmlSpan.closest('.page') as HTMLElement | null;
        const contentArea = pageEl?.querySelector('.page-content') as HTMLElement | null;
        if (pageEl && contentArea) {
          this.contentRectCache.set(pageIndex, contentArea.getBoundingClientRect());
        }
      }
    }
    
    // Now process spans using cached contentRects
    for (const entry of overlappingSpans) {
      const { element: htmlSpan, pmStart, pmEnd, pageIndex } = entry;
      
      const contentRect = this.contentRectCache.get(pageIndex);
      if (!contentRect) continue;
      
      const spanRect = htmlSpan.getBoundingClientRect();
      
      // Calculate the overlap range within this span
      const overlapStart = Math.max(minPos, pmStart);
      const overlapEnd = Math.min(maxPos, pmEnd);
      
      // Get text node for Range-based measurement
      const textNode = htmlSpan.firstChild;
      if (!textNode || textNode.nodeType !== Node.TEXT_NODE || !textNode.textContent) {
        // No text - use full span rect
        rects.push({
          pageIndex,
          x: spanRect.left - contentRect.left,
          y: spanRect.top - contentRect.top,
          width: spanRect.width,
          height: spanRect.height,
        });
        continue;
      }
      
      const text = textNode.textContent;
      const charStart = Math.max(0, overlapStart - pmStart);
      const charEnd = Math.min(text.length, overlapEnd - pmStart);
      
      if (charStart >= charEnd) continue;
      
      // For whole-span selection, skip Range API overhead
      if (charStart === 0 && charEnd === text.length) {
        rects.push({
          pageIndex,
          x: spanRect.left - contentRect.left,
          y: spanRect.top - contentRect.top,
          width: spanRect.width,
          height: spanRect.height,
        });
        continue;
      }
      
      // Use Range API to get exact selection rectangle
      try {
        const range = document.createRange();
        range.setStart(textNode, charStart);
        range.setEnd(textNode, charEnd);
        
        const rangeRects = range.getClientRects();
        for (const rangeRect of rangeRects) {
          rects.push({
            pageIndex,
            x: rangeRect.left - contentRect.left,
            y: rangeRect.top - contentRect.top,
            width: rangeRect.width,
            height: rangeRect.height,
          });
        }
      } catch {
        // Fallback to span rect
        rects.push({
          pageIndex,
          x: spanRect.left - contentRect.left,
          y: spanRect.top - contentRect.top,
          width: spanRect.width,
          height: spanRect.height,
        });
      }
    }
    
    // Merge adjacent/overlapping rectangles on same line
    return this.mergeSelectionRects(rects);
  }
  
  /**
   * Merge adjacent selection rectangles for cleaner rendering
   */
  private mergeSelectionRects(rects: SelectionRect[]): SelectionRect[] {
    if (rects.length <= 1) return rects;
    
    // Sort by page, then by y, then by x
    rects.sort((a, b) => {
      if (a.pageIndex !== b.pageIndex) return a.pageIndex - b.pageIndex;
      if (Math.abs(a.y - b.y) > 2) return a.y - b.y;
      return a.x - b.x;
    });
    
    const merged: SelectionRect[] = [];
    let current = rects[0];
    
    for (let i = 1; i < rects.length; i++) {
      const next = rects[i];
      
      // Check if on same page and same line (within 2px tolerance)
      const samePage = current.pageIndex === next.pageIndex;
      const sameLine = Math.abs(current.y - next.y) <= 2;
      const adjacent = next.x <= current.x + current.width + 2;
      
      if (samePage && sameLine && adjacent) {
        // Merge
        const newWidth = Math.max(current.x + current.width, next.x + next.width) - current.x;
        current = {
          ...current,
          width: newWidth,
          height: Math.max(current.height, next.height),
        };
      } else {
        merged.push(current);
        current = next;
      }
    }
    
    merged.push(current);
    return merged;
  }
  
  /**
   * Render the caret
   */
  private renderCaret(pos: CaretPosition): void {
    if (!this.caretElement || !this.layout) return;
    
    // Calculate absolute position
    const origin = getPageContentOrigin(this.layout, pos.pageIndex);
    const absoluteX = origin.x + pos.x;
    const absoluteY = origin.y + pos.y;
    
    // Use transform for GPU-accelerated positioning
    this.caretElement.style.transform = `translate(${absoluteX}px, ${absoluteY}px)`;
    this.caretElement.style.height = `${pos.height}px`;
    
    // Hide caret when cell selection is active, otherwise show it
    if (this.cellSelectionActive) {
      this.caretElement.style.display = 'none';
    } else {
      this.caretElement.style.display = 'block';
    }
    
    // Update blinking based on focus
    if (this.isFocused) {
      this.caretElement.classList.add('blinking');
    } else {
      this.caretElement.classList.remove('blinking');
    }
  }
  
  /**
   * Hide the caret
   */
  private hideCaret(): void {
    if (this.caretElement) {
      this.caretElement.style.display = 'none';
    }
  }
  
  /**
   * Clear selection rectangles (removes from DOM)
   */
  private clearSelectionRects(): void {
    for (const el of this.selectionElements) {
      el.remove();
    }
    this.selectionElements = [];
  }
  
  /**
   * Return selection rects to pool for reuse (performance optimization)
   */
  private returnRectsToPool(): void {
    for (const el of this.selectionElements) {
      el.style.display = 'none';
      this.selectionPool.push(el);
    }
    this.selectionElements = [];
  }
  
  /**
   * Get a rect element from pool or create new one
   */
  private getRectFromPool(): HTMLElement {
    if (this.selectionPool.length > 0) {
      const el = this.selectionPool.pop()!;
      el.style.display = '';
      return el;
    }
    const el = document.createElement('div');
    el.className = 'selection-rect';
    return el;
  }
  
  /**
   * Render selection rectangles using element pooling (performance optimization)
   */
  private renderSelectionRectsPooled(rects: SelectionRect[]): void {
    if (!this.overlayContainer || !this.layout) return;
    
    for (const rect of rects) {
      const el = this.getRectFromPool();
      
      // Calculate absolute position
      const origin = getPageContentOrigin(this.layout, rect.pageIndex);
      const absoluteX = origin.x + rect.x;
      const absoluteY = origin.y + rect.y;
      
      // Use transform for GPU-accelerated positioning
      el.style.transform = `translate(${absoluteX}px, ${absoluteY}px)`;
      el.style.width = `${rect.width}px`;
      el.style.height = `${rect.height}px`;
      
      if (!el.parentElement) {
        this.overlayContainer.appendChild(el);
      }
      this.selectionElements.push(el);
    }
  }
  
  /**
   * Clean up
   */
  destroy(): void {
    if (this.animationFrame !== null) {
      cancelAnimationFrame(this.animationFrame);
    }
    
    // Clean up scroll listener
    if (this.scrollContainer && this.scrollHandler) {
      this.scrollContainer.removeEventListener('scroll', this.scrollHandler);
      this.scrollHandler = null;
      this.scrollContainer = null;
    }
    
    this.clearSelectionRects();
    
    // Clean up pool
    for (const el of this.selectionPool) {
      el.remove();
    }
    this.selectionPool = [];
    
    if (this.caretElement) {
      this.caretElement.remove();
      this.caretElement = null;
    }
    
    this.overlayContainer = null;
  }
}

// ============================================================================
// React Component Helper
// ============================================================================

/**
 * Create selection overlay styles for inline use
 */
export function getSelectionOverlayStyles(): string {
  return `
    @keyframes caret-blink {
      0%, 100% { opacity: 1; }
      50% { opacity: 0; }
    }
    .selection-overlay {
      position: absolute;
      top: 0;
      left: 0;
      right: 0;
      bottom: 0;
      pointer-events: none;
      z-index: 100;
    }
    .selection-caret {
      position: absolute;
      width: 2px;
      background-color: #000;
      pointer-events: none;
    }
    .selection-caret.blinking {
      animation: caret-blink 1s step-end infinite;
    }
    .selection-rect {
      position: absolute;
      top: 0;
      left: 0;
      background-color: rgba(66, 133, 244, 0.3);
      pointer-events: none;
      will-change: transform;
    }
  `;
}

