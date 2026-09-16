/**
 * DOM Painter - Renders pages independently based on layout
 * 
 * This is the key component that makes the true layout engine work.
 * Instead of clipping a single continuous editor, we render each page's
 * content independently, showing only the fragments assigned to that page.
 * 
 * Key insight: Each page gets its own DOM tree with only the content
 * that belongs to that page. No clipping needed!
 */

import {
  FlowBlock,
  Run,
  ParagraphBlock,
  HeadingBlock,
  ListItemBlock,
  TableBlock,
  ImageBlock,
  isRunContainingBlock,
} from './flow-blocks';
import { Measure, hasLineData, LineMeasure } from './measurer';
import { DocumentLayout, PageLayout, PageFragment, PageConfig } from './true-layout-engine';
import {
  sanitizeHref,
  sanitizeImageSrc,
  safeCssColor,
  safeFontFamily,
  safeCssKeyword,
  safeLineHeight,
} from './sanitize';

// Re-export header/footer types from docs-core for convenience
export type {
  DynamicFieldType,
  DynamicFieldRun,
  HeaderFooterTextRun,
  HeaderFooterImageRun,
  HeaderFooterInlineContent,
  HeaderFooterParagraph,
  HeaderFooterContent,
} from '@b32nio/spindle-docs-core';

// Import types for local use
import type {
  DynamicFieldRun,
  HeaderFooterParagraph,
  HeaderFooterContent,
} from '@b32nio/spindle-docs-core';

/**
 * Context for resolving dynamic fields
 */
export interface DynamicFieldContext {
  /** Current page number (1-based) */
  pageNumber: number;
  /** Total number of pages */
  totalPages: number;
  /** Document title */
  title?: string;
  /** Current date */
  date?: Date;
}

/**
 * Callback when header/footer area is clicked
 */
export type HeaderFooterClickHandler = (
  type: 'header' | 'footer',
  pageIndex: number
) => void;

/**
 * Configuration for the DOM painter
 */
export interface PainterConfig {
  /** Page configuration */
  pageConfig: PageConfig;
  /** Scale factor (zoom) */
  scale: number;
  /** Default font family */
  fontFamily?: string;
  /** Default font size in pixels */
  fontSize?: number;
  /** Default line height multiplier */
  lineHeight?: number;
  /** Gap between pages */
  pageGap?: number;
  /** Header content configuration */
  header?: HeaderFooterContent;
  /** Footer content configuration */
  footer?: HeaderFooterContent;
  /** Document title (for dynamic field resolution) */
  documentTitle?: string;
  /** Callback when header/footer is clicked (for editing) */
  onHeaderFooterClick?: HeaderFooterClickHandler;
  /** Whether header/footer areas should be interactive */
  headerFooterEditable?: boolean;
}

/**
 * Rendered page information
 */
export interface RenderedPage {
  /** The page element */
  element: HTMLElement;
  /** Page index */
  pageIndex: number;
  /** Content area element */
  contentArea: HTMLElement;
}

// ============================================================================
// DOM Painter Class
// ============================================================================

/**
 * PM position info for a block
 */
export interface BlockPmPosition {
  start: number;
  end: number;
}

/**
 * Virtualization configuration
 */
export interface VirtualizationConfig {
  /** Enable virtualization (default: true) */
  enabled?: boolean;
  /** Number of pages to render at once (default: 5) */
  windowSize?: number;
  /** Number of extra pages to render outside visible area (default: 1) */
  overscan?: number;
  /** Scroll container element for detecting visible pages */
  scrollContainer?: HTMLElement | null;
  /** Callback when visible pages change (for invalidating caches) */
  onPagesChange?: () => void;
}

/**
 * Internal configuration with resolved defaults
 */
interface ResolvedPainterConfig {
  pageConfig: PageConfig;
  scale: number;
  fontFamily: string;
  fontSize: number;
  lineHeight: number;
  pageGap: number;
  header?: HeaderFooterContent;
  footer?: HeaderFooterContent;
  documentTitle?: string;
  onHeaderFooterClick?: HeaderFooterClickHandler;
  headerFooterEditable?: boolean;
}

/**
 * DomPainter - Renders document layout to DOM with optional page virtualization
 * 
 * When virtualization is enabled, only visible pages + a small buffer are rendered
 * to the DOM, dramatically improving scroll performance for large documents.
 */
export class DomPainter {
  private config: ResolvedPainterConfig;
  private blockLookup: Map<string, { block: FlowBlock; measure: Measure; index: number; pmPos?: BlockPmPosition }> = new Map();
  
  // Virtualization state
  private virtualEnabled = true;
  private virtualWindowSize = 5;
  private virtualOverscan = 1;
  private scrollContainer: HTMLElement | null = null;
  private currentLayout: DocumentLayout | null = null;
  private stackContainer: HTMLElement | null = null;
  private mountedPages: Map<number, RenderedPage> = new Map();
  private visibleStart = 0;
  private visibleEnd = 0;
  private onScrollHandler: (() => void) | null = null;
  private scrollRafId: number | null = null;
  private onPagesChangeCallback: (() => void) | null = null;
  
  constructor(config: PainterConfig) {
    this.config = {
      pageConfig: config.pageConfig,
      scale: config.scale,
      fontFamily: config.fontFamily ?? 'Arial, sans-serif',
      fontSize: config.fontSize ?? 11,
      lineHeight: config.lineHeight ?? 1.5,
      pageGap: config.pageGap ?? 24,
      header: config.header,
      footer: config.footer,
      documentTitle: config.documentTitle,
      onHeaderFooterClick: config.onHeaderFooterClick,
      headerFooterEditable: config.headerFooterEditable,
    };
  }
  
  /**
   * Update the painter configuration
   */
  updateConfig(config: Partial<PainterConfig>): void {
    Object.assign(this.config, config);
  }
  
  /**
   * Configure virtualization settings
   */
  setVirtualization(config: VirtualizationConfig): void {
    this.virtualEnabled = config.enabled ?? true;
    this.virtualWindowSize = Math.max(1, config.windowSize ?? 5);
    this.virtualOverscan = Math.max(0, config.overscan ?? 1);
    this.onPagesChangeCallback = config.onPagesChange ?? null;
    
    // Handle scroll container change
    if (config.scrollContainer !== undefined && config.scrollContainer !== this.scrollContainer) {
      this.unbindScrollHandler();
      this.scrollContainer = config.scrollContainer;
      if (this.scrollContainer && this.virtualEnabled) {
        this.bindScrollHandler();
      }
    }
  }
  
  /**
   * Bind scroll handler for virtualization
   */
  private bindScrollHandler(): void {
    if (!this.scrollContainer || this.onScrollHandler) return;
    
    this.onScrollHandler = () => {
      // Use RAF to throttle updates during scroll
      if (this.scrollRafId !== null) return;
      
      this.scrollRafId = requestAnimationFrame(() => {
        this.scrollRafId = null;
        this.updateVisiblePages();
      });
    };
    
    this.scrollContainer.addEventListener('scroll', this.onScrollHandler, { passive: true });
  }
  
  /**
   * Unbind scroll handler
   */
  private unbindScrollHandler(): void {
    if (this.scrollContainer && this.onScrollHandler) {
      this.scrollContainer.removeEventListener('scroll', this.onScrollHandler);
    }
    this.onScrollHandler = null;
    if (this.scrollRafId !== null) {
      cancelAnimationFrame(this.scrollRafId);
      this.scrollRafId = null;
    }
  }
  
  /**
   * Calculate which pages are currently visible
   */
  private getVisibleRange(): { start: number; end: number } {
    if (!this.scrollContainer || !this.currentLayout || this.currentLayout.pages.length === 0) {
      return { start: 0, end: Math.min(this.virtualWindowSize - 1, (this.currentLayout?.pages.length ?? 1) - 1) };
    }
    
    const pageHeight = this.getPageHeight();
    const gap = this.config.pageGap * this.config.scale;
    const pageWithGap = pageHeight + gap;
    
    // Get scroll position relative to content
    const scrollTop = this.scrollContainer.scrollTop;
    const viewportHeight = this.scrollContainer.clientHeight;
    
    // Calculate visible page range
    const firstVisible = Math.max(0, Math.floor(scrollTop / pageWithGap));
    const lastVisible = Math.min(
      this.currentLayout.pages.length - 1,
      Math.ceil((scrollTop + viewportHeight) / pageWithGap)
    );
    
    // Add overscan
    const start = Math.max(0, firstVisible - this.virtualOverscan);
    const end = Math.min(this.currentLayout.pages.length - 1, lastVisible + this.virtualOverscan);
    
    return { start, end };
  }
  
  /**
   * Update which pages are rendered based on scroll position
   */
  private updateVisiblePages(): void {
    if (!this.virtualEnabled || !this.currentLayout || !this.stackContainer) return;
    
    const { start, end } = this.getVisibleRange();
    
    // Skip if range hasn't changed
    if (start === this.visibleStart && end === this.visibleEnd) return;
    
    this.visibleStart = start;
    this.visibleEnd = end;
    
    const neededPages = new Set<number>();
    for (let i = start; i <= end; i++) {
      neededPages.add(i);
    }
    
    let pagesChanged = false;
    
    // Remove pages that are no longer needed
    for (const [pageIndex, page] of this.mountedPages.entries()) {
      if (!neededPages.has(pageIndex)) {
        page.element.remove();
        this.mountedPages.delete(pageIndex);
        pagesChanged = true;
      }
    }
    
    // Add pages that are now needed
    for (const pageIndex of neededPages) {
      if (!this.mountedPages.has(pageIndex)) {
        const pageLayout = this.currentLayout.pages[pageIndex];
        if (pageLayout) {
          const rendered = this.paintPage(pageLayout, this.currentLayout);
          this.mountedPages.set(pageIndex, rendered);
          this.stackContainer.appendChild(rendered.element);
          pagesChanged = true;
        }
      }
    }
    
    // Notify when pages change (for cache invalidation)
    if (pagesChanged && this.onPagesChangeCallback) {
      this.onPagesChangeCallback();
    }
  }
  
  /**
   * Set the data to paint
   * @param blocks - FlowBlocks to render
   * @param measures - Measurements for each block
   * @param pmPositions - Optional map of block IDs to PM positions
   */
  setData(blocks: FlowBlock[], measures: Measure[], pmPositions?: Map<string, BlockPmPosition>): void {
    
    // Build lookup map
    this.blockLookup.clear();
    for (let i = 0; i < blocks.length; i++) {
      const block = blocks[i];
      this.blockLookup.set(block.id, {
        block,
        measure: measures[i],
        index: i,
        pmPos: pmPositions?.get(block.id),
      });
    }
  }
  
  /**
   * Paint the entire document layout
   * With virtualization enabled, only visible pages are rendered initially
   */
  paint(layout: DocumentLayout, mount: HTMLElement): RenderedPage[] {
    this.currentLayout = layout;
    
    // Clear mount
    mount.innerHTML = '';
    this.mountedPages.clear();
    
    // Create page stack container with full height for proper scrolling
    this.stackContainer = document.createElement('div');
    this.stackContainer.className = 'page-stack';
    this.stackContainer.style.cssText = `
      position: relative;
      width: ${this.getPageWidth()}px;
      height: ${layout.totalHeight}px;
    `;
    
    mount.appendChild(this.stackContainer);
    
    // Calculate visible range
    const { start, end } = this.getVisibleRange();
    this.visibleStart = start;
    this.visibleEnd = end;
    
    const renderedPages: RenderedPage[] = [];
    
    if (this.virtualEnabled) {
      // Render only visible pages
      for (let i = start; i <= end && i < layout.pages.length; i++) {
        const rendered = this.paintPage(layout.pages[i], layout);
        renderedPages.push(rendered);
        this.mountedPages.set(i, rendered);
        this.stackContainer.appendChild(rendered.element);
      }
      
      // Bind scroll handler if we have a scroll container
      if (this.scrollContainer) {
        this.bindScrollHandler();
      }
    } else {
      // Render all pages (old behavior)
      for (const page of layout.pages) {
        const rendered = this.paintPage(page, layout);
        renderedPages.push(rendered);
        this.mountedPages.set(page.pageIndex, rendered);
        this.stackContainer.appendChild(rendered.element);
      }
    }
    
    return renderedPages;
  }
  
  /**
   * Cleanup resources
   */
  destroy(): void {
    this.unbindScrollHandler();
    this.mountedPages.clear();
    this.currentLayout = null;
    this.stackContainer = null;
  }
  
  /**
   * Paint a single page
   */
  paintPage(page: PageLayout, _layout: DocumentLayout): RenderedPage {
    const pageWidth = this.getPageWidth();
    const pageHeight = this.getPageHeight();
    const margins = this.getScaledMargins();
    const contentWidth = pageWidth - margins.left - margins.right;
    
    // Calculate page Y position
    const pageY = page.pageIndex * (pageHeight + this.config.pageGap * this.config.scale);
    
    // Create page container
    const pageEl = document.createElement('div');
    pageEl.className = 'page';
    pageEl.dataset.pageIndex = String(page.pageIndex);
    pageEl.style.cssText = `
      position: absolute;
      left: 0;
      top: ${pageY}px;
      width: ${pageWidth}px;
      height: ${pageHeight}px;
      background: white;
      overflow: hidden;
      contain: strict;
      transform: translateZ(0);
    `;
    
    // Add shadow via pseudo-element for better performance (avoids repaint on scroll)
    // We'll add a separate shadow element instead of box-shadow on main element
    const shadowEl = document.createElement('div');
    shadowEl.className = 'page-shadow';
    shadowEl.style.cssText = `
      position: absolute;
      inset: 0;
      box-shadow: 0 1px 3px rgba(0, 0, 0, 0.12), 0 1px 2px rgba(0, 0, 0, 0.24);
      pointer-events: none;
      z-index: -1;
    `;
    pageEl.appendChild(shadowEl);
    
    // Create content area
    const contentArea = document.createElement('div');
    contentArea.className = 'page-content';
    contentArea.style.cssText = `
      position: absolute;
      top: ${margins.top}px;
      left: ${margins.left}px;
      width: ${contentWidth}px;
      height: ${pageHeight - margins.top - margins.bottom}px;
      font-family: ${this.config.fontFamily};
      font-size: ${this.config.fontSize * this.config.scale}px;
      line-height: ${this.config.lineHeight};
      overflow: hidden;
    `;
    
    // Render each fragment
    for (const fragment of page.fragments) {
      const fragmentEl = this.renderFragment(fragment);
      if (fragmentEl) {
        contentArea.appendChild(fragmentEl);
      }
    }
    
    pageEl.appendChild(contentArea);
    
    // Create dynamic field context for header/footer
    const fieldContext: DynamicFieldContext = {
      pageNumber: page.pageIndex + 1,
      totalPages: _layout.pages.length,
      title: this.config.documentTitle,
      date: new Date(),
    };
    
    // Get header/footer margins (default to 48px / 0.5 inch if not specified)
    const headerMargin = (this.config.pageConfig.margins.header ?? 48) * this.config.scale;
    const footerMargin = (this.config.pageConfig.margins.footer ?? 48) * this.config.scale;
    
    // Render header if configured or editable
    const isEditable = this.config.headerFooterEditable ?? false;
    const emptyContent: HeaderFooterContent = { blocks: [] };
    
    if (this.config.header || isEditable) {
      const headerEl = this.renderHeaderFooter(
        'header',
        this.config.header ?? emptyContent,
        page.pageIndex,
        fieldContext,
        contentWidth,
        headerMargin,
        margins
      );
      if (headerEl) {
        pageEl.appendChild(headerEl);
      }
    }
    
    // Render footer if configured or editable
    if (this.config.footer || isEditable) {
      const footerEl = this.renderHeaderFooter(
        'footer',
        this.config.footer ?? emptyContent,
        page.pageIndex,
        fieldContext,
        contentWidth,
        footerMargin,
        margins
      );
      if (footerEl) {
        pageEl.appendChild(footerEl);
      }
    } else {
      // Fallback: Add simple page number if no footer is configured and not editable
      const pageNumber = document.createElement('div');
      pageNumber.className = 'page-number';
      pageNumber.textContent = String(page.pageIndex + 1);
      pageNumber.style.cssText = `
        position: absolute;
        bottom: ${Math.max(8, margins.bottom / 2 - 8)}px;
        left: 0;
        right: 0;
        text-align: center;
        font-size: ${10 * this.config.scale}px;
        color: #5f6368;
        pointer-events: none;
      `;
      pageEl.appendChild(pageNumber);
    }
    
    return {
      element: pageEl,
      pageIndex: page.pageIndex,
      contentArea,
    };
  }
  
  /**
   * Render a header or footer area
   */
  private renderHeaderFooter(
    type: 'header' | 'footer',
    content: HeaderFooterContent,
    pageIndex: number,
    fieldContext: DynamicFieldContext,
    contentWidth: number,
    areaMargin: number,
    margins: { top: number; bottom: number; left: number; right: number }
  ): HTMLElement | null {
    const isFirstPage = pageIndex === 0;
    const isEditable = this.config.headerFooterEditable ?? false;
    
    // Determine which blocks to use based on first page setting
    let blocks = content.blocks;
    if (content.differentFirstPage && isFirstPage) {
      blocks = content.firstPageBlocks ?? [];
    }
    
    // Create the header/footer container (even if empty, for editable click target)
    const container = document.createElement('div');
    container.className = `page-${type}`;
    container.dataset.hfType = type;
    container.dataset.pageIndex = String(pageIndex);
    
    // Base styles
    const baseStyles = `
      font-family: ${this.config.fontFamily};
      font-size: ${(this.config.fontSize - 1) * this.config.scale}px;
      line-height: ${this.config.lineHeight};
      color: #5f6368;
      min-height: ${20 * this.config.scale}px;
    `;
    
    // Interactive styles when editable
    const interactiveStyles = isEditable ? `
      pointer-events: auto;
      cursor: pointer;
      transition: background-color 0.15s;
    ` : `
      pointer-events: none;
    `;
    
    // Position the header or footer
    if (type === 'header') {
      container.style.cssText = `
        position: absolute;
        top: ${areaMargin}px;
        left: ${margins.left}px;
        width: ${contentWidth}px;
        ${baseStyles}
        ${interactiveStyles}
      `;
    } else {
      container.style.cssText = `
        position: absolute;
        bottom: ${areaMargin}px;
        left: ${margins.left}px;
        width: ${contentWidth}px;
        ${baseStyles}
        ${interactiveStyles}
      `;
    }
    
    // Add hover effect and click handler when editable
    if (isEditable) {
      container.addEventListener('mouseenter', () => {
        container.style.backgroundColor = 'rgba(26, 115, 232, 0.08)';
      });
      container.addEventListener('mouseleave', () => {
        container.style.backgroundColor = 'transparent';
      });
      container.addEventListener('click', (e) => {
        e.stopPropagation();
        this.config.onHeaderFooterClick?.(type, pageIndex);
      });
    }
    
    // If no blocks and editable, show placeholder
    if (blocks.length === 0) {
      if (isEditable) {
        const placeholder = document.createElement('div');
        placeholder.className = 'hf-placeholder';
        placeholder.textContent = `Click to add ${type}`;
        placeholder.style.cssText = `
          color: #9aa0a6;
          font-style: italic;
          text-align: center;
          padding: 4px;
        `;
        container.appendChild(placeholder);
        return container;
      }
      return null;
    }
    
    // Render each paragraph block
    for (const block of blocks) {
      const paraEl = this.renderHeaderFooterParagraph(block, fieldContext);
      container.appendChild(paraEl);
    }
    
    return container;
  }
  
  /**
   * Render a header/footer paragraph
   */
  private renderHeaderFooterParagraph(
    block: HeaderFooterParagraph,
    fieldContext: DynamicFieldContext
  ): HTMLElement {
    const para = document.createElement('p');
    para.className = 'hf-paragraph';
    para.style.cssText = `
      margin: 0;
      padding: 0;
      text-align: ${safeCssKeyword(block.alignment) || 'center'};
    `;

    // Render inline content
    for (const item of block.content) {
      if (item.type === 'text') {
        const span = document.createElement('span');
        span.textContent = item.text;
        
        let styles = '';
        if (item.bold) styles += 'font-weight: bold;';
        if (item.italic) styles += 'font-style: italic;';
        if (item.fontSize) styles += `font-size: ${item.fontSize * this.config.scale}px;`;
        const itemFont = safeFontFamily(item.fontFamily);
        if (itemFont) styles += `font-family: ${itemFont};`;
        const itemColor = safeCssColor(item.color);
        if (itemColor) styles += `color: ${itemColor};`;

        span.style.cssText = styles;
        para.appendChild(span);
      } else if (item.type === 'dynamicField') {
        const span = document.createElement('span');
        span.className = 'dynamic-field';
        span.textContent = this.resolveDynamicField(item, fieldContext);
        para.appendChild(span);
      } else if (item.type === 'image') {
        const img = document.createElement('img');
        img.src = sanitizeImageSrc(item.src);
        img.alt = item.alt || '';
        img.style.cssText = `
          width: ${(item.width || 24) * this.config.scale}px;
          height: ${(item.height || 24) * this.config.scale}px;
          object-fit: contain;
          vertical-align: middle;
        `;
        para.appendChild(img);
      }
    }
    
    return para;
  }
  
  /**
   * Resolve a dynamic field to its actual value
   */
  private resolveDynamicField(
    field: DynamicFieldRun,
    context: DynamicFieldContext
  ): string {
    switch (field.fieldType) {
      case 'pageNumber':
        return String(context.pageNumber);
      
      case 'totalPages':
        return String(context.totalPages);
      
      case 'date': {
        const date = context.date || new Date();
        // Support basic format strings, default to locale date
        if (field.format) {
          return this.formatDate(date, field.format);
        }
        return date.toLocaleDateString();
      }
      
      case 'time': {
        const date = context.date || new Date();
        if (field.format) {
          return this.formatTime(date, field.format);
        }
        return date.toLocaleTimeString();
      }
      
      case 'title':
        return context.title || 'Untitled';
      
      default:
        return '';
    }
  }
  
  /**
   * Format a date using a format string
   */
  private formatDate(date: Date, format: string): string {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    
    return format
      .replace('YYYY', String(year))
      .replace('YY', String(year).slice(-2))
      .replace('MM', month)
      .replace('M', String(date.getMonth() + 1))
      .replace('DD', day)
      .replace('D', String(date.getDate()));
  }
  
  /**
   * Format a time using a format string
   */
  private formatTime(date: Date, format: string): string {
    const hours24 = date.getHours();
    const hours12 = hours24 % 12 || 12;
    const minutes = String(date.getMinutes()).padStart(2, '0');
    const seconds = String(date.getSeconds()).padStart(2, '0');
    const ampm = hours24 >= 12 ? 'PM' : 'AM';
    
    return format
      .replace('HH', String(hours24).padStart(2, '0'))
      .replace('H', String(hours24))
      .replace('hh', String(hours12).padStart(2, '0'))
      .replace('h', String(hours12))
      .replace('mm', minutes)
      .replace('ss', seconds)
      .replace('A', ampm)
      .replace('a', ampm.toLowerCase());
  }
  
  /**
   * Render a single fragment
   */
  private renderFragment(fragment: PageFragment): HTMLElement | null {
    const data = this.blockLookup.get(fragment.blockId);
    if (!data) return null;
    
    const { block, measure, pmPos } = data;
    
    // Create fragment container
    const fragmentEl = document.createElement('div');
    fragmentEl.className = 'fragment';
    fragmentEl.dataset.blockId = fragment.blockId;
    fragmentEl.dataset.fromLine = String(fragment.fromLine);
    fragmentEl.dataset.toLine = String(fragment.toLine);
    
    // Add PM position data attributes for click-to-position mapping
    if (pmPos) {
      fragmentEl.dataset.pmStart = String(pmPos.start);
      fragmentEl.dataset.pmEnd = String(pmPos.end);
    }
    
    fragmentEl.style.cssText = `
      position: absolute;
      top: ${fragment.y}px;
      left: ${fragment.x}px;
      width: 100%;
      overflow: hidden;
    `;
    
    // Render content based on block type
    switch (block.kind) {
      case 'paragraph':
        this.renderParagraphFragment(fragmentEl, block, measure, fragment);
        break;
      case 'heading':
        this.renderHeadingFragment(fragmentEl, block, measure, fragment);
        break;
      case 'listItem':
        this.renderListItemFragment(fragmentEl, block, measure, fragment);
        break;
      case 'table':
        this.renderTableFragment(fragmentEl, block as TableBlock, fragment);
        break;
      case 'image':
        this.renderImageFragment(fragmentEl, block as ImageBlock, fragment);
        break;
      case 'horizontalRule':
        this.renderHorizontalRule(fragmentEl);
        break;
    }
    
    return fragmentEl;
  }
  
  /**
   * Render a paragraph fragment
   * 
   * This is where the magic happens for line-level pagination.
   * We render ONLY the text content for the lines in this fragment,
   * not the full paragraph with clipping.
   */
  private renderParagraphFragment(
    container: HTMLElement,
    block: ParagraphBlock,
    measure: Measure,
    fragment: PageFragment
  ): void {
    const data = this.blockLookup.get(block.id);
    const pmPos = data?.pmPos;
    
    const para = document.createElement('p');
    para.className = 'paragraph-fragment';
    
    // Get paragraph styles
    const attrs = block.attrs || {};
    // Indent level (0, 1, 2, etc.) converts to pixels (each level = 24px base)
    const indentLevel = attrs.indent || 0;
    const indentPx = indentLevel * 24 * this.config.scale;
    para.style.cssText = `
      margin: 0;
      padding: 0;
      text-align: ${safeCssKeyword(attrs.alignment) || 'left'};
      line-height: ${safeLineHeight(attrs.lineHeight) ?? this.config.lineHeight};
      padding-left: ${indentPx + (attrs.leftIndent || 0) * this.config.scale}px;
      padding-right: ${(attrs.rightIndent || 0) * this.config.scale}px;
    `;
    
    // Only apply first-line indent if this is the first fragment
    if (fragment.isFirstFragment && attrs.firstLineIndent) {
      para.style.textIndent = `${attrs.firstLineIndent * this.config.scale}px`;
    }
    
    // Add space before if this is the first fragment
    if (fragment.isFirstFragment && attrs.spaceBefore) {
      para.style.marginTop = `${attrs.spaceBefore * this.config.scale}px`;
    }
    
    // Add space after if this is the last fragment
    if (fragment.isLastFragment && attrs.spaceAfter) {
      para.style.marginBottom = `${(attrs.spaceAfter || 8) * this.config.scale}px`;
    }
    
    // Check if we have line segment data for precise rendering
    if (hasLineData(measure) && measure.lines.length > 0) {
      const lines = measure.lines.slice(fragment.fromLine, fragment.toLine);
      const hasSegments = lines.some(line => line.segments && line.segments.length > 0);
      
      if (hasSegments) {
        // Render only the text for these specific lines with PM position tracking
        this.renderLinesContentWithPmPositions(para, block.runs, lines, measure.lines, fragment.fromLine, pmPos);
      } else {
        // Fallback: render all runs with PM positions (will be clipped)
        this.renderRuns(para, block.runs, pmPos);
        this.applyFragmentClipping(para, container, measure, fragment);
        return;
      }
    } else {
      // No line data - render full content with PM positions
      this.renderRuns(para, block.runs, pmPos);
    }
    
    container.appendChild(para);
  }
  
  /**
   * Render lines content with PM position data attributes for click-to-position mapping.
   * Embeds data-pm-start/data-pm-end attributes in spans for accurate cursor positioning.
   */
  private renderLinesContentWithPmPositions(
    container: HTMLElement,
    runs: Run[],
    fragmentLines: LineMeasure[],
    allLines: LineMeasure[],
    fromLineIndex: number,
    pmPos?: BlockPmPosition
  ): void {
    if (fragmentLines.length === 0) {
      container.innerHTML = '&#8203;';
      return;
    }
    
    // Calculate the character offset at the start of this fragment
    // by counting characters in lines before fromLineIndex
    let charOffsetAtFragmentStart = 0;
    for (let i = 0; i < fromLineIndex && i < allLines.length; i++) {
      const line = allLines[i];
      if (line.segments) {
        for (const seg of line.segments) {
          charOffsetAtFragmentStart += seg.text.length;
        }
      }
    }
    
    // Track current character offset as we render
    let currentCharOffset = charOffsetAtFragmentStart;
    
    // Render each line with a wrapper div containing PM positions
    for (let lineIdx = 0; lineIdx < fragmentLines.length; lineIdx++) {
      const line = fragmentLines[lineIdx];
      if (!line.segments || line.segments.length === 0) continue;
      
      // Calculate PM positions for this line
      const lineStartChar = currentCharOffset;
      let lineEndChar = lineStartChar;
      for (const seg of line.segments) {
        lineEndChar += seg.text.length;
      }
      
      // Create line wrapper with PM positions
      const lineDiv = document.createElement('div');
      lineDiv.className = 'line';
      
      if (pmPos) {
        // PM positions: block start + 1 (enter node) + char offset
        const linepmStart = pmPos.start + 1 + lineStartChar;
        const linepmEnd = pmPos.start + 1 + lineEndChar;
        lineDiv.dataset.pmStart = String(linepmStart);
        lineDiv.dataset.pmEnd = String(linepmEnd);
      }
      
      // Render spans within this line
      let spanCharOffset = lineStartChar;
      for (const segment of line.segments) {
        const run = runs[segment.runIndex];
        if (!run || segment.text.length === 0) continue;
        
        const spanStart = spanCharOffset;
        const spanEnd = spanCharOffset + segment.text.length;
        spanCharOffset = spanEnd;
        
        if (run.kind === 'text') {
          const span = document.createElement('span');
          span.textContent = segment.text;
          span.style.cssText = this.getRunStyles(run);
          
          // Add PM position data attributes
          if (pmPos) {
            span.dataset.pmStart = String(pmPos.start + 1 + spanStart);
            span.dataset.pmEnd = String(pmPos.start + 1 + spanEnd);
          }

          this.applyCommentDecoration(span, run);
          lineDiv.appendChild(span);
        } else if (run.kind === 'link') {
          const a = document.createElement('a');
          a.textContent = segment.text;
          a.href = sanitizeHref(run.href);
          a.style.cssText = `
            color: ${safeCssColor(run.color) || '#1a73e8'};
            text-decoration: underline;
            ${run.bold ? 'font-weight: bold;' : ''}
            ${run.italic ? 'font-style: italic;' : ''}
          `;
          
          // Add PM position data attributes
          if (pmPos) {
            a.dataset.pmStart = String(pmPos.start + 1 + spanStart);
            a.dataset.pmEnd = String(pmPos.start + 1 + spanEnd);
          }

          this.applyCommentDecoration(a, run);
          lineDiv.appendChild(a);
        }
      }
      
      // Update character offset for next line
      currentCharOffset = lineEndChar;
      
      // Add line to container if it has content
      if (lineDiv.childNodes.length > 0) {
        container.appendChild(lineDiv);
      }
    }
    
    // If container is empty, add zero-width space
    if (container.childNodes.length === 0) {
      container.innerHTML = '&#8203;';
    }
  }
  
  /**
   * Apply clipping for fragments when segment data is not available
   * This is a fallback that uses the old negative-margin approach
   */
  private applyFragmentClipping(
    element: HTMLElement,
    container: HTMLElement,
    measure: Measure,
    fragment: PageFragment
  ): void {
    if (fragment.isFirstFragment && fragment.isLastFragment) {
      container.appendChild(element);
      return;
    }
    
    // Apply line clipping using a wrapper
    const wrapper = document.createElement('div');
    wrapper.style.cssText = `
      overflow: hidden;
      height: ${fragment.height}px;
    `;
    
    // If not the first fragment, use negative margin to "scroll" to the right lines
    if (!fragment.isFirstFragment && hasLineData(measure)) {
      let offset = 0;
      for (let i = 0; i < fragment.fromLine; i++) {
        offset += measure.lines[i].height * this.config.scale;
      }
      element.style.marginTop = `-${offset}px`;
    }
    
    wrapper.appendChild(element);
    container.appendChild(wrapper);
  }
  
  /**
   * Render a heading fragment
   */
  private renderHeadingFragment(
    container: HTMLElement,
    block: HeadingBlock,
    measure: Measure,
    fragment: PageFragment
  ): void {
    const data = this.blockLookup.get(block.id);
    const pmPos = data?.pmPos;
    
    // Clamp to a valid h1-h6 tag — an untrusted level (string / out of range)
    // would make createElement throw InvalidCharacterError and abort the render.
    const headingLevel = Math.min(6, Math.max(1, Math.trunc(Number(block.level)) || 1));
    const heading = document.createElement(`h${headingLevel}`);
    heading.className = 'heading-fragment';
    
    // Get heading styles
    const attrs = block.attrs || {};
    const fontSizes: Record<number, number> = {
      1: 24,
      2: 20,
      3: 16,
      4: 14,
      5: 12,
      6: 11,
    };
    const fontSize = (fontSizes[block.level] || 16) * this.config.scale;
    
    // Indent level (0, 1, 2, etc.) converts to pixels (each level = 24px base)
    const indentLevel = attrs.indent || 0;
    const indentPx = indentLevel * 24 * this.config.scale;
    
    heading.style.cssText = `
      margin: 0;
      padding: 0;
      padding-left: ${indentPx}px;
      font-size: ${fontSize}px;
      font-weight: ${block.level <= 2 ? 400 : 700};
      text-align: ${safeCssKeyword(attrs.alignment) || 'left'};
      line-height: 1.3;
    `;
    
    // Add spacing
    if (fragment.isFirstFragment) {
      heading.style.marginTop = `${(attrs.spaceBefore || 16) * this.config.scale}px`;
    }
    if (fragment.isLastFragment) {
      heading.style.marginBottom = `${(attrs.spaceAfter || 8) * this.config.scale}px`;
    }
    
    // Check if we have line segment data for precise rendering
    if (hasLineData(measure) && measure.lines.length > 0) {
      const lines = measure.lines.slice(fragment.fromLine, fragment.toLine);
      const hasSegments = lines.some(line => line.segments && line.segments.length > 0);
      
      if (hasSegments) {
        // Use PM position-aware rendering
        this.renderLinesContentWithPmPositions(heading, block.runs, lines, measure.lines, fragment.fromLine, pmPos);
        container.appendChild(heading);
        return;
      }
    }
    
    // Fallback: render all runs with PM positions and clipping
    this.renderRuns(heading, block.runs, pmPos);
    this.applyFragmentClipping(heading, container, measure, fragment);
  }
  
  /**
   * Render a list item fragment
   */
  private renderListItemFragment(
    container: HTMLElement,
    block: ListItemBlock,
    measure: Measure,
    fragment: PageFragment
  ): void {
    const data = this.blockLookup.get(block.id);
    // For list items, pmPos points to list_item node. Content is at +1 (paragraph inside list_item).
    // renderRuns adds another +1, so we need to add +1 here to make total offset +2.
    const pmPos = data?.pmPos ? { start: data.pmPos.start + 1, end: data.pmPos.end } : undefined;
    
    // Create list structure
    const list = document.createElement(block.listType === 'bullet' ? 'ul' : 'ol') as HTMLOListElement | HTMLUListElement;
    const item = document.createElement('li');
    
    // For ordered lists, set the start attribute to show correct number
    if (block.listType === 'ordered' && block.listIndex !== undefined) {
      (list as HTMLOListElement).start = block.listIndex;
    }
    
    const indent = 24 * (block.level + 1) * this.config.scale;
    list.style.cssText = `
      margin: 0;
      padding-left: ${indent}px;
      list-style-position: outside;
    `;
    
    const attrs = block.attrs || {};
    item.style.cssText = `
      margin: 0;
      text-align: ${safeCssKeyword(attrs.alignment) || 'left'};
      line-height: ${safeLineHeight(attrs.lineHeight) ?? this.config.lineHeight};
    `;
    
    // Add spacing
    if (fragment.isFirstFragment) {
      list.style.marginTop = `${(attrs.spaceBefore || 0) * this.config.scale}px`;
    }
    if (fragment.isLastFragment) {
      list.style.marginBottom = `${(attrs.spaceAfter || 4) * this.config.scale}px`;
    }
    
    // Check if we have line segment data for precise rendering
    if (hasLineData(measure) && measure.lines.length > 0) {
      const lines = measure.lines.slice(fragment.fromLine, fragment.toLine);
      const hasSegments = lines.some(line => line.segments && line.segments.length > 0);
      
      if (hasSegments) {
        this.renderLinesContentWithPmPositions(item, block.runs, lines, measure.lines, fragment.fromLine, pmPos);
        list.appendChild(item);
        container.appendChild(list);
        return;
      }
    }
    
    // Fallback: render all runs with PM positions and clipping
    this.renderRuns(item, block.runs, pmPos);
    list.appendChild(item);
    
    // Handle partial fragments
    if (!fragment.isFirstFragment || !fragment.isLastFragment) {
      const wrapper = document.createElement('div');
      wrapper.style.cssText = `
        overflow: hidden;
        height: ${fragment.height}px;
      `;
      
      if (!fragment.isFirstFragment && hasLineData(measure)) {
        let offset = 0;
        for (let i = 0; i < fragment.fromLine; i++) {
          offset += measure.lines[i].height * this.config.scale;
        }
        list.style.marginTop = `-${offset}px`;
      }
      
      wrapper.appendChild(list);
      container.appendChild(wrapper);
    } else {
      container.appendChild(list);
    }
  }
  
  /**
   * Render a table fragment with PM position data for cell selection
   */
  private renderTableFragment(
    container: HTMLElement,
    block: TableBlock,
    _fragment: PageFragment
  ): void {
    const data = this.blockLookup.get(block.id);
    const tablePmPos = data?.pmPos;
    
    const table = document.createElement('table');
    table.className = 'table-block';
    table.style.cssText = `
      border-collapse: collapse;
      width: 100%;
      margin: ${16 * this.config.scale}px 0;
    `;
    
    // Add PM position data for the table
    if (tablePmPos) {
      table.dataset.pmStart = String(tablePmPos.start);
      table.dataset.pmEnd = String(tablePmPos.end);
    }
    
    for (const row of block.rows) {
      const tr = document.createElement('tr');
      
      for (const cell of row.cells) {
        const td = document.createElement('td');
        td.className = 'table-cell';
        
        // Build cell styles including optional background color and width
        let cellStyles = `
          border: 1px solid #dadce0;
          padding: ${4 * this.config.scale}px ${8 * this.config.scale}px;
          vertical-align: top;
          min-height: 1em;
        `;
        const cellBg = safeCssColor(cell.backgroundColor);
        if (cellBg) {
          cellStyles += `background-color: ${cellBg};`;
        }
        // Apply column width if specified
        if (cell.colwidth && cell.colwidth.length > 0 && cell.colwidth[0]) {
          cellStyles += `width: ${cell.colwidth[0] * this.config.scale}px;`;
        }
        td.style.cssText = cellStyles;
        
        // Add PM position data for click-to-position mapping
        if (cell.pmStart !== undefined && cell.pmEnd !== undefined) {
          td.dataset.pmStart = String(cell.pmStart);
          td.dataset.pmEnd = String(cell.pmEnd);
        }
        
        // Render cell content with PM positions
        let charOffset = 0;
        for (const cellBlock of cell.blocks) {
          if (isRunContainingBlock(cellBlock)) {
            const p = document.createElement('p');
            p.className = 'cell-paragraph';
            
            // Get paragraph attributes including alignment
            const attrs = cellBlock.attrs || {};
            p.style.margin = '0';
            p.style.minHeight = '1em';
            p.style.textAlign = safeCssKeyword(attrs.alignment) || 'left';
            
            // Calculate PM position for this paragraph within the cell
            // Cell structure: cell_start + 1 (enter cell) + paragraph content
            const paragraphPmStart = cell.pmStart !== undefined 
              ? cell.pmStart + 1 + charOffset 
              : undefined;
            
            // Render runs with PM position tracking
            if (paragraphPmStart !== undefined) {
              let runCharOffset = 0;
              for (const run of cellBlock.runs) {
                if (run.kind === 'text') {
                  const span = document.createElement('span');
                  span.textContent = run.text;
                  span.style.cssText = this.getRunStyles(run);
                  
                  // Add PM position for this span
                  const spanStart = paragraphPmStart + 1 + runCharOffset; // +1 to enter paragraph
                  const spanEnd = spanStart + run.text.length;
                  span.dataset.pmStart = String(spanStart);
                  span.dataset.pmEnd = String(spanEnd);
                  runCharOffset += run.text.length;

                  this.applyCommentDecoration(span, run);
                  p.appendChild(span);
                } else if (run.kind === 'link') {
                  // A link run must render as an anchor — the previous loop only
                  // handled text/lineBreak, so links in a cell rendered as nothing.
                  const a = document.createElement('a');
                  a.textContent = run.text;
                  a.href = sanitizeHref(run.href);
                  a.style.cssText = `color: ${safeCssColor(run.color) || '#1a73e8'}; text-decoration: underline;${run.bold ? ' font-weight: bold;' : ''}${run.italic ? ' font-style: italic;' : ''}`;
                  const linkStart = paragraphPmStart + 1 + runCharOffset;
                  a.dataset.pmStart = String(linkStart);
                  a.dataset.pmEnd = String(linkStart + run.text.length);
                  runCharOffset += run.text.length;
                  this.applyCommentDecoration(a, run);
                  p.appendChild(a);
                } else if (run.kind === 'image') {
                  const img = document.createElement('img');
                  img.src = sanitizeImageSrc(run.src);
                  img.alt = run.alt || '';
                  img.width = run.width * this.config.scale;
                  img.height = run.height * this.config.scale;
                  runCharOffset += 1; // an inline image is one PM position
                  p.appendChild(img);
                } else if (run.kind === 'lineBreak') {
                  p.appendChild(document.createElement('br'));
                  runCharOffset += 1; // a hard break is one PM position — advance,
                  // or every subsequent cell's PM positions drift by one.
                }
              }
              
              // Add PM positions to the paragraph
              p.dataset.pmStart = String(paragraphPmStart);
              p.dataset.pmEnd = String(paragraphPmStart + 1 + runCharOffset + 1); // +1 enter, +1 exit
              
              charOffset += runCharOffset + 2; // paragraph nodeSize
            } else {
              // Fallback without PM positions
              this.renderRuns(p, cellBlock.runs);
            }
            
            // If paragraph is empty, add zero-width space
            if (p.childNodes.length === 0) {
              p.innerHTML = '&#8203;';
              // Still set PM positions for empty paragraphs
              if (cell.pmStart !== undefined) {
                p.dataset.pmStart = String(cell.pmStart + 1);
                p.dataset.pmEnd = String(cell.pmStart + 3); // Empty paragraph is 2 nodeSize
              }
            }
            
            td.appendChild(p);
          }
        }
        
        // If cell is empty, ensure it has at least some content for clicking
        if (td.childNodes.length === 0) {
          const emptyP = document.createElement('p');
          emptyP.style.margin = '0';
          emptyP.style.minHeight = '1em';
          emptyP.innerHTML = '&#8203;';
          if (cell.pmStart !== undefined) {
            emptyP.dataset.pmStart = String(cell.pmStart + 1);
            emptyP.dataset.pmEnd = String(cell.pmStart + 3);
          }
          td.appendChild(emptyP);
        }
        
        if (cell.colspan) td.colSpan = cell.colspan;
        if (cell.rowspan) td.rowSpan = cell.rowspan;
        
        tr.appendChild(td);
      }
      
      table.appendChild(tr);
    }
    
    container.appendChild(table);
  }
  
  /**
   * Render an image fragment
   */
  private renderImageFragment(
    container: HTMLElement,
    block: ImageBlock,
    _fragment: PageFragment
  ): void {
    const data = this.blockLookup.get(block.id);
    const pmPos = data?.pmPos;
    
    const wrapper = document.createElement('div');
    wrapper.className = 'image-block';
    wrapper.style.cssText = `
      text-align: ${safeCssKeyword(block.alignment) || 'center'};
      margin: ${8 * this.config.scale}px 0;
      position: relative;
    `;
    
    // Add PM position data for drag operations
    if (pmPos) {
      wrapper.dataset.pmStart = String(pmPos.start);
      wrapper.dataset.pmEnd = String(pmPos.end);
    }
    
    const img = document.createElement('img');
    img.src = sanitizeImageSrc(block.src);
    img.alt = block.alt || '';
    img.draggable = false; // Prevent native image drag
    img.style.cssText = `
      max-width: 100%;
      width: ${block.width * this.config.scale}px;
      height: ${block.height * this.config.scale}px;
      object-fit: contain;
      pointer-events: auto;
    `;
    
    wrapper.appendChild(img);
    container.appendChild(wrapper);
  }
  
  /**
   * Render a horizontal rule
   */
  private renderHorizontalRule(container: HTMLElement): void {
    const hr = document.createElement('hr');
    hr.style.cssText = `
      border: none;
      border-top: 1px solid #dadce0;
      margin: ${12 * this.config.scale}px 0;
    `;
    container.appendChild(hr);
  }
  
  /**
   * Render runs to a container element
   */
  private renderRuns(container: HTMLElement, runs: Run[], pmPos?: BlockPmPosition): void {
    let charOffset = 0;
    
    for (const run of runs) {
      if (run.kind === 'text') {
        const span = document.createElement('span');
        span.textContent = run.text;
        span.style.cssText = this.getRunStyles(run);
        
        // Add PM position data attributes
        if (pmPos) {
          const spanStart = charOffset;
          const spanEnd = charOffset + run.text.length;
          span.dataset.pmStart = String(pmPos.start + 1 + spanStart);
          span.dataset.pmEnd = String(pmPos.start + 1 + spanEnd);
          charOffset = spanEnd;
        }

        this.applyCommentDecoration(span, run);
        container.appendChild(span);
      } else if (run.kind === 'lineBreak') {
        container.appendChild(document.createElement('br'));
      } else if (run.kind === 'image') {
        const img = document.createElement('img');
        img.src = sanitizeImageSrc(run.src);
        img.width = run.width * this.config.scale;
        img.height = run.height * this.config.scale;
        img.alt = run.alt || '';
        img.style.verticalAlign = 'middle';
        container.appendChild(img);
      } else if (run.kind === 'link') {
        const a = document.createElement('a');
        a.textContent = run.text;
        a.href = sanitizeHref(run.href);
        a.style.cssText = `
          color: ${safeCssColor(run.color) || '#1a73e8'};
          text-decoration: underline;
          ${run.bold ? 'font-weight: bold;' : ''}
          ${run.italic ? 'font-style: italic;' : ''}
        `;
        
        // Add PM position data attributes
        if (pmPos) {
          const spanStart = charOffset;
          const spanEnd = charOffset + run.text.length;
          a.dataset.pmStart = String(pmPos.start + 1 + spanStart);
          a.dataset.pmEnd = String(pmPos.start + 1 + spanEnd);
          charOffset = spanEnd;
        }

        this.applyCommentDecoration(a, run);
        container.appendChild(a);
      }
    }
    
    // If no runs, add zero-width space for proper height
    if (runs.length === 0) {
      container.innerHTML = '&#8203;';
    }
  }
  
  /**
   * Decorate an element that's covered by a comment thread: a highlight tint,
   * a pointer cursor, and a data attribute the click handler keys off.
   */
  private applyCommentDecoration(el: HTMLElement, run: Run): void {
    const threadId =
      run.kind === 'text' || run.kind === 'link' ? run.commentThreadId : undefined;
    if (!threadId) return;
    el.setAttribute('data-comment-thread', threadId);
    el.style.backgroundColor = 'rgba(99, 102, 241, 0.18)';
    el.style.cursor = 'pointer';
  }

  /**
   * Get CSS styles for a text run
   */
  private getRunStyles(run: Run): string {
    if (run.kind !== 'text') return '';
    
    const styles: string[] = [];
    
    if (run.bold) styles.push('font-weight: bold');
    if (run.italic) styles.push('font-style: italic');
    if (run.underline) styles.push('text-decoration: underline');
    if (run.strikethrough) styles.push('text-decoration: line-through');
    if (run.fontSize) styles.push(`font-size: ${run.fontSize * this.config.scale}px`);
    const runFont = safeFontFamily(run.fontFamily);
    if (runFont) styles.push(`font-family: ${runFont}`);
    const runColor = safeCssColor(run.color);
    if (runColor) styles.push(`color: ${runColor}`);
    const runBgColor = safeCssColor(run.backgroundColor);
    if (runBgColor) styles.push(`background-color: ${runBgColor}`);
    if (run.superscript) styles.push('vertical-align: super; font-size: 0.8em');
    if (run.subscript) styles.push('vertical-align: sub; font-size: 0.8em');
    
    return styles.join('; ');
  }
  
  /**
   * Get scaled page width
   */
  private getPageWidth(): number {
    return this.config.pageConfig.width * this.config.scale;
  }
  
  /**
   * Get scaled page height
   */
  private getPageHeight(): number {
    return this.config.pageConfig.height * this.config.scale;
  }
  
  /**
   * Get scaled margins
   */
  private getScaledMargins(): { top: number; bottom: number; left: number; right: number } {
    const m = this.config.pageConfig.margins;
    const s = this.config.scale;
    return {
      top: m.top * s,
      bottom: m.bottom * s,
      left: m.left * s,
      right: m.right * s,
    };
  }
}

// ============================================================================
// React Component Wrapper
// ============================================================================

/**
 * Props for the PageView component
 */
export interface PageViewProps {
  page: PageLayout;
  layout: DocumentLayout;
  blocks: FlowBlock[];
  measures: Measure[];
  config: PainterConfig;
  onClick?: (e: React.MouseEvent, pageIndex: number) => void;
}

/**
 * Create a React-friendly page renderer
 * This is a helper for integrating with React components
 */
export function createPageElement(
  page: PageLayout,
  layout: DocumentLayout,
  blocks: FlowBlock[],
  measures: Measure[],
  config: PainterConfig
): HTMLElement {
  const painter = new DomPainter(config);
  painter.setData(blocks, measures);
  return painter.paintPage(page, layout).element;
}

