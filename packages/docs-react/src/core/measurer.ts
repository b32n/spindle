/**
 * DOM Measurer - Measures FlowBlocks by rendering them to a hidden DOM element
 * 
 * This is a critical component of the true layout engine. It:
 * 1. Renders each block to a hidden measurement container
 * 2. Measures line heights using getClientRects()
 * 3. Returns precise measurements for layout calculation
 * 
 * The key insight is that we need LINE-LEVEL measurements, not just block heights.
 * This allows us to split paragraphs across pages at line boundaries.
 */

import {
  FlowBlock,
  Run,
  ParagraphBlock,
  HeadingBlock,
  ListItemBlock,
  TableBlock,
  ImageBlock,
  HorizontalRuleBlock,
  PageBreakBlock,
  isRunContainingBlock,
  isTextRun,
} from './flow-blocks';
import {
  sanitizeHref,
  sanitizeImageSrc,
  safeCssColor,
  safeFontFamily,
  safeCssKeyword,
  safeLineHeight,
  safeCssNumber,
} from './sanitize';

// ============================================================================
// Measurement Types
// ============================================================================

/**
 * A segment of a run that appears on a specific line
 */
export interface LineRunSegment {
  /** Index of the run in the block's runs array */
  runIndex: number;
  /** Start character index within the run's text */
  startOffset: number;
  /** End character index within the run's text (exclusive) */
  endOffset: number;
  /** The actual text content for this segment */
  text: string;
}

/**
 * Measurement for a single line of text
 */
export interface LineMeasure {
  /** Height of this line in pixels */
  height: number;
  /** Width of this line in pixels */
  width: number;
  /** Baseline offset from top of line */
  ascent: number;
  /** Descent below baseline */
  descent: number;
  /** Y offset from the top of the block */
  yOffset: number;
  /** Run segments that make up this line (for rendering partial blocks) */
  segments?: LineRunSegment[];
}

/**
 * Measurement for a paragraph or text block
 */
export interface ParagraphMeasure {
  kind: 'paragraph';
  /** Block ID */
  blockId: string;
  /** Total height of the block */
  totalHeight: number;
  /** Individual line measurements */
  lines: LineMeasure[];
  /** Space before the paragraph */
  spaceBefore: number;
  /** Space after the paragraph */
  spaceAfter: number;
}

/**
 * Measurement for a heading
 */
export interface HeadingMeasure {
  kind: 'heading';
  blockId: string;
  totalHeight: number;
  lines: LineMeasure[];
  spaceBefore: number;
  spaceAfter: number;
}

/**
 * Measurement for a list item
 */
export interface ListItemMeasure {
  kind: 'listItem';
  blockId: string;
  totalHeight: number;
  lines: LineMeasure[];
  spaceBefore: number;
  spaceAfter: number;
  /** Indent for the bullet/number */
  bulletIndent: number;
}

/**
 * Measurement for a table
 */
export interface TableMeasure {
  kind: 'table';
  blockId: string;
  totalHeight: number;
  /** Row heights */
  rowHeights: number[];
}

/**
 * Measurement for an image
 */
export interface ImageMeasure {
  kind: 'image';
  blockId: string;
  totalHeight: number;
  width: number;
}

/**
 * Measurement for a horizontal rule
 */
export interface HorizontalRuleMeasure {
  kind: 'horizontalRule';
  blockId: string;
  totalHeight: number;
}

/**
 * Measurement for a page break
 */
export interface PageBreakMeasure {
  kind: 'pageBreak';
  blockId: string;
  totalHeight: number;
  forceBreak: true;
}

/**
 * Union type for all measurements
 */
export type Measure =
  | ParagraphMeasure
  | HeadingMeasure
  | ListItemMeasure
  | TableMeasure
  | ImageMeasure
  | HorizontalRuleMeasure
  | PageBreakMeasure;

// ============================================================================
// DOM Measurer Class
// ============================================================================

/**
 * Configuration for the measurer
 */
export interface MeasurerConfig {
  /** Width of the content area in pixels */
  contentWidth: number;
  /** Default font family */
  fontFamily?: string;
  /** Default font size in pixels */
  fontSize?: number;
  /** Default line height multiplier */
  lineHeight?: number;
}

/**
 * DomMeasurer - Measures FlowBlocks by rendering to a hidden container
 */
export class DomMeasurer {
  private measureContainer: HTMLDivElement;
  private config: Required<MeasurerConfig>;
  private cache: Map<string, Measure> = new Map();
  
  constructor(config: MeasurerConfig) {
    this.config = {
      contentWidth: config.contentWidth,
      fontFamily: config.fontFamily ?? 'Arial, sans-serif',
      fontSize: config.fontSize ?? 11,
      lineHeight: config.lineHeight ?? 1.5,
    };
    
    // Create hidden measurement container
    this.measureContainer = document.createElement('div');
    this.measureContainer.className = 'flow-measure-container';
    this.measureContainer.style.cssText = `
      position: absolute;
      visibility: hidden;
      width: ${this.config.contentWidth}px;
      left: -9999px;
      top: 0;
      font-family: ${this.config.fontFamily};
      font-size: ${this.config.fontSize}px;
      line-height: ${this.config.lineHeight};
      white-space: pre-wrap;
      word-wrap: break-word;
    `;
    document.body.appendChild(this.measureContainer);
  }
  
  /**
   * Update the content width (e.g., when page margins change)
   */
  updateContentWidth(width: number): void {
    this.config.contentWidth = width;
    this.measureContainer.style.width = `${width}px`;
    this.clearCache();
  }
  
  /**
   * Clear the measurement cache
   */
  clearCache(): void {
    this.cache.clear();
  }
  
  /**
   * Measure a single block
   */
  measureBlock(block: FlowBlock): Measure {
    // Check cache first
    const cached = this.cache.get(block.id);
    if (cached) {
      return cached;
    }
    
    let measure: Measure;
    
    switch (block.kind) {
      case 'paragraph':
        measure = this.measureParagraph(block);
        break;
      case 'heading':
        measure = this.measureHeading(block);
        break;
      case 'listItem':
        measure = this.measureListItem(block);
        break;
      case 'table':
        measure = this.measureTable(block);
        break;
      case 'image':
        measure = this.measureImage(block);
        break;
      case 'horizontalRule':
        measure = this.measureHorizontalRule(block);
        break;
      case 'pageBreak':
        measure = this.measurePageBreak(block);
        break;
    }
    
    // Cache the result
    this.cache.set(block.id, measure);
    return measure;
  }
  
  /**
   * Measure multiple blocks
   */
  measureBlocks(blocks: FlowBlock[]): Measure[] {
    return blocks.map(block => this.measureBlock(block));
  }
  
  /**
   * Measure a paragraph block with line-level precision
   */
  private measureParagraph(block: ParagraphBlock): ParagraphMeasure {
    // Clear container
    this.measureContainer.innerHTML = '';
    
    // Create paragraph element
    const para = document.createElement('p');
    para.style.cssText = this.getParagraphStyles(block);
    
    // Add content
    this.renderRuns(para, block.runs);
    this.measureContainer.appendChild(para);
    
    // Measure lines with run tracking
    const lines = this.measureLines(para, block.runs);
    
    // Get spacing
    const computedStyle = window.getComputedStyle(para);
    const spaceBefore = parseFloat(computedStyle.marginTop) || 0;
    const spaceAfter = parseFloat(computedStyle.marginBottom) || 0;
    
    // Calculate total height
    const totalHeight = lines.reduce((sum, line) => sum + line.height, 0) + spaceBefore + spaceAfter;
    
    return {
      kind: 'paragraph',
      blockId: block.id,
      totalHeight,
      lines,
      spaceBefore,
      spaceAfter,
    };
  }
  
  /**
   * Measure a heading block
   */
  private measureHeading(block: HeadingBlock): HeadingMeasure {
    // Clear container
    this.measureContainer.innerHTML = '';
    
    // Create heading element
    const heading = document.createElement(`h${block.level}`);
    heading.style.cssText = this.getHeadingStyles(block);
    
    // Add content
    this.renderRuns(heading, block.runs);
    this.measureContainer.appendChild(heading);
    
    // Measure lines with run tracking
    const lines = this.measureLines(heading, block.runs);
    
    // Get spacing
    const computedStyle = window.getComputedStyle(heading);
    const spaceBefore = parseFloat(computedStyle.marginTop) || 0;
    const spaceAfter = parseFloat(computedStyle.marginBottom) || 0;
    
    const totalHeight = lines.reduce((sum, line) => sum + line.height, 0) + spaceBefore + spaceAfter;
    
    return {
      kind: 'heading',
      blockId: block.id,
      totalHeight,
      lines,
      spaceBefore,
      spaceAfter,
    };
  }
  
  /**
   * Measure a list item block
   */
  private measureListItem(block: ListItemBlock): ListItemMeasure {
    // Clear container
    this.measureContainer.innerHTML = '';
    
    // Create list structure
    const list = document.createElement(block.listType === 'bullet' ? 'ul' : 'ol');
    const item = document.createElement('li');
    
    const indent = 24 * (block.level + 1); // 24px per indent level
    list.style.cssText = `
      margin: 0;
      padding-left: ${indent}px;
      list-style-position: outside;
    `;
    item.style.cssText = this.getListItemStyles(block);
    
    // Add content
    this.renderRuns(item, block.runs);
    list.appendChild(item);
    this.measureContainer.appendChild(list);
    
    // Measure lines with run tracking
    const lines = this.measureLines(item, block.runs);
    
    // Get spacing
    const computedStyle = window.getComputedStyle(item);
    const spaceBefore = parseFloat(computedStyle.marginTop) || 0;
    const spaceAfter = parseFloat(computedStyle.marginBottom) || 4;
    
    const totalHeight = lines.reduce((sum, line) => sum + line.height, 0) + spaceBefore + spaceAfter;
    
    return {
      kind: 'listItem',
      blockId: block.id,
      totalHeight,
      lines,
      spaceBefore,
      spaceAfter,
      bulletIndent: indent,
    };
  }
  
  /**
   * Measure a table block
   */
  private measureTable(block: TableBlock): TableMeasure {
    // Clear container
    this.measureContainer.innerHTML = '';
    
    // Margin value must match dom-painter.ts (16px at scale 1)
    const tableMargin = 16;
    
    // Create table
    const table = document.createElement('table');
    table.style.cssText = `
      border-collapse: collapse;
      width: 100%;
      margin: ${tableMargin}px 0;
    `;
    
    const rowHeights: number[] = [];
    
    for (const row of block.rows) {
      const tr = document.createElement('tr');
      
      for (const cell of row.cells) {
        const td = document.createElement('td');
        let cellStyles = `
          border: 1px solid #dadce0;
          padding: 4px 8px;
          vertical-align: top;
        `;
        
        // Apply column width if specified (use the same value as dom-painter)
        if (cell.colwidth && cell.colwidth.length > 0 && cell.colwidth[0]) {
          cellStyles += `width: ${cell.colwidth[0]}px;`;
        }
        
        td.style.cssText = cellStyles;
        
        // Render cell content (simplified - just get text)
        for (const cellBlock of cell.blocks) {
          if (isRunContainingBlock(cellBlock)) {
            const p = document.createElement('p');
            p.style.margin = '0';
            this.renderRuns(p, cellBlock.runs);
            td.appendChild(p);
          }
        }
        
        if (cell.colspan) td.colSpan = cell.colspan;
        if (cell.rowspan) td.rowSpan = cell.rowspan;
        
        tr.appendChild(td);
      }
      
      table.appendChild(tr);
    }
    
    this.measureContainer.appendChild(table);
    
    // Measure row heights
    const rows = table.querySelectorAll('tr');
    rows.forEach((tr) => {
      rowHeights.push(tr.getBoundingClientRect().height);
    });
    
    // getBoundingClientRect doesn't include margins, so we add them explicitly
    // Top margin + table height + bottom margin
    const tableHeight = table.getBoundingClientRect().height;
    const totalHeight = tableHeight + (tableMargin * 2);
    
    return {
      kind: 'table',
      blockId: block.id,
      totalHeight,
      rowHeights,
    };
  }
  
  /**
   * Measure an image block
   */
  private measureImage(block: ImageBlock): ImageMeasure {
    // Images don't need DOM measurement - we know their dimensions
    const margin = 8;
    const totalHeight = block.height + margin * 2;
    
    return {
      kind: 'image',
      blockId: block.id,
      totalHeight,
      width: block.width,
    };
  }
  
  /**
   * Measure a horizontal rule
   */
  private measureHorizontalRule(block: HorizontalRuleBlock): HorizontalRuleMeasure {
    return {
      kind: 'horizontalRule',
      blockId: block.id,
      totalHeight: 24, // Fixed height for HR
    };
  }
  
  /**
   * Measure a page break
   */
  private measurePageBreak(block: PageBreakBlock): PageBreakMeasure {
    return {
      kind: 'pageBreak',
      blockId: block.id,
      totalHeight: 0, // Page breaks have no height
      forceBreak: true,
    };
  }
  
  /**
   * Measure individual lines within an element using getClientRects
   * Also tracks which text content belongs to each line for proper fragment rendering
   * 
   * IMPORTANT: getClientRects() returns glyph bounding boxes, which are based on
   * font metrics and do NOT include line-height spacing. We must use the computed
   * line-height for accurate height calculations.
   */
  private measureLines(element: HTMLElement, runs?: Run[]): LineMeasure[] {
    const lines: LineMeasure[] = [];
    const elementRect = element.getBoundingClientRect();
    
    // Get computed line height - this is the ACTUAL line height used for rendering
    // getClientRects() returns glyph heights which are smaller than line-height
    const computedStyle = window.getComputedStyle(element);
    const computedLineHeight = parseFloat(computedStyle.lineHeight);
    const defaultLineHeight = computedLineHeight || this.config.fontSize * this.config.lineHeight;
    
    // Collect all text nodes with their positions
    const textNodes: Array<{ node: Text; runIndex: number; startInRun: number }> = [];
    let runIndex = 0;
    
    const collectTextNodes = (node: Node) => {
      if (node.nodeType === Node.TEXT_NODE) {
        const textNode = node as Text;
        if (textNode.textContent && textNode.textContent.length > 0) {
          // Find which run this text node belongs to
          // This is approximate - we track by order
          textNodes.push({ node: textNode, runIndex, startInRun: 0 });
        }
      } else if (node.nodeType === Node.ELEMENT_NODE) {
        const el = node as HTMLElement;
        // Check if this is a span (run container) or br
        if (el.tagName === 'BR') {
          runIndex++;
        } else if (el.tagName === 'IMG') {
          // An inline image is its own run (renderRuns appends a bare <img>).
          // It has no text nodes, but must still advance runIndex or every run
          // after it maps to the wrong slot and the painter drops/mis-styles it.
          runIndex++;
        } else if (el.tagName === 'SPAN' || el.tagName === 'A') {
          // This is a run container
          for (const child of Array.from(el.childNodes)) {
            collectTextNodes(child);
          }
          runIndex++;
        } else {
          for (const child of Array.from(el.childNodes)) {
            collectTextNodes(child);
          }
        }
      }
    };
    
    for (const child of Array.from(element.childNodes)) {
      collectTextNodes(child);
    }
    
    if (textNodes.length === 0) {
      // Empty element - return a single line with default height
      lines.push({
        height: defaultLineHeight,
        width: elementRect.width,
        ascent: defaultLineHeight * 0.8,
        descent: defaultLineHeight * 0.2,
        yOffset: 0,
        segments: [],
      });
      return lines;
    }
    
    // Use word-based approach for efficiency (check whole words instead of chars)
    // This is much faster than character-by-character for long paragraphs
    interface LineInfo {
      top: number;
      width: number;
      maxGlyphHeight: number;  // Track the tallest glyph on this line for mixed font sizes
      segments: LineRunSegment[];
    }
    
    const lineInfos: LineInfo[] = [];
    let currentLine: LineInfo | null = null;
    
    for (const { node, runIndex } of textNodes) {
      const text = node.textContent || '';
      if (text.length === 0) continue;
      
      // Find line breaks using binary search within words
      let startIndex = 0;
      
      while (startIndex < text.length) {
        // Find the end of the current word/segment
        let endIndex = startIndex + 1;
        
        // Skip to end of word or use chunks of ~20 chars for efficiency
        while (endIndex < text.length && endIndex - startIndex < 20) {
          if (text[endIndex] === ' ' || text[endIndex] === '\n') {
            endIndex++;
            break;
          }
          endIndex++;
        }
        
        // Get rect for this segment
        const range = document.createRange();
        range.setStart(node, startIndex);
        range.setEnd(node, Math.min(endIndex, text.length));
        
        const rects = range.getClientRects();
        if (rects.length === 0) {
          startIndex = endIndex;
          continue;
        }
        
        // Process each rect (a segment might span multiple lines)
        for (let r = 0; r < rects.length; r++) {
          const rect = rects[r];
          const relativeTop = rect.top - elementRect.top;
          
          // Determine which part of the text is on this line
          // For simplicity, if there's only one rect, all text is on that line
          // If multiple rects, we need to find the break point
          let segmentStart = startIndex;
          let segmentEnd = endIndex;
          
          if (rects.length > 1) {
            // Multiple rects means line break within this segment
            // Use binary search to find the break point
            segmentEnd = this.findLineBreakPoint(node, startIndex, endIndex, relativeTop, elementRect.top);
            if (r > 0) {
              segmentStart = this.findLineBreakPoint(node, startIndex, endIndex, relativeTop - 5, elementRect.top);
            }
          }
          
          const segmentText = text.substring(segmentStart, segmentEnd);
          
          // Check if this is a new line - use defaultLineHeight for comparison threshold
          // since line tops are spaced by line-height, not glyph height
          const isNewLine = currentLine === null || Math.abs(relativeTop - currentLine.top) > defaultLineHeight * 0.5;
          
          if (isNewLine) {
            // Start a new line
            if (currentLine !== null) {
              lineInfos.push(currentLine);
            }
            
            currentLine = {
              top: relativeTop,
              width: rect.width,
              maxGlyphHeight: rect.height,  // Track glyph height for this line
              segments: segmentText.length > 0 ? [{
                runIndex,
                startOffset: segmentStart,
                endOffset: segmentEnd,
                text: segmentText,
              }] : [],
            };
          } else if (currentLine && segmentText.length > 0) {
            // Continue current line
            currentLine.width = Math.max(currentLine.width, rect.right - elementRect.left);
            // Update max glyph height if this segment is taller (different font size)
            currentLine.maxGlyphHeight = Math.max(currentLine.maxGlyphHeight, rect.height);
            
            // Extend or add segment
            const lastSegment = currentLine.segments[currentLine.segments.length - 1];
            if (lastSegment && lastSegment.runIndex === runIndex && lastSegment.endOffset === segmentStart) {
              // Extend existing segment
              lastSegment.endOffset = segmentEnd;
              lastSegment.text += segmentText;
            } else {
              // Add new segment
              currentLine.segments.push({
                runIndex,
                startOffset: segmentStart,
                endOffset: segmentEnd,
                text: segmentText,
              });
            }
          }
        }
        
        startIndex = endIndex;
      }
    }
    
    // Don't forget the last line
    if (currentLine !== null) {
      lineInfos.push(currentLine);
    }
    
    // Convert to LineMeasure format
    // For lines with mixed font sizes, we need to calculate the actual line height
    // based on the tallest glyph in that line, applying the line-height multiplier
    const lineHeightMultiplier = this.config.lineHeight;
    
    for (const info of lineInfos) {
      // Calculate line height based on the actual glyph height for this line
      // This handles mixed font sizes within a paragraph
      // Use the larger of: defaultLineHeight OR (maxGlyphHeight * lineHeightMultiplier)
      const computedHeight = info.maxGlyphHeight * lineHeightMultiplier;
      const lineHeight = Math.max(defaultLineHeight, computedHeight);
      
      lines.push({
        height: lineHeight,
        width: info.width,
        ascent: lineHeight * 0.8,
        descent: lineHeight * 0.2,
        yOffset: info.top,
        segments: info.segments,
      });
    }
    
    // If no lines were detected, use fallback
    if (lines.length === 0) {
      const allText = runs?.filter(isTextRun).map(r => r.text).join('') || '';
      lines.push({
        height: elementRect.height || defaultLineHeight,
        width: elementRect.width,
        ascent: (elementRect.height || defaultLineHeight) * 0.8,
        descent: (elementRect.height || defaultLineHeight) * 0.2,
        yOffset: 0,
        segments: allText ? [{ runIndex: 0, startOffset: 0, endOffset: allText.length, text: allText }] : [],
      });
    }
    
    return lines;
  }
  
  /**
   * Binary search to find where a line break occurs within a text range
   */
  private findLineBreakPoint(
    node: Text,
    start: number,
    end: number,
    targetY: number,
    elementTop: number
  ): number {
    if (end - start <= 1) {
      return end;
    }
    
    const mid = Math.floor((start + end) / 2);
    const range = document.createRange();
    range.setStart(node, mid);
    range.setEnd(node, mid + 1);
    
    const rects = range.getClientRects();
    if (rects.length === 0) {
      return mid;
    }
    
    const midY = rects[0].top - elementTop;
    
    if (Math.abs(midY - targetY) < 2) {
      // On the same line, search later
      return this.findLineBreakPoint(node, mid, end, targetY, elementTop);
    } else if (midY < targetY) {
      // Before target line, search later
      return this.findLineBreakPoint(node, mid, end, targetY, elementTop);
    } else {
      // After target line, search earlier
      return this.findLineBreakPoint(node, start, mid, targetY, elementTop);
    }
  }
  
  /**
   * Render runs to a DOM element
   */
  private renderRuns(container: HTMLElement, runs: Run[]): void {
    for (const run of runs) {
      if (run.kind === 'text') {
        const span = document.createElement('span');
        span.textContent = run.text;
        span.style.cssText = this.getRunStyles(run);
        container.appendChild(span);
      } else if (run.kind === 'lineBreak') {
        container.appendChild(document.createElement('br'));
      } else if (run.kind === 'image') {
        const img = document.createElement('img');
        img.src = sanitizeImageSrc(run.src);
        img.width = run.width;
        img.height = run.height;
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
        container.appendChild(a);
      }
    }
    
    // If no runs, add a zero-width space to ensure element has height
    if (runs.length === 0) {
      container.innerHTML = '&#8203;';
    }
  }
  
  /**
   * Get CSS styles for a paragraph
   */
  private getParagraphStyles(block: ParagraphBlock): string {
    const attrs = block.attrs || {};
    // These numeric attrs come from untrusted document/collab JSON; coerce each
    // (a string like "8px 0; background: url(//evil)" would otherwise inject a
    // CSS declaration into this measurement element's style).
    return `
      margin: ${safeCssNumber(attrs.spaceBefore, 0)}px 0 ${safeCssNumber(attrs.spaceAfter, 8)}px 0;
      text-align: ${safeCssKeyword(attrs.alignment) || 'left'};
      line-height: ${safeLineHeight(attrs.lineHeight) ?? this.config.lineHeight};
      text-indent: ${safeCssNumber(attrs.firstLineIndent, 0)}px;
      padding-left: ${safeCssNumber(attrs.leftIndent, 0)}px;
      padding-right: ${safeCssNumber(attrs.rightIndent, 0)}px;
    `;
  }
  
  /**
   * Get CSS styles for a heading
   */
  private getHeadingStyles(block: HeadingBlock): string {
    const attrs = block.attrs || {};
    const fontSizes: Record<number, number> = {
      1: 24,
      2: 20,
      3: 16,
      4: 14,
      5: 12,
      6: 11,
    };
    const fontSize = fontSizes[block.level] || 16;
    
    return `
      margin: ${safeCssNumber(attrs.spaceBefore, 16)}px 0 ${safeCssNumber(attrs.spaceAfter, 8)}px 0;
      font-size: ${fontSize}px;
      font-weight: ${block.level <= 2 ? 400 : 700};
      text-align: ${safeCssKeyword(attrs.alignment) || 'left'};
      line-height: 1.3;
    `;
  }
  
  /**
   * Get CSS styles for a list item
   */
  private getListItemStyles(block: ListItemBlock): string {
    const attrs = block.attrs || {};
    return `
      margin: 0 0 4px 0;
      text-align: ${safeCssKeyword(attrs.alignment) || 'left'};
      line-height: ${safeLineHeight(attrs.lineHeight) ?? this.config.lineHeight};
    `;
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
    if (run.fontSize) styles.push(`font-size: ${safeCssNumber(run.fontSize, 0)}px`);
    const runFont = safeFontFamily(run.fontFamily);
    if (runFont) styles.push(`font-family: ${runFont}`);
    const runColor = safeCssColor(run.color);
    if (runColor) styles.push(`color: ${runColor}`);
    const runBgColor = safeCssColor(run.backgroundColor);
    if (runBgColor) styles.push(`background-color: ${runBgColor}`);
    if (run.superscript) styles.push('vertical-align: super; font-size: 0.8em');
    if (run.subscript) styles.push('vertical-align: sub; font-size: 0.8em');
    if (run.smallCaps) styles.push('font-variant: small-caps');
    if (run.allCaps) styles.push('text-transform: uppercase');
    if (run.hidden) styles.push('opacity: 0.5');
    
    return styles.join('; ');
  }
  
  /**
   * Destroy the measurer and clean up DOM
   */
  destroy(): void {
    if (this.measureContainer.parentNode) {
      this.measureContainer.parentNode.removeChild(this.measureContainer);
    }
    this.cache.clear();
  }
}

// ============================================================================
// Helper Functions
// ============================================================================

/**
 * Check if a measure has line-level data
 */
export function hasLineData(measure: Measure): measure is ParagraphMeasure | HeadingMeasure | ListItemMeasure {
  return 'lines' in measure && Array.isArray(measure.lines);
}

/**
 * Get the number of lines in a measure
 */
export function getLineCount(measure: Measure): number {
  if (hasLineData(measure)) {
    return measure.lines.length;
  }
  return 1; // Non-text blocks are treated as single "line"
}

/**
 * Get the height of a specific line range
 */
export function getLineRangeHeight(
  measure: Measure,
  fromLine: number,
  toLine: number
): number {
  if (!hasLineData(measure)) {
    return measure.totalHeight;
  }
  
  let height = 0;
  const lines = measure.lines;
  
  for (let i = fromLine; i < toLine && i < lines.length; i++) {
    height += lines[i].height;
  }
  
  // Add spacing if this is the start of the block
  if (fromLine === 0) {
    height += measure.spaceBefore;
  }
  
  // Add spacing if this is the end of the block
  if (toLine >= lines.length) {
    height += measure.spaceAfter;
  }
  
  return height;
}

