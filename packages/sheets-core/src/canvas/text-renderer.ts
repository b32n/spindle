// Text Renderer - Handles text measurement and rendering

import type { CanvasTheme, TextStyle, Rect } from './types';
import { INDENT_PX } from './style-paint';

/** Whether the style asks for an underline or a strike-through line. */
const hasDecoration = (style: Partial<TextStyle>): boolean => (style.textDecoration !== undefined && style.textDecoration !== 'none') || style.strikethrough === true;

/**
 * Cached text metrics for a specific font configuration
 */
interface FontMetricsCache {
  font: string;
  lineHeight: number;
  baseline: number;
}

/**
 * Handles text measurement and rendering with caching
 */
export class TextRenderer {
  private theme: CanvasTheme;
  private metricsCache: Map<string, FontMetricsCache> = new Map();
  private measureCanvas: HTMLCanvasElement;
  private measureCtx: CanvasRenderingContext2D;
  
  constructor(theme: CanvasTheme) {
    this.theme = theme;
    
    // Create offscreen canvas for text measurement
    this.measureCanvas = document.createElement('canvas');
    const ctx = this.measureCanvas.getContext('2d');
    if (!ctx) {
      throw new Error('Failed to create measurement context');
    }
    this.measureCtx = ctx;
  }
  
  /**
   * Build CSS font string from style
   */
  buildFontString(style: Partial<TextStyle>): string {
    const fontWeight = style.fontWeight ?? 'normal';
    const fontStyle = style.fontStyle ?? 'normal';
    const fontSize = style.fontSize ?? this.theme.cellFontSize;
    const fontFamily = style.fontFamily ?? this.theme.cellFont;
    
    return `${fontStyle} ${fontWeight} ${fontSize}px ${fontFamily}`;
  }
  
  /**
   * Get font metrics (with caching)
   */
  getFontMetrics(font: string): FontMetricsCache {
    const cached = this.metricsCache.get(font);
    if (cached) {
      return cached;
    }
    
    // Measure font metrics
    this.measureCtx.font = font;
    const metrics = this.measureCtx.measureText('Mg');
    
    // Approximate line height as 1.2 * font size
    const fontSize = parseInt(font.match(/(\d+)px/)?.[1] ?? '11', 10);
    const lineHeight = fontSize * 1.2;
    
    // Use actual metrics if available, otherwise approximate
    const baseline = metrics.actualBoundingBoxAscent ?? fontSize * 0.8;
    
    const fontMetrics: FontMetricsCache = {
      font,
      lineHeight,
      baseline,
    };
    
    this.metricsCache.set(font, fontMetrics);
    return fontMetrics;
  }
  
  /**
   * Measure text width
   */
  measureText(text: string, style: Partial<TextStyle>): number {
    const font = this.buildFontString(style);
    this.measureCtx.font = font;
    return this.measureCtx.measureText(text).width;
  }
  
  /**
   * Truncate text to fit within maxWidth, adding ellipsis if needed
   */
  truncateText(text: string, maxWidth: number, style: Partial<TextStyle>): string {
    const font = this.buildFontString(style);
    this.measureCtx.font = font;
    
    const fullWidth = this.measureCtx.measureText(text).width;
    if (fullWidth <= maxWidth) {
      return text;
    }
    
    const ellipsis = '…';
    const ellipsisWidth = this.measureCtx.measureText(ellipsis).width;
    const availableWidth = maxWidth - ellipsisWidth;
    
    if (availableWidth <= 0) {
      return ellipsis;
    }
    
    // Binary search for the right truncation point
    let low = 0;
    let high = text.length;
    
    while (low < high) {
      const mid = Math.ceil((low + high) / 2);
      const truncated = text.substring(0, mid);
      const width = this.measureCtx.measureText(truncated).width;
      
      if (width <= availableWidth) {
        low = mid;
      } else {
        high = mid - 1;
      }
    }
    
    return text.substring(0, low) + ellipsis;
  }
  
  /**
   * Render text within bounds
   */
  renderText(
    ctx: CanvasRenderingContext2D,
    text: string,
    bounds: Rect,
    style: Partial<TextStyle>,
    padding: number = 4
  ): void {
    if (!text) return;

    // Rotated text follows a separate path (clipped to the cell, not truncated).
    const rotation = style.textRotation ?? 0;
    if (rotation === 0 && style.shrinkToFit) style = this.shrunkToFit(text, style, bounds.width - padding * 2 - this.indentOf(style));

    const font = this.buildFontString(style);
    const metrics = this.getFontMetrics(font);

    ctx.font = font;
    ctx.fillStyle = style.color ?? this.theme.cellTextColor;

    if (rotation !== 0) {
      this.renderRotatedText(ctx, text, bounds, style, metrics, rotation);
      return;
    }

    // Indent moves the text in from the edge it is aligned to (centred text has none)
    const indent = this.indentOf(style);
    const align = style.textAlign ?? 'left';
    const leftPad = padding + (align === 'left' ? indent : 0);
    const rightPad = padding + (align === 'right' ? indent : 0);

    // Calculate available width for text
    const availableWidth = bounds.width - leftPad - rightPad;
    
    // Truncate if necessary
    const displayText = this.truncateText(text, availableWidth, style);
    
    // Calculate x position based on alignment
    let x: number;
    const textAlign = style.textAlign ?? 'left';
    
    switch (textAlign) {
      case 'center':
        ctx.textAlign = 'center';
        x = bounds.x + bounds.width / 2;
        break;
      case 'right':
        ctx.textAlign = 'right';
        x = bounds.x + bounds.width - rightPad;
        break;
      case 'left':
      default:
        ctx.textAlign = 'left';
        x = bounds.x + leftPad;
        break;
    }
    
    // Calculate y position based on vertical alignment
    let y: number;
    const verticalAlign = style.verticalAlign ?? 'middle';
    
    switch (verticalAlign) {
      case 'top':
        y = bounds.y + padding + metrics.baseline;
        break;
      case 'bottom':
        y = bounds.y + bounds.height - padding;
        break;
      case 'middle':
      default:
        y = bounds.y + (bounds.height + metrics.baseline) / 2;
        break;
    }
    
    // Apply text decoration
    ctx.fillText(displayText, x, y);
    
    // Render text decoration (underline, strikethrough)
    if (hasDecoration(style)) {
      const textWidth = this.measureCtx.measureText(displayText).width;
      
      // Adjust x for decoration based on alignment
      let decorationX: number;
      switch (textAlign) {
        case 'center':
          decorationX = x - textWidth / 2;
          break;
        case 'right':
          decorationX = x - textWidth;
          break;
        default:
          decorationX = x;
      }
      
      ctx.strokeStyle = style.color ?? this.theme.cellTextColor;
      this.drawDecorations(ctx, style, decorationX, textWidth, y, metrics.baseline);
    }
  }

  /** The indent in pixels (a level moves the text `INDENT_PX`). */
  private indentOf(style: Partial<TextStyle>): number {
    return style.indent && style.indent > 0 ? style.indent * INDENT_PX : 0;
  }

  /** `style` with a font size at which `text` fits `width` (never larger than its own: shrink to fit does not enlarge). */
  private shrunkToFit(text: string, style: Partial<TextStyle>, width: number): Partial<TextStyle> {
    const natural = this.measureText(text, style);
    if (natural <= width || width <= 0) return style;
    const size = style.fontSize ?? this.theme.cellFontSize;
    return { ...style, fontSize: Math.max(size * (width / natural), 1) };
  }

  /**
   * Underline and strike-through lines from `x` over `width`, `baselineY` being the text's baseline: a double underline is two lines
   * and an accounting one runs the full width of the cell's text the same way (the file's accounting variants differ in where the
   * line sits against the cell edge, which the canvas does not draw).
   */
  private drawDecorations(ctx: CanvasRenderingContext2D, style: Partial<TextStyle>, x: number, width: number, baselineY: number, baseline: number): void {
    ctx.lineWidth = 1;
    const line = (y: number): void => {
      ctx.beginPath();
      ctx.moveTo(x, y);
      ctx.lineTo(x + width, y);
      ctx.stroke();
    };
    if (style.textDecoration === 'underline') {
      const double = style.underlineKind === 'double' || style.underlineKind === 'doubleAccounting';
      line(baselineY + 2);
      if (double) line(baselineY + 4);
    }
    if (style.textDecoration === 'line-through' || style.strikethrough) line(baselineY - baseline / 3);
  }

  /**
   * Render text rotated by an angle. Drawn about the cell centre and clipped
   * to the cell bounds — rotated text is not truncated or auto-fitted.
   */
  private renderRotatedText(
    ctx: CanvasRenderingContext2D,
    text: string,
    bounds: Rect,
    style: Partial<TextStyle>,
    metrics: FontMetricsCache,
    rotation: number
  ): void {
    // Canvas rotation is clockwise; a positive user angle should tilt text
    // upward, so negate.
    const rad = (-rotation * Math.PI) / 180;
    const cx = bounds.x + bounds.width / 2;
    const cy = bounds.y + bounds.height / 2;

    ctx.save();
    // Clip so rotated text cannot spill outside the cell.
    ctx.beginPath();
    ctx.rect(bounds.x, bounds.y, bounds.width, bounds.height);
    ctx.clip();

    ctx.translate(cx, cy);
    ctx.rotate(rad);
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText(text, 0, 0);

    if (hasDecoration(style)) {
      const textWidth = ctx.measureText(text).width;
      ctx.strokeStyle = style.color ?? this.theme.cellTextColor;
      // (about the text's centre, as it is drawn: the baseline is the middle, the underline a little below it)
      this.drawDecorations(ctx, style, -textWidth / 2, textWidth, metrics.baseline * 0.4 - 2, metrics.baseline * 1.2 - 6);
    }

    ctx.restore();
  }

  /**
   * Render wrapped text (for cells with text wrapping enabled)
   */
  renderWrappedText(
    ctx: CanvasRenderingContext2D,
    text: string,
    bounds: Rect,
    style: Partial<TextStyle>,
    padding: number = 4
  ): void {
    if (!text) return;
    
    const font = this.buildFontString(style);
    const metrics = this.getFontMetrics(font);
    
    ctx.font = font;
    ctx.fillStyle = style.color ?? this.theme.cellTextColor;
    
    const availableWidth = bounds.width - padding * 2;
    const words = text.split(' ');
    const lines: string[] = [];
    let currentLine = '';
    
    // Word wrap
    for (const word of words) {
      const testLine = currentLine ? `${currentLine} ${word}` : word;
      const testWidth = this.measureCtx.measureText(testLine).width;
      
      if (testWidth <= availableWidth) {
        currentLine = testLine;
      } else {
        if (currentLine) {
          lines.push(currentLine);
        }
        currentLine = word;
      }
    }
    if (currentLine) {
      lines.push(currentLine);
    }
    
    // Calculate starting y position
    const totalTextHeight = lines.length * metrics.lineHeight;
    let startY: number;
    const verticalAlign = style.verticalAlign ?? 'middle';
    
    switch (verticalAlign) {
      case 'top':
        startY = bounds.y + padding + metrics.baseline;
        break;
      case 'bottom':
        startY = bounds.y + bounds.height - padding - totalTextHeight + metrics.baseline;
        break;
      case 'middle':
      default:
        startY = bounds.y + (bounds.height - totalTextHeight) / 2 + metrics.baseline;
        break;
    }
    
    // Calculate x position based on alignment
    const textAlign = style.textAlign ?? 'left';
    let x: number;
    
    switch (textAlign) {
      case 'center':
        ctx.textAlign = 'center';
        x = bounds.x + bounds.width / 2;
        break;
      case 'right':
        ctx.textAlign = 'right';
        x = bounds.x + bounds.width - padding;
        break;
      case 'left':
      default:
        ctx.textAlign = 'left';
        x = bounds.x + padding;
        break;
    }
    
    // Render each line
    let y = startY;
    for (const line of lines) {
      // Only render if within bounds
      if (y + metrics.lineHeight > bounds.y && y - metrics.baseline < bounds.y + bounds.height) {
        ctx.fillText(line, x, y);
      }
      y += metrics.lineHeight;
    }
  }
  
  /**
   * Clear the metrics cache
   */
  clearCache(): void {
    this.metricsCache.clear();
  }
}

