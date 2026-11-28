// Cell Renderer - Handles cell content rendering

import type { CanvasTheme, Rect, TextStyle, CellBorders } from './types';
import type { Cell, CellStyle, CellFormat } from '../types';
import { TextRenderer } from './text-renderer';
import { GridRenderer } from './grid-renderer';
import { formatNumber } from '../utils/format-utils';
import { fillColorOf } from './style-paint';

/**
 * Renders cell content including backgrounds, text, and borders
 */
export class CellRenderer {
  private theme: CanvasTheme;
  private textRenderer: TextRenderer;
  private gridRenderer: GridRenderer;
  
  constructor(theme: CanvasTheme, textRenderer: TextRenderer) {
    this.theme = theme;
    this.textRenderer = textRenderer;
    this.gridRenderer = new GridRenderer(theme);
  }
  
  /**
   * Render a cell with all its content
   */
  renderCell(
    ctx: CanvasRenderingContext2D,
    cell: Cell | undefined,
    bounds: Rect,
    style: CellStyle | undefined,
    format: CellFormat | undefined
  ): void {
    // 1. Render background
    this.renderBackground(ctx, bounds, style);
    
    // 2. Render cell content (text/value)
    if (cell) {
      const displayValue = this.formatCellValue(cell, format);
      if (displayValue) {
        const textStyle = this.cellStyleToTextStyle(style);
        
        // Rotation takes precedence over wrapping; renderText handles rotation.
        if (style?.textWrap && !style?.textRotation) {
          this.textRenderer.renderWrappedText(ctx, displayValue, bounds, textStyle);
        } else {
          this.textRenderer.renderText(ctx, displayValue, bounds, textStyle);
        }
      }
    }
    
    // 3. Render borders (if any)
    if (style) this.renderBorders(ctx, bounds, style);
  }

  private renderBorders(ctx: CanvasRenderingContext2D, bounds: Rect, style: CellStyle): void {
    const borders = this.cellStyleToBorders(style);
    if (borders.top || borders.right || borders.bottom || borders.left) {
      this.gridRenderer.renderCellBorders(ctx, bounds, borders);
    }
  }
  
  /**
   * Render an empty cell (just background)
   */
  renderEmptyCell(ctx: CanvasRenderingContext2D, bounds: Rect, style?: CellStyle): void {
    this.renderBackground(ctx, bounds, style);
    if (style) this.renderBorders(ctx, bounds, style);
  }

  /**
   * Render the comment marker — a small filled triangle in the cell's
   * top-right corner, shown on cells that have an open comment thread.
   */
  renderCommentIndicator(ctx: CanvasRenderingContext2D, bounds: Rect): void {
    const size = 6;
    const x = bounds.x + bounds.width;
    const y = bounds.y;
    ctx.beginPath();
    ctx.moveTo(x - size, y);
    ctx.lineTo(x, y);
    ctx.lineTo(x, y + size);
    ctx.closePath();
    ctx.fillStyle = '#6366f1';
    ctx.fill();
  }
  
  /**
   * Render cell background
   */
  private renderBackground(
    ctx: CanvasRenderingContext2D,
    bounds: Rect,
    style: CellStyle | undefined
  ): void {
    const bgColor = fillColorOf(style) ?? this.theme.cellBackgroundColor;
    
    ctx.fillStyle = bgColor;
    ctx.fillRect(bounds.x, bounds.y, bounds.width, bounds.height);
  }
  
  /**
   * Format cell value for display
   */
  private formatCellValue(cell: Cell, format: CellFormat | undefined): string {
    const value = cell.value;

    if (value === null || value === undefined) {
      return '';
    }

    // If there's a formula, show the computed value (not the formula)
    // The formula is shown in the formula bar, not in the cell
    
    if (typeof value === 'number') {
      return this.formatNumber(value, format);
    }
    
    if (typeof value === 'boolean') {
      return value ? 'TRUE' : 'FALSE';
    }
    
    return String(value);
  }
  
  /**
   * Format a number according to the cell format
   */
  private formatNumber(value: number, format: CellFormat | undefined): string {
    // Use the comprehensive formatNumber function from format-utils
    // Pass an empty format object if none is provided
    return formatNumber(value, format || {});
  }
  
  
  /**
   * Convert CellStyle to TextStyle for rendering
   */
  private cellStyleToTextStyle(style: CellStyle | undefined): Partial<TextStyle> {
    if (!style) {
      return {};
    }
    
    return {
      fontFamily: style.fontFamily,
      fontSize: style.fontSize,
      fontWeight: style.bold ? 'bold' : 'normal',
      fontStyle: style.italic ? 'italic' : 'normal',
      color: style.fontColor,
      textAlign: style.textAlign,
      verticalAlign: style.verticalAlign,
      textDecoration: style.textDecoration,
      underlineKind: style.underlineKind,
      strikethrough: style.strikethrough,
      indent: style.indent,
      // (shrinking is for one line: wrapped text and rotated text keep their size)
      shrinkToFit: style.shrinkToFit && !style.textWrap && !style.textRotation ? true : undefined,
      textRotation: style.textRotation,
    };
  }
  
  /**
   * Convert CellStyle borders to CellBorders
   */
  private cellStyleToBorders(style: CellStyle): CellBorders {
    return {
      top: this.gridRenderer.parseBorderString(style.borderTop, style.borderTopKind),
      right: this.gridRenderer.parseBorderString(style.borderRight, style.borderRightKind),
      bottom: this.gridRenderer.parseBorderString(style.borderBottom, style.borderBottomKind),
      left: this.gridRenderer.parseBorderString(style.borderLeft, style.borderLeftKind),
    };
  }
  
  /**
   * Render a hyperlink cell (with special styling)
   */
  renderHyperlinkCell(
    ctx: CanvasRenderingContext2D,
    cell: Cell,
    bounds: Rect,
    style: CellStyle | undefined,
    format: CellFormat | undefined
  ): void {
    // Render background
    this.renderBackground(ctx, bounds, style);
    
    // Render text with hyperlink styling
    const displayValue = this.formatCellValue(cell, format);
    if (displayValue) {
      const textStyle: Partial<TextStyle> = {
        ...this.cellStyleToTextStyle(style),
        color: '#1a73e8', // Blue color for hyperlinks
        textDecoration: 'underline',
      };
      
      this.textRenderer.renderText(ctx, displayValue, bounds, textStyle);
    }
    
    // Render borders
    if (style) this.renderBorders(ctx, bounds, style);
  }
}

