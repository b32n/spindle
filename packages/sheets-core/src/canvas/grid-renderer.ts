// Grid Renderer - Handles grid lines and cell borders

import type { CanvasTheme, Viewport, BorderStyle, CellBorders, Rect } from './types';
import type { BorderKind } from '../types';
import { BORDER_PAINT, borderKindOf } from './style-paint';

/**
 * Renders grid lines and cell borders
 */
export class GridRenderer {
  private theme: CanvasTheme;
  
  constructor(theme: CanvasTheme) {
    this.theme = theme;
  }
  
  /**
   * Render grid lines for the visible area
   */
  renderGridLines(
    ctx: CanvasRenderingContext2D,
    viewport: Viewport,
    rowHeights: Map<number, number>,
    colWidths: Map<number, number>,
    defaultRowHeight: number,
    defaultColWidth: number,
    headerWidth: number,
    headerHeight: number,
    startRow: number,
    endRow: number,
    startCol: number,
    endCol: number,
    hiddenRows?: Set<number>,
    hiddenCols?: Set<number>,
    // First row/column to accumulate offsets from. For a frozen region's
    // scrolling pane, headerHeight/headerWidth already include the frozen extent,
    // so accumulation must start at the frozen boundary — otherwise the frozen
    // rows/cols are counted twice and the grid lines shift by the frozen extent
    // (misaligned with the cells, which accumulate from the boundary).
    firstAccumRow: number = 0,
    firstAccumCol: number = 0
  ): void {
    const { scrollTop, scrollLeft, width, height } = viewport;
    const hidRows = hiddenRows ?? new Set<number>();
    const hidCols = hiddenCols ?? new Set<number>();

    ctx.strokeStyle = this.theme.gridLineColor;
    ctx.lineWidth = this.theme.gridLineWidth;

    // Calculate starting positions (skip hidden rows/cols)
    let startY = headerHeight;
    for (let r = firstAccumRow; r < startRow; r++) {
      if (!hidRows.has(r)) {
        startY += rowHeights.get(r) ?? defaultRowHeight;
      }
    }
    startY -= scrollTop;

    let startX = headerWidth;
    for (let c = firstAccumCol; c < startCol; c++) {
      if (!hidCols.has(c)) {
        startX += colWidths.get(c) ?? defaultColWidth;
      }
    }
    startX -= scrollLeft;
    
    ctx.beginPath();
    
    // Draw horizontal lines (row separators) - skip hidden rows
    let y = startY;
    for (let row = startRow; row <= endRow; row++) {
      if (hidRows.has(row)) continue;
      // Snap to pixel grid for crisp lines
      const lineY = Math.round(y) + 0.5;
      ctx.moveTo(headerWidth, lineY);
      ctx.lineTo(width, lineY);
      y += rowHeights.get(row) ?? defaultRowHeight;
    }
    
    // Draw vertical lines (column separators) - skip hidden columns
    let x = startX;
    for (let col = startCol; col <= endCol; col++) {
      if (hidCols.has(col)) continue;
      // Snap to pixel grid for crisp lines
      const lineX = Math.round(x) + 0.5;
      ctx.moveTo(lineX, headerHeight);
      ctx.lineTo(lineX, height);
      x += colWidths.get(col) ?? defaultColWidth;
    }
    
    ctx.stroke();
  }
  
  /**
   * Parse a CSS border string into a BorderStyle object
   */
  parseBorderString(border: string | undefined, kind?: BorderKind): BorderStyle | undefined {
    if (!border || border === 'none') {
      return undefined;
    }
    
    // Parse formats like "1px solid #000000" or "2px dashed red"
    const parts = border.split(' ');
    if (parts.length < 2) {
      return undefined;
    }
    
    const widthStr = parts[0];
    const style = parts[1] as 'solid' | 'dashed' | 'dotted';
    const color = parts[2] ?? '#000000';
    
    const width = parseInt(widthStr, 10);
    if (isNaN(width)) {
      return undefined;
    }
    
    return { width, style, color, kind: borderKindOf(border, kind) };
  }
  
  /**
   * Render borders for a single cell
   */
  renderCellBorders(
    ctx: CanvasRenderingContext2D,
    bounds: Rect,
    borders: CellBorders
  ): void {
    const { x, y, width, height } = bounds;
    
    // Render each border if specified
    if (borders.top) {
      this.renderBorderLine(ctx, x, y, x + width, y, borders.top);
    }
    
    if (borders.right) {
      this.renderBorderLine(ctx, x + width, y, x + width, y + height, borders.right);
    }
    
    if (borders.bottom) {
      this.renderBorderLine(ctx, x, y + height, x + width, y + height, borders.bottom);
    }
    
    if (borders.left) {
      this.renderBorderLine(ctx, x, y, x, y + height, borders.left);
    }
  }
  
  /**
   * Render a single border line with the specified style
   */
  private renderBorderLine(
    ctx: CanvasRenderingContext2D,
    x1: number,
    y1: number,
    x2: number,
    y2: number,
    style: BorderStyle
  ): void {
    ctx.strokeStyle = style.color;

    // The stroke of the exact line style when the file had one, else of the CSS style
    const paint = style.kind ? BORDER_PAINT[style.kind] : undefined;
    const width = paint ? paint.width : style.width;
    ctx.setLineDash(paint ? [...paint.dash] : style.style === 'dashed' ? [4, 2] : style.style === 'dotted' ? [1, 1] : []);

    const horizontal = y1 === y2;
    // A double line is two thin lines either side of the edge, the way Excel draws it
    const offsets = paint?.double || style.style === 'double' ? [-1, 1] : [0];
    ctx.lineWidth = offsets.length > 1 ? 1 : width;

    for (const offset of offsets) {
      ctx.beginPath();

      // Snap to pixel grid for crisp lines
      const snapX1 = Math.round(x1) + 0.5 + (horizontal ? 0 : offset);
      const snapY1 = Math.round(y1) + 0.5 + (horizontal ? offset : 0);
      const snapX2 = Math.round(x2) + 0.5 + (horizontal ? 0 : offset);
      const snapY2 = Math.round(y2) + 0.5 + (horizontal ? offset : 0);

      ctx.moveTo(snapX1, snapY1);
      ctx.lineTo(snapX2, snapY2);
      ctx.stroke();
    }
    
    // Reset line dash
    ctx.setLineDash([]);
  }
  
  /**
   * Render borders for all visible cells that have custom borders
   */
  renderAllCellBorders(
    ctx: CanvasRenderingContext2D,
    cellBorders: Map<string, CellBorders>,
    viewport: Viewport,
    rowHeights: Map<number, number>,
    colWidths: Map<number, number>,
    defaultRowHeight: number,
    defaultColWidth: number,
    headerWidth: number,
    headerHeight: number,
    startRow: number,
    endRow: number,
    startCol: number,
    endCol: number
  ): void {
    const { scrollTop, scrollLeft } = viewport;
    
    // Calculate starting positions
    let startY = headerHeight;
    for (let r = 0; r < startRow; r++) {
      startY += rowHeights.get(r) ?? defaultRowHeight;
    }
    startY -= scrollTop;
    
    let startX = headerWidth;
    for (let c = 0; c < startCol; c++) {
      startX += colWidths.get(c) ?? defaultColWidth;
    }
    startX -= scrollLeft;
    
    // Iterate through visible cells and render borders
    let y = startY;
    for (let row = startRow; row < endRow; row++) {
      const rowHeight = rowHeights.get(row) ?? defaultRowHeight;
      let x = startX;
      
      for (let col = startCol; col < endCol; col++) {
        const colWidth = colWidths.get(col) ?? defaultColWidth;
        const cellKey = `${row}:${col}`;
        const borders = cellBorders.get(cellKey);
        
        if (borders) {
          this.renderCellBorders(ctx, { x, y, width: colWidth, height: rowHeight }, borders);
        }
        
        x += colWidth;
      }
      
      y += rowHeight;
    }
  }
}

