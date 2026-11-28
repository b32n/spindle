// Main Canvas Renderer - Coordinates all sub-renderers

import type {
  CanvasRendererConfig,
  RenderState,
  Viewport,
  CellPosition,
  DirtyRegion,
  CanvasTheme,
  Rect,
} from './types';
import { DEFAULT_THEME } from './types';
import { TextRenderer } from './text-renderer';
import { GridRenderer } from './grid-renderer';
import { CellRenderer } from './cell-renderer';
import { HeaderRenderer } from './header-renderer';
import { SelectionRenderer } from './selection-renderer';
import { HitTester } from './hit-testing';
import {
  calculateFreezeDimensions,
  type FreezeRegion,
  type FreezeDimensions
} from '../features/freeze';

/** Inputs for the shared visible-range computation. */
interface VisibleRangeParams {
  scrollTop: number;
  scrollLeft: number;
  width: number;
  height: number;
  headerWidth: number;
  headerHeight: number;
  frozenWidth: number;
  frozenHeight: number;
  frozenRows: number;
  frozenCols: number;
  rowCount: number;
  colCount: number;
  rowHeights: Map<number, number>;
  colWidths: Map<number, number>;
  defaultRowHeight: number;
  defaultColWidth: number;
  hiddenRows: Set<number>;
  hiddenCols: Set<number>;
}

/**
 * Single source of truth for the visible cell range. Both the render path
 * (calculateVisibleRange) and the cell-load path (calculateVisibleRangeForDimensions)
 * delegate here so the window that is loaded exactly matches the window that is
 * drawn — divergence between the two used to leave visible cells blank because
 * one path skipped hidden/frozen rows and the other did not.
 */
export function computeVisibleRange(
  p: VisibleRangeParams
): { startRow: number; endRow: number; startCol: number; endCol: number } {
  // Visible rows (skip hidden; frozen rows don't scroll so start past them)
  let y = 0;
  let startRow = p.frozenRows;
  while (y < p.scrollTop && startRow < p.rowCount) {
    if (!p.hiddenRows.has(startRow)) y += p.rowHeights.get(startRow) ?? p.defaultRowHeight;
    startRow++;
  }
  if (startRow > p.frozenRows) startRow--; // include partially visible row

  let endRow = startRow;
  const visibleHeight = p.height - p.headerHeight - p.frozenHeight;
  while (y < p.scrollTop + visibleHeight && endRow < p.rowCount) {
    if (!p.hiddenRows.has(endRow)) y += p.rowHeights.get(endRow) ?? p.defaultRowHeight;
    endRow++;
  }
  endRow = Math.min(endRow + 1, p.rowCount); // include partially visible row

  // Visible columns (skip hidden; frozen cols don't scroll)
  let x = 0;
  let startCol = p.frozenCols;
  while (x < p.scrollLeft && startCol < p.colCount) {
    if (!p.hiddenCols.has(startCol)) x += p.colWidths.get(startCol) ?? p.defaultColWidth;
    startCol++;
  }
  if (startCol > p.frozenCols) startCol--; // include partially visible column

  let endCol = startCol;
  const visibleWidth = p.width - p.headerWidth - p.frozenWidth;
  while (x < p.scrollLeft + visibleWidth && endCol < p.colCount) {
    if (!p.hiddenCols.has(endCol)) x += p.colWidths.get(endCol) ?? p.defaultColWidth;
    endCol++;
  }
  endCol = Math.min(endCol + 1, p.colCount); // include partially visible column

  return { startRow, endRow, startCol, endCol };
}

/**
 * Main canvas renderer that coordinates all rendering operations
 */
export class CanvasRenderer {
  private canvas: HTMLCanvasElement;
  private ctx: CanvasRenderingContext2D;
  private devicePixelRatio: number;
  
  // Configuration
  private defaultRowHeight: number;
  private defaultColWidth: number;
  private headerHeight: number;
  private headerWidth: number;
  private theme: CanvasTheme;
  
  // Viewport state
  private viewport: Viewport;
  
  // Current render state
  private renderState: RenderState | null = null;
  // Cells covered by a merged region ("row:col"), rebuilt each frame. Used to
  // skip them in the cell loop — the merged-region pass draws them instead.
  private mergedCellSet: Set<string> = new Set();

  // Sub-renderers
  private textRenderer: TextRenderer;
  private gridRenderer: GridRenderer;
  private cellRenderer: CellRenderer;
  private headerRenderer: HeaderRenderer;
  private selectionRenderer: SelectionRenderer;
  
  // Hit tester
  private hitTester: HitTester;
  
  // Dirty tracking
  private isDirty: boolean = true;
  private dirtyRegions: DirtyRegion[] = [];
  
  // Animation frame handle
  private animationFrameId: number | null = null;
  
  // Freeze panes state
  private freezeDimensions: FreezeDimensions = { frozenWidth: 0, frozenHeight: 0 };
  private frozenRows: number = 0;
  private frozenCols: number = 0;
  
  constructor(config: CanvasRendererConfig) {
    this.canvas = config.canvas;
    const ctx = this.canvas.getContext('2d');
    if (!ctx) {
      throw new Error('Failed to get 2D context from canvas');
    }
    this.ctx = ctx;
    
    this.devicePixelRatio = config.devicePixelRatio ?? window.devicePixelRatio ?? 1;
    this.defaultRowHeight = config.defaultRowHeight;
    this.defaultColWidth = config.defaultColWidth;
    this.headerHeight = config.headerHeight;
    this.headerWidth = config.headerWidth;
    this.theme = { ...DEFAULT_THEME, ...config.theme };
    
    // Initialize viewport
    this.viewport = {
      scrollTop: 0,
      scrollLeft: 0,
      width: this.canvas.width / this.devicePixelRatio,
      height: this.canvas.height / this.devicePixelRatio,
      startRow: 0,
      endRow: 0,
      startCol: 0,
      endCol: 0,
    };
    
    // Initialize sub-renderers
    this.textRenderer = new TextRenderer(this.theme);
    this.gridRenderer = new GridRenderer(this.theme);
    this.cellRenderer = new CellRenderer(this.theme, this.textRenderer);
    this.headerRenderer = new HeaderRenderer(this.theme);
    this.selectionRenderer = new SelectionRenderer(this.theme);
    
    // Initialize hit tester
    this.hitTester = new HitTester(
      this.headerWidth,
      this.headerHeight,
      this.defaultRowHeight,
      this.defaultColWidth
    );
    
    // Set up canvas for high DPI
    this.setupCanvas();
  }
  
  /**
   * Set up canvas for high DPI displays
   */
  private setupCanvas(): void {
    const { width, height } = this.canvas.getBoundingClientRect();
    
    // Set actual size in memory (scaled for device pixel ratio)
    this.canvas.width = width * this.devicePixelRatio;
    this.canvas.height = height * this.devicePixelRatio;
    
    // Scale the context to match device pixel ratio
    this.ctx.scale(this.devicePixelRatio, this.devicePixelRatio);
    
    // Update viewport dimensions
    this.viewport.width = width;
    this.viewport.height = height;
  }
  
  /**
   * Resize the canvas
   */
  resize(width: number, height: number): void {
    this.canvas.style.width = `${width}px`;
    this.canvas.style.height = `${height}px`;
    this.canvas.width = width * this.devicePixelRatio;
    this.canvas.height = height * this.devicePixelRatio;
    
    this.ctx.scale(this.devicePixelRatio, this.devicePixelRatio);
    
    this.viewport.width = width;
    this.viewport.height = height;
    
    this.invalidate();
  }
  
  /**
   * Set the viewport scroll position
   */
  setViewport(scrollTop: number, scrollLeft: number): void {
    if (this.viewport.scrollTop === scrollTop && this.viewport.scrollLeft === scrollLeft) {
      return;
    }
    
    this.viewport.scrollTop = scrollTop;
    this.viewport.scrollLeft = scrollLeft;
    
    // Update hit tester
    this.hitTester.setScroll(scrollTop, scrollLeft);
    
    // Recalculate visible range if we have render state
    if (this.renderState) {
      this.calculateVisibleRange(this.renderState);
    }
    
    this.invalidate();
  }
  
  /**
   * Calculate the visible range based on current viewport and optionally provided dimensions.
   * This can be called before setState to get an accurate viewport for cell loading.
   */
  calculateVisibleRangeForDimensions(
    rowCount: number,
    colCount: number,
    rowHeights?: Map<number, number>,
    colWidths?: Map<number, number>,
    hiddenRows?: Set<number>,
    hiddenCols?: Set<number>,
    frozenRows: number = 0,
    frozenCols: number = 0
  ): void {
    const rh = rowHeights ?? new Map<number, number>();
    const cw = colWidths ?? new Map<number, number>();
    const hr = hiddenRows ?? new Set<number>();
    const hc = hiddenCols ?? new Set<number>();
    // Freeze offsets must match the render path, otherwise the loaded window
    // shifts relative to the drawn one.
    const freeze = calculateFreezeDimensions(
      frozenRows,
      frozenCols,
      rh,
      cw,
      this.defaultRowHeight,
      this.defaultColWidth,
      hr,
      hc
    );
    const r = computeVisibleRange({
      scrollTop: this.viewport.scrollTop,
      scrollLeft: this.viewport.scrollLeft,
      width: this.viewport.width,
      height: this.viewport.height,
      headerWidth: this.headerWidth,
      headerHeight: this.headerHeight,
      frozenWidth: freeze.frozenWidth,
      frozenHeight: freeze.frozenHeight,
      frozenRows,
      frozenCols,
      rowCount,
      colCount,
      rowHeights: rh,
      colWidths: cw,
      defaultRowHeight: this.defaultRowHeight,
      defaultColWidth: this.defaultColWidth,
      hiddenRows: hr,
      hiddenCols: hc,
    });
    this.viewport.startRow = r.startRow;
    this.viewport.endRow = r.endRow;
    this.viewport.startCol = r.startCol;
    this.viewport.endCol = r.endCol;
  }
  
  /**
   * Update the render state
   */
  setState(state: RenderState): void {
    this.renderState = state;
    
    // Calculate freeze dimensions
    this.frozenRows = state.frozenRows ?? 0;
    this.frozenCols = state.frozenCols ?? 0;
    this.freezeDimensions = calculateFreezeDimensions(
      this.frozenRows,
      this.frozenCols,
      state.rowHeights,
      state.colWidths,
      this.defaultRowHeight,
      this.defaultColWidth,
      state.hiddenRows,
      state.hiddenCols
    );
    
    // Update hit tester with row/col dimensions, hidden rows/cols, and freeze config
    this.hitTester.setDimensions(
      state.rowHeights,
      state.colWidths,
      this.defaultRowHeight,
      this.defaultColWidth,
      state.rowCount,
      state.colCount,
      state.hiddenRows,
      state.hiddenCols
    );
    
    // Update hit tester with freeze config
    this.hitTester.setFreezeConfig(
      this.frozenRows,
      this.frozenCols,
      this.freezeDimensions.frozenWidth,
      this.freezeDimensions.frozenHeight
    );
    
    // Calculate visible range immediately so getViewport() returns correct values
    this.calculateVisibleRange(state);
    
    this.invalidate();
  }
  
  /**
   * Mark the canvas as needing redraw
   */
  invalidate(region?: DirtyRegion): void {
    this.isDirty = true;
    if (region) {
      this.dirtyRegions.push(region);
    } else {
      this.dirtyRegions = [{ type: 'all' }];
    }
    
    // Schedule a render on the next animation frame
    if (this.animationFrameId === null) {
      this.animationFrameId = requestAnimationFrame(() => {
        this.animationFrameId = null;
        this.render();
      });
    }
  }
  
  /**
   * Force an immediate render
   */
  renderNow(): void {
    if (this.animationFrameId !== null) {
      cancelAnimationFrame(this.animationFrameId);
      this.animationFrameId = null;
    }
    this.render();
  }
  
  /**
   * Main render method
   */
  private render(): void {
    if (!this.isDirty) {
      return;
    }

    if (!this.renderState) {
      return;
    }
    
    const state = this.renderState;
    const { frozenWidth, frozenHeight } = this.freezeDimensions;
    const hasFrozenRows = this.frozenRows > 0;
    const hasFrozenCols = this.frozenCols > 0;

    // Rebuild the covered-cell lookup for this frame.
    this.mergedCellSet.clear();
    if (state.mergedRegions) {
      for (const region of state.mergedRegions) {
        for (let r = region.startRow; r <= region.endRow; r++) {
          for (let c = region.startCol; c <= region.endCol; c++) {
            this.mergedCellSet.add(`${r}:${c}`);
          }
        }
      }
    }

    // Calculate visible range
    this.calculateVisibleRange(state);
    
    // Clear the canvas
    this.ctx.clearRect(0, 0, this.viewport.width, this.viewport.height);
    
    // Save context state
    this.ctx.save();
    
    // Render in order (back to front):
    // For freeze panes, we render 4 regions in order:
    // 1. Main scrollable area (scrolls both ways)
    // 2. Left frozen column area (scrolls vertically only)
    // 3. Top frozen row area (scrolls horizontally only)
    // 4. Top-left corner (never scrolls)
    
    // --- MAIN SCROLLABLE AREA ---
    this.renderRegion(state, 'main');
    
    // --- LEFT FROZEN AREA (if frozen cols exist) ---
    if (hasFrozenCols) {
      this.renderRegion(state, 'left');
    }
    
    // --- TOP FROZEN AREA (if frozen rows exist) ---
    if (hasFrozenRows) {
      this.renderRegion(state, 'top');
    }
    
    // --- TOP-LEFT FROZEN CORNER (if both frozen rows and cols exist) ---
    if (hasFrozenRows && hasFrozenCols) {
      this.renderRegion(state, 'top-left');
    }
    
    // --- FREEZE DIVIDER LINES ---
    if (hasFrozenRows || hasFrozenCols) {
      this.renderFreezeDividers(frozenWidth, frozenHeight);
    }
    
    // --- HEADERS ---
    this.renderHeaders(state);
    
    // --- CORNER CELL ---
    this.renderCornerCell();
    
    // Restore context state
    this.ctx.restore();
    
    // Reset dirty state
    this.isDirty = false;
    this.dirtyRegions = [];
  }
  
  /**
   * Render a specific freeze region (cells, grid lines, selection)
   */
  private renderRegion(state: RenderState, region: FreezeRegion): void {
    const { frozenWidth, frozenHeight } = this.freezeDimensions;
    const { scrollTop, scrollLeft, width, height } = this.viewport;
    
    // Calculate clip rect based on region
    let clipX: number, clipY: number, clipW: number, clipH: number;
    let effectiveScrollTop: number, effectiveScrollLeft: number;
    let startRow: number, endRow: number, startCol: number, endCol: number;
    
    switch (region) {
      case 'top-left':
        // Never scrolls, shows frozen rows and cols
        clipX = this.headerWidth;
        clipY = this.headerHeight;
        clipW = frozenWidth;
        clipH = frozenHeight;
        effectiveScrollTop = 0;
        effectiveScrollLeft = 0;
        startRow = 0;
        endRow = this.frozenRows;
        startCol = 0;
        endCol = this.frozenCols;
        break;
        
      case 'top':
        // Scrolls horizontally only, shows frozen rows
        clipX = this.headerWidth + frozenWidth;
        clipY = this.headerHeight;
        clipW = width - this.headerWidth - frozenWidth;
        clipH = frozenHeight;
        effectiveScrollTop = 0;
        effectiveScrollLeft = scrollLeft;
        startRow = 0;
        endRow = this.frozenRows;
        startCol = this.viewport.startCol;
        endCol = this.viewport.endCol;
        // Ensure we don't render frozen cols in this region
        if (startCol < this.frozenCols) startCol = this.frozenCols;
        break;
        
      case 'left':
        // Scrolls vertically only, shows frozen cols
        clipX = this.headerWidth;
        clipY = this.headerHeight + frozenHeight;
        clipW = frozenWidth;
        clipH = height - this.headerHeight - frozenHeight;
        effectiveScrollTop = scrollTop;
        effectiveScrollLeft = 0;
        startRow = this.viewport.startRow;
        endRow = this.viewport.endRow;
        startCol = 0;
        endCol = this.frozenCols;
        // Ensure we don't render frozen rows in this region
        if (startRow < this.frozenRows) startRow = this.frozenRows;
        break;
        
      case 'main':
      default:
        // Scrolls both ways
        clipX = this.headerWidth + frozenWidth;
        clipY = this.headerHeight + frozenHeight;
        clipW = width - this.headerWidth - frozenWidth;
        clipH = height - this.headerHeight - frozenHeight;
        effectiveScrollTop = scrollTop;
        effectiveScrollLeft = scrollLeft;
        startRow = this.viewport.startRow;
        endRow = this.viewport.endRow;
        startCol = this.viewport.startCol;
        endCol = this.viewport.endCol;
        // Ensure we don't render frozen rows/cols in main region
        if (startRow < this.frozenRows) startRow = this.frozenRows;
        if (startCol < this.frozenCols) startCol = this.frozenCols;
        break;
    }
    
    // Skip if clip region is invalid
    if (clipW <= 0 || clipH <= 0) return;
    
    // Render cells, grid lines, and selection for this region
    this.ctx.save();
    this.ctx.beginPath();
    this.ctx.rect(clipX, clipY, clipW, clipH);
    this.ctx.clip();
    
    // Render cells
    this.renderCellsInRegion(
      state, 
      startRow, endRow, startCol, endCol,
      effectiveScrollTop, effectiveScrollLeft,
      region
    );
    
    // Render grid lines
    this.renderGridLinesInRegion(
      state,
      startRow, endRow, startCol, endCol,
      effectiveScrollTop, effectiveScrollLeft,
      region
    );

    // Render merged regions on top of the grid lines (covers interior lines).
    this.renderMergedRegions(state);

    // Render selection
    this.renderSelectionInRegion(
      state,
      effectiveScrollTop, effectiveScrollLeft,
      region
    );
    
    this.ctx.restore();
  }
  
  /**
   * Render freeze divider lines
   */
  private renderFreezeDividers(frozenWidth: number, frozenHeight: number): void {
    const { width, height } = this.viewport;
    
    this.ctx.save();
    this.ctx.strokeStyle = this.theme.freezeDividerColor;
    this.ctx.lineWidth = this.theme.freezeDividerWidth;
    
    // Vertical divider (after frozen columns)
    if (this.frozenCols > 0 && frozenWidth > 0) {
      const dividerX = this.headerWidth + frozenWidth;
      this.ctx.beginPath();
      this.ctx.moveTo(dividerX, this.headerHeight);
      this.ctx.lineTo(dividerX, height);
      this.ctx.stroke();
      
      // Add subtle shadow
      this.ctx.fillStyle = this.theme.freezeShadowColor;
      this.ctx.fillRect(dividerX, this.headerHeight, 4, height - this.headerHeight);
    }
    
    // Horizontal divider (after frozen rows)
    if (this.frozenRows > 0 && frozenHeight > 0) {
      const dividerY = this.headerHeight + frozenHeight;
      this.ctx.beginPath();
      this.ctx.moveTo(this.headerWidth, dividerY);
      this.ctx.lineTo(width, dividerY);
      this.ctx.stroke();
      
      // Add subtle shadow
      this.ctx.fillStyle = this.theme.freezeShadowColor;
      this.ctx.fillRect(this.headerWidth, dividerY, width - this.headerWidth, 4);
    }
    
    this.ctx.restore();
  }
  
  /**
   * Calculate which rows and columns are visible
   */
  private calculateVisibleRange(state: RenderState): void {
    const r = computeVisibleRange({
      scrollTop: this.viewport.scrollTop,
      scrollLeft: this.viewport.scrollLeft,
      width: this.viewport.width,
      height: this.viewport.height,
      headerWidth: this.headerWidth,
      headerHeight: this.headerHeight,
      frozenWidth: this.freezeDimensions.frozenWidth,
      frozenHeight: this.freezeDimensions.frozenHeight,
      frozenRows: this.frozenRows,
      frozenCols: this.frozenCols,
      rowCount: state.rowCount,
      colCount: state.colCount,
      rowHeights: state.rowHeights,
      colWidths: state.colWidths,
      defaultRowHeight: this.defaultRowHeight,
      defaultColWidth: this.defaultColWidth,
      hiddenRows: state.hiddenRows ?? new Set<number>(),
      hiddenCols: state.hiddenCols ?? new Set<number>(),
    });
    this.viewport.startRow = r.startRow;
    this.viewport.endRow = r.endRow;
    this.viewport.startCol = r.startCol;
    this.viewport.endCol = r.endCol;
  }
  
  /**
   * Render cells in a specific region
   */
  private renderCellsInRegion(
    state: RenderState,
    startRow: number,
    endRow: number,
    startCol: number,
    endCol: number,
    effectiveScrollTop: number,
    effectiveScrollLeft: number,
    region: FreezeRegion
  ): void {
    const hiddenRows = state.hiddenRows ?? new Set<number>();
    const hiddenCols = state.hiddenCols ?? new Set<number>();
    const { frozenWidth, frozenHeight } = this.freezeDimensions;
    
    // Calculate starting Y position based on region
    let startY: number;
    if (region === 'top-left' || region === 'top') {
      // Frozen rows - start at header
      startY = this.headerHeight;
      for (let r = 0; r < startRow; r++) {
        if (!hiddenRows.has(r)) {
          startY += state.rowHeights.get(r) ?? this.defaultRowHeight;
        }
      }
    } else {
      // Non-frozen rows - account for frozen area and scroll
      startY = this.headerHeight + frozenHeight;
      for (let r = this.frozenRows; r < startRow; r++) {
        if (!hiddenRows.has(r)) {
          startY += state.rowHeights.get(r) ?? this.defaultRowHeight;
        }
      }
      startY -= effectiveScrollTop;
    }
    
    // Calculate starting X position based on region
    let startX: number;
    if (region === 'top-left' || region === 'left') {
      // Frozen cols - start at header
      startX = this.headerWidth;
      for (let c = 0; c < startCol; c++) {
        if (!hiddenCols.has(c)) {
          startX += state.colWidths.get(c) ?? this.defaultColWidth;
        }
      }
    } else {
      // Non-frozen cols - account for frozen area and scroll
      startX = this.headerWidth + frozenWidth;
      for (let c = this.frozenCols; c < startCol; c++) {
        if (!hiddenCols.has(c)) {
          startX += state.colWidths.get(c) ?? this.defaultColWidth;
        }
      }
      startX -= effectiveScrollLeft;
    }
    
    // Render visible cells (skip hidden and filtered rows/cols)
    let y = startY;
    for (let row = startRow; row < endRow; row++) {
      // Skip hidden rows
      if (hiddenRows.has(row)) continue;
      // Skip filtered rows (only show rows that pass all filters)
      if (state.filteredRows && !state.filteredRows.has(row)) continue;
      
      const rowHeight = state.rowHeights.get(row) ?? this.defaultRowHeight;
      let x = startX;
      
      for (let col = startCol; col < endCol; col++) {
        // Skip hidden columns
        if (hiddenCols.has(col)) continue;
        
        const colWidth = state.colWidths.get(col) ?? this.defaultColWidth;
        const cellKey = `${row}:${col}`;

        // Cells inside a merged region are drawn by the merged-region pass.
        if (this.mergedCellSet.has(cellKey)) {
          x += colWidth;
          continue;
        }

        const cell = state.cells.get(cellKey);
        // an empty cell shows the default format of its row, or else of its column (a filled one has its own style)
        const style = cell?.styleId
          ? state.styles.get(cell.styleId)
          : cell === undefined
            ? (state.rowFormats?.get(row) ?? state.colFormats?.get(col))?.style
            : undefined;
        const format = cell?.formatId ? state.formats.get(cell.formatId) : undefined;

        const bounds: Rect = { x, y, width: colWidth, height: rowHeight };

        // Skip rendering if this is the editing cell
        const isEditing = state.editingCell?.row === row && state.editingCell?.col === col;
        if (!isEditing) {
          this.cellRenderer.renderCell(this.ctx, cell, bounds, style, format);
        } else {
          // Render empty cell background for editing cell
          this.cellRenderer.renderEmptyCell(this.ctx, bounds);
        }

        // Comment marker — drawn last so it sits above cell content.
        if (state.commentedCells?.has(cellKey)) {
          this.cellRenderer.renderCommentIndicator(this.ctx, bounds);
        }

        x += colWidth;
      }
      
      y += rowHeight;
    }
  }
  
  /**
   * Draw merged regions on top of the grid lines — the union background plus
   * the anchor cell's content spanning the whole region. Called inside each
   * freeze region's clip; getCellBounds positions each region correctly and
   * the clip restricts which ones are visible.
   */
  private renderMergedRegions(state: RenderState): void {
    const regions = state.mergedRegions;
    if (!regions || regions.length === 0) return;

    for (const region of regions) {
      const tl = this.getCellBounds(region.startRow, region.startCol);
      const br = this.getCellBounds(region.endRow, region.endCol);
      if (!tl || !br) continue;

      const bounds: Rect = {
        x: tl.x,
        y: tl.y,
        width: br.x + br.width - tl.x,
        height: br.y + br.height - tl.y,
      };
      if (bounds.width <= 0 || bounds.height <= 0) continue;

      const cellKey = `${region.startRow}:${region.startCol}`;
      const cell = state.cells.get(cellKey);
      const isEditing =
        state.editingCell?.row === region.startRow &&
        state.editingCell?.col === region.startCol;

      if (isEditing) {
        this.cellRenderer.renderEmptyCell(this.ctx, bounds);
      } else {
        const style = cell?.styleId ? state.styles.get(cell.styleId) : undefined;
        const format = cell?.formatId ? state.formats.get(cell.formatId) : undefined;
        this.cellRenderer.renderCell(this.ctx, cell, bounds, style, format);
      }

      if (state.commentedCells?.has(cellKey)) {
        this.cellRenderer.renderCommentIndicator(this.ctx, bounds);
      }
    }
  }

  /**
   * Render grid lines in a specific region
   */
  private renderGridLinesInRegion(
    state: RenderState,
    startRow: number,
    endRow: number,
    startCol: number,
    endCol: number,
    effectiveScrollTop: number,
    effectiveScrollLeft: number,
    region: FreezeRegion
  ): void {
    const { frozenWidth, frozenHeight } = this.freezeDimensions;
    
    // Create a modified viewport for this region
    const regionViewport: Viewport = {
      ...this.viewport,
      scrollTop: effectiveScrollTop,
      scrollLeft: effectiveScrollLeft,
      startRow,
      endRow,
      startCol,
      endCol,
    };
    
    // Determine header offset based on region
    let headerWidthOffset = this.headerWidth;
    let headerHeightOffset = this.headerHeight;
    
    // When the region's header offset already includes the frozen extent, the
    // grid-line accumulation must start at the frozen boundary (matching how the
    // cells are positioned) rather than at index 0.
    let firstAccumRow = 0;
    let firstAccumCol = 0;
    if (region === 'main' || region === 'top') {
      headerWidthOffset = this.headerWidth + frozenWidth;
      firstAccumCol = this.frozenCols;
    }
    if (region === 'main' || region === 'left') {
      headerHeightOffset = this.headerHeight + frozenHeight;
      firstAccumRow = this.frozenRows;
    }

    this.gridRenderer.renderGridLines(
      this.ctx,
      regionViewport,
      state.rowHeights,
      state.colWidths,
      this.defaultRowHeight,
      this.defaultColWidth,
      headerWidthOffset,
      headerHeightOffset,
      startRow,
      endRow,
      startCol,
      endCol,
      state.hiddenRows,
      state.hiddenCols,
      firstAccumRow,
      firstAccumCol
    );
  }
  
  /**
   * Render selection in a specific region
   */
  private renderSelectionInRegion(
    state: RenderState,
    effectiveScrollTop: number,
    effectiveScrollLeft: number,
    region: FreezeRegion
  ): void {
    if (!state.selection && !state.activeCell && (!state.formulaRanges || state.formulaRanges.length === 0)) {
      return;
    }
    
    const { frozenWidth, frozenHeight } = this.freezeDimensions;
    
    // Create a modified viewport for this region
    const regionViewport: Viewport = {
      ...this.viewport,
      scrollTop: effectiveScrollTop,
      scrollLeft: effectiveScrollLeft,
    };
    
    // Determine header offset based on region
    let headerWidthOffset = this.headerWidth;
    let headerHeightOffset = this.headerHeight;
    
    if (region === 'main' || region === 'top') {
      headerWidthOffset = this.headerWidth + frozenWidth;
    }
    if (region === 'main' || region === 'left') {
      headerHeightOffset = this.headerHeight + frozenHeight;
    }
    
    this.selectionRenderer.render(
      this.ctx,
      state.selection,
      state.activeCell,
      regionViewport,
      state.rowHeights,
      state.colWidths,
      this.defaultRowHeight,
      this.defaultColWidth,
      headerWidthOffset,
      headerHeightOffset,
      state.formulaRanges,
      state.hiddenRows,
      state.hiddenCols,
      this.frozenRows,
      this.frozenCols,
      region
    );
  }
  
  
  /**
   * Render row and column headers
   */
  private renderHeaders(state: RenderState): void {
    const { startRow, endRow, startCol, endCol, scrollTop, scrollLeft } = this.viewport;
    const { frozenWidth, frozenHeight } = this.freezeDimensions;
    const hasFrozenRows = this.frozenRows > 0;
    const hasFrozenCols = this.frozenCols > 0;

    // --- COLUMN HEADERS ---
    if (hasFrozenCols) {
      // Render frozen column headers (don't scroll horizontally)
      this.ctx.save();
      this.ctx.beginPath();
      this.ctx.rect(this.headerWidth, 0, frozenWidth, this.headerHeight);
      this.ctx.clip();

      this.headerRenderer.renderColumnHeaders(
        this.ctx,
        0,
        this.frozenCols,
        0, // No horizontal scroll for frozen columns
        state.colWidths,
        this.defaultColWidth,
        this.headerWidth,
        this.headerHeight,
        state.selection,
        state.activeCell,
        state.hiddenCols,
        undefined, // No skip - render from column 0
        state.filters
      );

      this.ctx.restore();
    }

    // Render scrollable column headers
    this.ctx.save();
    this.ctx.beginPath();
    const colHeaderX = this.headerWidth + frozenWidth;
    const colHeaderWidth = this.viewport.width - this.headerWidth - frozenWidth;
    this.ctx.rect(colHeaderX, 0, colHeaderWidth, this.headerHeight);
    this.ctx.clip();
    
    this.headerRenderer.renderColumnHeaders(
      this.ctx,
      hasFrozenCols ? this.frozenCols : startCol,
      endCol,
      scrollLeft,
      state.colWidths,
      this.defaultColWidth,
      this.headerWidth + frozenWidth, // Start after frozen area
      this.headerHeight,
      state.selection,
      state.activeCell,
      state.hiddenCols,
      hasFrozenCols ? this.frozenCols : undefined, // Skip accumulating frozen column widths
      state.filters
    );

    this.ctx.restore();

    // --- ROW HEADERS ---
    if (hasFrozenRows) {
      // Render frozen row headers (don't scroll vertically)
      this.ctx.save();
      this.ctx.beginPath();
      this.ctx.rect(0, this.headerHeight, this.headerWidth, frozenHeight);
      this.ctx.clip();

      this.headerRenderer.renderRowHeaders(
        this.ctx,
        0,
        this.frozenRows,
        0, // No vertical scroll for frozen rows
        state.rowHeights,
        this.defaultRowHeight,
        this.headerWidth,
        this.headerHeight,
        state.selection,
        state.activeCell,
        state.hiddenRows,
        undefined // No skip - render from row 0
      );

      this.ctx.restore();
    }

    // Render scrollable row headers
    this.ctx.save();
    this.ctx.beginPath();
    const rowHeaderY = this.headerHeight + frozenHeight;
    const rowHeaderHeight = this.viewport.height - this.headerHeight - frozenHeight;
    this.ctx.rect(0, rowHeaderY, this.headerWidth, rowHeaderHeight);
    this.ctx.clip();
    
    this.headerRenderer.renderRowHeaders(
      this.ctx,
      hasFrozenRows ? this.frozenRows : startRow,
      endRow,
      scrollTop,
      state.rowHeights,
      this.defaultRowHeight,
      this.headerWidth,
      this.headerHeight + frozenHeight, // Start after frozen area
      state.selection,
      state.activeCell,
      state.hiddenRows,
      hasFrozenRows ? this.frozenRows : undefined // Skip accumulating frozen row heights
    );

    this.ctx.restore();
  }
  
  /**
   * Render the corner cell
   */
  private renderCornerCell(): void {
    this.headerRenderer.renderCornerCell(
      this.ctx,
      this.headerWidth,
      this.headerHeight
    );
  }
  
  /**
   * Get the cell at a canvas point
   */
  getCellAtPoint(x: number, y: number): CellPosition | null {
    return this.hitTester.getCellAt(x, y);
  }
  
  /**
   * Get header at a canvas point
   */
  getHeaderAtPoint(x: number, y: number) {
    return this.hitTester.getHeaderAt(x, y);
  }
  
  /**
   * Get resize handle at a canvas point
   */
  getResizeHandleAtPoint(x: number, y: number) {
    return this.hitTester.getResizeHandleAt(x, y);
  }
  
  /**
   * Check if point is on fill handle
   */
  isFillHandleAtPoint(x: number, y: number): boolean {
    if (!this.renderState?.selection?.activeCell) {
      return false;
    }
    return this.hitTester.getFillHandleAt(
      x,
      y,
      this.renderState.selection.ranges[0]
    );
  }
  
  /**
   * Get cell bounds in canvas coordinates (accounting for freeze panes)
   */
  getCellBounds(row: number, col: number): Rect | null {
    if (!this.renderState) return null;
    
    const { scrollTop, scrollLeft } = this.viewport;
    const { frozenWidth, frozenHeight } = this.freezeDimensions;
    const state = this.renderState;
    const hiddenRows = state.hiddenRows ?? new Set<number>();
    const hiddenCols = state.hiddenCols ?? new Set<number>();
    
    // If the cell itself is hidden, return null
    if (hiddenRows.has(row) || hiddenCols.has(col)) {
      return null;
    }
    
    // Determine if cell is in frozen area
    const isRowFrozen = row < this.frozenRows;
    const isColFrozen = col < this.frozenCols;
    
    // Calculate x position (skip hidden columns)
    let x: number;
    if (isColFrozen) {
      // Cell is in frozen columns - no scroll offset
      x = this.headerWidth;
      for (let c = 0; c < col; c++) {
        if (!hiddenCols.has(c)) {
          x += state.colWidths.get(c) ?? this.defaultColWidth;
        }
      }
    } else {
      // Cell is in scrollable columns - apply scroll offset
      x = this.headerWidth + frozenWidth;
      for (let c = this.frozenCols; c < col; c++) {
        if (!hiddenCols.has(c)) {
          x += state.colWidths.get(c) ?? this.defaultColWidth;
        }
      }
      x -= scrollLeft;
    }
    
    // Calculate y position (skip hidden rows)
    let y: number;
    if (isRowFrozen) {
      // Cell is in frozen rows - no scroll offset
      y = this.headerHeight;
      for (let r = 0; r < row; r++) {
        if (!hiddenRows.has(r)) {
          y += state.rowHeights.get(r) ?? this.defaultRowHeight;
        }
      }
    } else {
      // Cell is in scrollable rows - apply scroll offset
      y = this.headerHeight + frozenHeight;
      for (let r = this.frozenRows; r < row; r++) {
        if (!hiddenRows.has(r)) {
          y += state.rowHeights.get(r) ?? this.defaultRowHeight;
        }
      }
      y -= scrollTop;
    }
    
    const width = state.colWidths.get(col) ?? this.defaultColWidth;
    const height = state.rowHeights.get(row) ?? this.defaultRowHeight;
    
    return { x, y, width, height };
  }
  
  /**
   * Get the current viewport
   */
  getViewport(): Viewport {
    return { ...this.viewport };
  }

  /** Width of the row-number column on the left (in CSS pixels). */
  getHeaderWidth(): number {
    return this.headerWidth;
  }

  /** Height of the column-letter row on top (in CSS pixels). */
  getHeaderHeight(): number {
    return this.headerHeight;
  }
  
  /**
   * Get the hit tester
   */
  getHitTester(): HitTester {
    return this.hitTester;
  }
  
  /**
   * Get the current freeze dimensions
   */
  getFreezeDimensions(): FreezeDimensions {
    return { ...this.freezeDimensions };
  }
  
  /**
   * Get frozen row/col counts
   */
  getFreezeConfig(): { frozenRows: number; frozenCols: number } {
    return { frozenRows: this.frozenRows, frozenCols: this.frozenCols };
  }
  
  /**
   * Cleanup resources
   */
  destroy(): void {
    if (this.animationFrameId !== null) {
      cancelAnimationFrame(this.animationFrameId);
      this.animationFrameId = null;
    }
  }
}

