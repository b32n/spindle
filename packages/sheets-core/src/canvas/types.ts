// Canvas rendering types for Spindle Sheets

import type { BorderKind, Cell, CellStyle, CellFormat, Selection, ColumnFilter, Range } from '../types';

/**
 * Rectangle bounds
 */
export interface Rect {
  x: number;
  y: number;
  width: number;
  height: number;
}

/**
 * Point coordinates
 */
export interface Point {
  x: number;
  y: number;
}

/**
 * Cell position in the grid
 */
export interface CellPosition {
  row: number;
  col: number;
}

/**
 * Viewport state - what's currently visible
 */
export interface Viewport {
  scrollTop: number;
  scrollLeft: number;
  width: number;
  height: number;
  startRow: number;
  endRow: number;
  startCol: number;
  endCol: number;
}

/**
 * Text style for rendering
 */
export interface TextStyle {
  fontFamily: string;
  fontSize: number;
  fontWeight: 'normal' | 'bold';
  fontStyle: 'normal' | 'italic';
  color: string;
  textAlign: 'left' | 'center' | 'right';
  verticalAlign: 'top' | 'middle' | 'bottom';
  textDecoration?: 'none' | 'underline' | 'line-through';
  /** An underline other than a single line (`double`, `singleAccounting`, `doubleAccounting`). */
  underlineKind?: 'double' | 'singleAccounting' | 'doubleAccounting';
  /** Strike-through drawn together with an underline. */
  strikethrough?: boolean;
  /** Indent levels from the edge the text is aligned to. */
  indent?: number;
  /** Scale the text down to fit the width of the cell (single line only). */
  shrinkToFit?: boolean;
  textRotation?: number; // degrees, -90..90; 0/undefined = horizontal
}

/**
 * Border style for a single edge
 */
export interface BorderStyle {
  width: number;
  color: string;
  style: 'solid' | 'dashed' | 'dotted' | 'double';
  /** The exact line style (one of the thirteen the file format has); how it is stroked comes from `BORDER_PAINT`. */
  kind?: BorderKind;
}

/**
 * Complete border configuration for a cell
 */
export interface CellBorders {
  top?: BorderStyle;
  right?: BorderStyle;
  bottom?: BorderStyle;
  left?: BorderStyle;
}

/**
 * Header hit test result
 */
export interface HeaderHit {
  type: 'row' | 'column';
  index: number;
  isResize: boolean;
}

/**
 * Resize handle hit test result
 */
export interface ResizeHandle {
  type: 'row' | 'column';
  index: number;
}

/**
 * Theme colors for the canvas
 */
export interface CanvasTheme {
  // Grid
  gridLineColor: string;
  gridLineWidth: number;
  
  // Headers
  headerBackgroundColor: string;
  headerTextColor: string;
  headerBorderColor: string;
  headerFont: string;
  headerFontSize: number;
  
  // Cells
  cellBackgroundColor: string;
  cellTextColor: string;
  cellFont: string;
  cellFontSize: number;
  
  // Selection
  selectionBorderColor: string;
  selectionBorderWidth: number;
  selectionFillColor: string;
  activeCellBorderColor: string;
  activeCellBorderWidth: number;
  
  // Fill handle
  fillHandleColor: string;
  fillHandleSize: number;
  
  // Formula reference highlighting colors (border, fill pairs)
  formulaReferenceColors: Array<{ border: string; fill: string }>;
  
  // Freeze panes
  freezeDividerColor: string;
  freezeDividerWidth: number;
  freezeShadowColor: string;
}

/**
 * Default theme
 */
export const DEFAULT_THEME: CanvasTheme = {
  // Grid
  gridLineColor: '#e2e2e2',
  gridLineWidth: 1,
  
  // Headers
  headerBackgroundColor: '#f8f9fa',
  headerTextColor: '#5f6368',
  headerBorderColor: '#e8eaed',
  headerFont: 'Arial',
  headerFontSize: 11,
  
  // Cells
  cellBackgroundColor: '#ffffff',
  cellTextColor: '#000000',
  cellFont: 'Arial',
  cellFontSize: 11,
  
  // Selection
  selectionBorderColor: '#1a73e8',
  selectionBorderWidth: 2,
  selectionFillColor: 'rgba(26, 115, 232, 0.1)',
  activeCellBorderColor: '#1a73e8',
  activeCellBorderWidth: 2,
  
  // Fill handle
  fillHandleColor: '#1a73e8',
  fillHandleSize: 6,
  
  // Formula reference highlighting colors (like Google Sheets)
  formulaReferenceColors: [
    { border: '#4285F4', fill: 'rgba(66, 133, 244, 0.15)' },  // Blue
    { border: '#EA4335', fill: 'rgba(234, 67, 53, 0.15)' },   // Red
    { border: '#FBBC04', fill: 'rgba(251, 188, 4, 0.15)' },   // Yellow
    { border: '#34A853', fill: 'rgba(52, 168, 83, 0.15)' },   // Green
    { border: '#9C27B0', fill: 'rgba(156, 39, 176, 0.15)' },  // Purple
    { border: '#FF9800', fill: 'rgba(255, 152, 0, 0.15)' },   // Orange
    { border: '#00BCD4', fill: 'rgba(0, 188, 212, 0.15)' },   // Cyan
    { border: '#E91E63', fill: 'rgba(233, 30, 99, 0.15)' },   // Pink
  ],
  
  // Freeze panes
  freezeDividerColor: '#c0c0c0',
  freezeDividerWidth: 2,
  freezeShadowColor: 'rgba(0, 0, 0, 0.08)',
};

/**
 * Configuration for the canvas renderer
 */
export interface CanvasRendererConfig {
  canvas: HTMLCanvasElement;
  devicePixelRatio?: number;
  defaultRowHeight: number;
  defaultColWidth: number;
  headerHeight: number;
  headerWidth: number;
  theme?: Partial<CanvasTheme>;
}

/**
 * Formula range for highlighting cell references in formulas
 */
export interface FormulaRangeHighlight {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
  colorIndex: number; // Index for color selection
}

/**
 * State required for rendering
 */
export interface RenderState {
  cells: Map<string, Cell>;
  styles: Map<string, CellStyle>;
  formats: Map<string, CellFormat>;
  selection: Selection | null;
  activeCell: CellPosition | null;
  editingCell: CellPosition | null;
  rowHeights: Map<number, number>;
  colWidths: Map<number, number>;
  rowCount: number;
  colCount: number;
  /** Formula ranges to highlight (cell references in formulas) */
  formulaRanges?: FormulaRangeHighlight[];
  /** Hidden rows (Set of row indices) */
  hiddenRows?: Set<number>;
  /** Hidden columns (Set of column indices) */
  hiddenCols?: Set<number>;
  /** Number of frozen rows (rows that stay visible when scrolling vertically) */
  frozenRows?: number;
  /** Number of frozen columns (columns that stay visible when scrolling horizontally) */
  frozenCols?: number;
  /** Active filters (column -> filter) */
  filters?: Map<number, ColumnFilter>;
  /** Filtered rows (rows that should be visible after filtering) */
  filteredRows?: Set<number>;
  /** Cells with an open comment thread, keyed "row:col" — drives the corner marker. */
  commentedCells?: Set<string>;
  /** Merged cell regions as numeric ranges (resolved from stable IDs). */
  mergedRegions?: Range[];
  /**
   * Default formats of columns and rows (index -> style, and number format): what an empty cell there shows (fill and borders). A row's
   * wins over its column's. Only the formatted ones are in the maps.
   */
  colFormats?: Map<number, { style?: CellStyle; format?: CellFormat }>;
  rowFormats?: Map<number, { style?: CellStyle; format?: CellFormat }>;
}

/**
 * Region that needs to be redrawn
 */
export interface DirtyRegion {
  type: 'all' | 'cells' | 'headers' | 'selection';
  bounds?: Rect;
}

/**
 * Cursor types for different interactions
 */
export type CursorType = 
  | 'default'
  | 'pointer'
  | 'cell'
  | 'col-resize'
  | 'row-resize'
  | 'crosshair'
  | 'grab'
  | 'grabbing';

/**
 * Mouse event data with grid coordinates
 */
export interface CanvasMouseEvent {
  // Screen coordinates relative to canvas
  x: number;
  y: number;
  // Grid coordinates (if over a cell)
  cell: CellPosition | null;
  // Header info (if over a header)
  header: HeaderHit | null;
  // Resize handle (if near one)
  resizeHandle: ResizeHandle | null;
  // Fill handle
  isFillHandle: boolean;
  // Original event
  originalEvent: MouseEvent;
}

/**
 * Scroll event data
 */
export interface CanvasScrollEvent {
  scrollTop: number;
  scrollLeft: number;
  deltaX: number;
  deltaY: number;
}

