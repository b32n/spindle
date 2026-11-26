// Formula parser types

export interface CellReference {
  row: number;
  col: number;
  rowAbsolute: boolean;
  colAbsolute: boolean;
  sheetName?: string; // Optional sheet name for cross-sheet references
}

/** Sheet limits the whole-column / whole-row forms (`A:A`, `1:1`) expand to (Excel's). */
export const MAX_SHEET_ROWS = 1_048_576;
export const MAX_SHEET_COLS = 16_384;

export interface RangeReference {
  start: CellReference;
  end: CellReference;
  /**
   * Set for a whole-column (`A:C`, rows unbounded) or whole-row (`1:3`, columns
   * unbounded) range. The corners then carry the nominal sheet extent on the
   * unbounded axis (rows 0..MAX_SHEET_ROWS-1, columns 0..MAX_SHEET_COLS-1);
   * evaluation clamps that axis to the used range.
   */
  whole?: 'column' | 'row';
}

/**
 * Node kinds. `error` is an error literal (`#N/A`); `array` an array constant
 * (`value` is the row-major matrix); `percent` a postfix `%` on `left`;
 * `implicit` the `@` implicit-intersection operator on `left`; `spillref` the
 * spill operator (`A1#`, `cellRef` is the anchor); `ref3d` a 3D reference
 * (`Sheet1:Sheet3!A1`, `cellRef` or `rangeRef` on the first sheet, last sheet in
 * `sheetEnd`). The reference-combining operators are `operator` nodes: `:`
 * (range over arbitrary references), ` ` (intersection) and `,` (union).
 */
export type FormulaNodeType =
  | 'number'
  | 'string'
  | 'cell'
  | 'range'
  | 'function'
  | 'operator'
  | 'variable'
  | 'error'
  | 'array'
  | 'percent'
  | 'implicit'
  | 'spillref'
  | 'ref3d';

export interface ParsedFormulaNode {
  type: FormulaNodeType;
  value?: unknown;
  cellRef?: CellReference;
  rangeRef?: RangeReference;
  /** Last sheet of a 3D reference (the first sheet is the ref's `sheetName`). */
  sheetEnd?: string;
  functionName?: string;
  args?: ParsedFormulaNode[];
  operator?: string;
  left?: ParsedFormulaNode;
  right?: ParsedFormulaNode;
}

export interface ParseResult {
  ast: ParsedFormulaNode;
  dependencies: Set<string>; // cellKeys
  error?: string;
}

export interface EvaluationContext {
  getCellValue: (row: number, col: number, sheetId?: string, sheetName?: string) => unknown;
  getRangeValues: (range: RangeReference, sheetId?: string, sheetName?: string) => unknown[][];
  getSheetIdByName?: (sheetName: string) => string | undefined; // Helper to resolve sheet name to sheet ID
  isCellFormula?: (row: number, col: number, sheetName?: string) => boolean; // Whether a cell holds a formula (for ISFORMULA)
  /** Sheet names in workbook order (resolves the span of a 3D reference). */
  getSheetNames?: () => string[];
  /** The spill region anchored at a cell (for `A1#`), or undefined when that cell is not a spill anchor. */
  getSpillRegion?: (
    row: number,
    col: number,
    sheetName?: string
  ) => { rows: number; cols: number; values: unknown[][] } | undefined;
}

