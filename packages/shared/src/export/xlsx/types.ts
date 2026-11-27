// Generic xlsx interchange shapes. Deliberately independent of any editor's
// own data model (e.g. sheets-core's Cell/CellStyle/CellFormat) — shared
// sits below the editor packages in the dependency graph and can't import
// their types. Callers resolve their own pooled styles/formats down to
// concrete values before calling exportXlsx.

import type { XlsxStyles, XlsxWorkbookInfo } from './package-types';

export type SpreadsheetCellValue = string | number | boolean | null;

export interface SpreadsheetBorder {
  style: 'thin' | 'medium' | 'thick' | 'dashed' | 'dotted' | 'double';
  /** #rrggbb */
  color?: string;
}

export interface SpreadsheetCellStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  fontFamily?: string;
  fontSize?: number;
  /** #rrggbb */
  fontColor?: string;
  /** #rrggbb */
  backgroundColor?: string;
  horizontalAlign?: 'left' | 'center' | 'right';
  verticalAlign?: 'top' | 'middle' | 'bottom';
  wrapText?: boolean;
  /** Degrees, -90..90. */
  textRotation?: number;
  border?: {
    top?: SpreadsheetBorder;
    right?: SpreadsheetBorder;
    bottom?: SpreadsheetBorder;
    left?: SpreadsheetBorder;
  };
}

export interface SpreadsheetNumberFormat {
  /** Excel-syntax format code, e.g. "#,##0.00", "0.00%", "$#,##0.00", "m/d/yyyy". */
  code: string;
}

export interface SpreadsheetCell {
  value: SpreadsheetCellValue;
  /**
   * A1-style formula text, no leading '='. `value` still carries the last
   * computed result — exportXlsx writes both (xlsx caches a formula's result
   * alongside its expression so spreadsheet apps have something to display
   * before they recalculate).
   */
  formula?: string;
  style?: SpreadsheetCellStyle;
  numberFormat?: SpreadsheetNumberFormat;
  hyperlink?: string;
}

export interface SpreadsheetMergedRange {
  /** 0-based, inclusive. */
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}

export interface SpreadsheetSheet {
  name: string;
  /** Sparse. Key `${row}:${col}`, 0-based. */
  cells: Map<string, SpreadsheetCell>;
  columnWidths?: Map<number, number>;
  rowHeights?: Map<number, number>;
  hiddenRows?: Set<number>;
  hiddenCols?: Set<number>;
  frozenRows?: number;
  frozenCols?: number;
  mergedRanges?: SpreadsheetMergedRange[];
}

export interface SpreadsheetExportInput {
  sheets: SpreadsheetSheet[];
  activeSheetIndex?: number;
}

export interface SpreadsheetImportResult {
  sheets: SpreadsheetSheet[];
  activeSheetIndex?: number;
  /** Non-fatal issues (unrecognized cell shape, unresolvable shared formula, etc.) — surfaced rather than silently dropped. */
  warnings: string[];
  /** Workbook-level facts from the own reader (S2): sheet states, defined names, calculation properties, `date1904`. Absent from the `exceljs` path. */
  workbook?: XlsxWorkbookInfo;
  /** The style tables with theme, indexed and tint colours resolved (S2). Absent from the `exceljs` path. */
  styles?: XlsxStyles;
}

/** Options for the loss detector (`detectXlsxLoss`, and `inspectXlsx` in sheets-core). All optional. */
export interface DetectXlsxLossOptions {
  /** `'summary'` (default) drives the import-confirmation dialog; `'full'` adds the finer ids the fidelity harness records. */
  detail?: 'summary' | 'full';
  /** Upper-case names of the functions the formula engine implements; without it `xlsx.formulas.unsupported` is not reported. */
  knownFunctions?: ReadonlySet<string>;
  /** True when the editor reproduces a number format code exactly; without it `xlsx.numberFormats.approximated` is not reported. */
  isNumberFormatExact?: (code: string) => boolean;
}
