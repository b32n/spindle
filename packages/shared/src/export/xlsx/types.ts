// Generic xlsx interchange shapes. Deliberately independent of any editor's
// own data model (e.g. sheets-core's Cell/CellStyle/CellFormat) — shared
// sits below the editor packages in the dependency graph and can't import
// their types. Callers resolve their own pooled styles/formats down to
// concrete values before calling exportXlsx.

import type { ImportLossItem } from '../loss';
import type { XlsxFormulaInfo, XlsxSharedString, XlsxSheetDetails, XlsxStyles, XlsxTheme, XlsxWorkbookInfo } from './package-types';

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
  /** Index into `SpreadsheetImportResult.styles.cellXfs` of the cell's format (own reader, S2). */
  styleIndex?: number;
  /** Index into `SpreadsheetImportResult.sharedStrings` when the value is a shared string (own reader, S2). */
  sharedString?: number;
  /** Set on a cell whose formula was shared or an array formula (own reader, S2). */
  formulaInfo?: XlsxFormulaInfo;
  /**
   * The formula is written exactly as given: the engine could not evaluate it, so it keeps the text the file had (future-function
   * prefixes included) next to the result the file saved for it (S4). Without it the writer adds the prefixes the file format needs.
   */
  cachedOnly?: true;
  /** A1 range the formula is an array formula over (`t="array"`, written from the anchor cell; a single cell for a one-cell array formula). */
  arrayRef?: string;
  /** With `arrayRef`: a dynamic-array formula (spills), written with the cell metadata (`cm`) Excel marks it with; without it a legacy (CSE) array formula. */
  dynamicArray?: true;
  /** A cell inside the range of an array formula, not its anchor: an error literal as its value is written as an error value (`t="e"`), as the file kept it, not as text. */
  inArray?: true;
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
  /** What the own reader read besides the cells: properties, views, defaults, column and row formats, links, raw extensions (S2). */
  details?: XlsxSheetDetails;
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
  /** Workbook-level facts from the own reader (S2): sheet states, defined names, calculation properties, `date1904`. Always set by `importXlsx` (the own reader). */
  workbook?: XlsxWorkbookInfo;
  /** The style tables with theme, indexed and tint colours resolved (S2). Always set by `importXlsx` (the own reader). */
  styles?: XlsxStyles;
  /** The shared-strings table in file order, rich runs kept (S2). Always set by `importXlsx` (the own reader). */
  sharedStrings?: XlsxSharedString[];
  /** The theme: its 12 colours in a cell colour's `theme` order, and the major/minor fonts (S2). Always set by `importXlsx` (the own reader). */
  theme?: XlsxTheme;
  /**
   * What the import limits (`ImportLimits`) cut from this file, as loss items (id `xlsx.package.sizeLimit`) a host can show with the
   * rest of the import report: cells past `maxTotalCells`, sheets past `maxSheets`, ranges, rows, columns, extension lists, formulas,
   * parts past `maxPartElements`... Each says what was cut and the option that raises the limit. Empty (and always set by the own
   * reader) when nothing was cut. `importFromXlsx` appends them to the loss report it returns.
   */
  limitLoss?: ImportLossItem[];
}

/** Options for the loss detector (`detectXlsxLoss`, and `inspectXlsx` in sheets-core). All optional. */
export interface DetectXlsxLossOptions {
  /** `'summary'` (default) drives the import-confirmation dialog; `'full'` adds the finer ids the fidelity harness records. */
  detail?: 'summary' | 'full';
  /** Upper-case names of the functions the formula engine implements; without it `xlsx.formulas.unsupported` is not reported. */
  knownFunctions?: ReadonlySet<string>;
  /** True when the editor reproduces a number format code exactly; without it `xlsx.numberFormats.approximated` is not reported. */
  isNumberFormatExact?: (code: string) => boolean;
  /**
   * True when the formula engine can evaluate a formula (the text as the file has it, no leading `=`); the rest are kept as their
   * saved value and reported as `xlsx.formulas.cachedOnly`. Without it only the functions `knownFunctions` lacks make a formula cached-only.
   */
  isFormulaEvaluable?: (formula: string) => boolean;
}
