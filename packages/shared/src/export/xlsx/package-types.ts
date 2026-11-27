// Plain-data shapes of the parts of an XLSX package that the own reader (`@b32nio/spindle-io`, S2) understands: workbook,
// styles, shared strings, theme. Everything here is JSON-like (no classes, Maps or functions), so a result can be handed
// across a worker boundary (S20). They are reported through the optional fields of `SpreadsheetImportResult`; editors that
// do not know them keep working on the fields they already read.

/**
 * A colour as written in the file, plus what it resolves to. `rgb`, `theme`, `indexed` and `auto` are the file's own
 * attributes (one of them is normally set); `tint` applies to any of them. `resolved` is `#rrggbb` once the theme, the
 * indexed palette and the tint are applied, and absent for `auto` (the application decides) or a colour that cannot be read.
 */
export interface XlsxColor {
  rgb?: string;
  theme?: number;
  indexed?: number;
  auto?: boolean;
  tint?: number;
  resolved?: string;
}

export interface XlsxFont {
  bold?: boolean;
  italic?: boolean;
  strike?: boolean;
  /** `single` when `<u/>` has no value. */
  underline?: 'single' | 'double' | 'singleAccounting' | 'doubleAccounting';
  verticalAlign?: 'baseline' | 'superscript' | 'subscript';
  size?: number;
  name?: string;
  family?: number;
  charset?: number;
  /** `major`, `minor` or `none`: the theme font scheme the font belongs to. */
  scheme?: string;
  color?: XlsxColor;
}

export interface XlsxPatternFill {
  type: 'pattern';
  /** `none`, `solid`, `gray125`, `darkGray`, `lightGrid`, ... as written. */
  pattern: string;
  fg?: XlsxColor;
  bg?: XlsxColor;
}

export interface XlsxGradientFill {
  type: 'gradient';
  gradientType: 'linear' | 'path';
  /** Degrees, linear gradients. */
  degree?: number;
  left?: number;
  right?: number;
  top?: number;
  bottom?: number;
  stops: Array<{ position: number; color: XlsxColor }>;
}

export type XlsxFill = XlsxPatternFill | XlsxGradientFill;

export interface XlsxBorderSide {
  /** `thin`, `medium`, `dashed`, `hair`, `mediumDashDotDot`, ... as written. */
  style: string;
  color?: XlsxColor;
}

export interface XlsxBorder {
  left?: XlsxBorderSide;
  right?: XlsxBorderSide;
  top?: XlsxBorderSide;
  bottom?: XlsxBorderSide;
  diagonal?: XlsxBorderSide;
  diagonalUp?: boolean;
  diagonalDown?: boolean;
  outline?: boolean;
}

export interface XlsxAlignment {
  horizontal?: string;
  vertical?: string;
  /** Degrees as written: 0 to 180, or 255 for stacked text. */
  textRotation?: number;
  wrapText?: boolean;
  indent?: number;
  relativeIndent?: number;
  justifyLastLine?: boolean;
  shrinkToFit?: boolean;
  readingOrder?: number;
}

export interface XlsxProtection {
  locked?: boolean;
  hidden?: boolean;
}

/** One entry of `cellXfs` or `cellStyleXfs`. The `apply*` flags are kept as written (absent when the file has none). */
export interface XlsxCellFormat {
  numFmtId: number;
  fontId: number;
  fillId: number;
  borderId: number;
  /** The `cellStyleXfs` entry this format derives from (`cellXfs` only). */
  xfId?: number;
  applyNumberFormat?: boolean;
  applyFont?: boolean;
  applyFill?: boolean;
  applyBorder?: boolean;
  applyAlignment?: boolean;
  applyProtection?: boolean;
  quotePrefix?: boolean;
  pivotButton?: boolean;
  alignment?: XlsxAlignment;
  protection?: XlsxProtection;
}

/** A named cell style (`cellStyles`), e.g. `Normal`, `Hyperlink`, `Good`. */
export interface XlsxNamedStyle {
  name: string;
  /** Index into `cellStyleXfs`. */
  xfId: number;
  builtinId?: number;
  hidden?: boolean;
}

/** A differential format (`dxfs`): only what it overrides. */
export interface XlsxDifferentialFormat {
  font?: XlsxFont;
  numFmt?: { id: number; code: string };
  fill?: XlsxFill;
  alignment?: XlsxAlignment;
  border?: XlsxBorder;
  protection?: XlsxProtection;
}

export interface XlsxStyles {
  /** Custom number formats (`numFmts`); built-in ids are not listed (ids 0 to 49 that have no entry here are the built-ins). */
  numFmts: Array<{ id: number; code: string }>;
  fonts: XlsxFont[];
  fills: XlsxFill[];
  borders: XlsxBorder[];
  cellStyleXfs: XlsxCellFormat[];
  cellXfs: XlsxCellFormat[];
  cellStyles: XlsxNamedStyle[];
  dxfs: XlsxDifferentialFormat[];
  /** The workbook's own indexed palette (`colors/indexedColors`), `#rrggbb` per index; absent when it uses the default. */
  indexedColors?: string[];
}

/** One run of a rich-text shared string. */
export interface XlsxTextRun {
  text: string;
  font?: XlsxFont;
}

/** A shared-string item: the plain text, and the formatting runs when the item had any. Phonetic text is not part of `text`. */
export interface XlsxSharedString {
  text: string;
  runs?: XlsxTextRun[];
}

export interface XlsxTheme {
  /** `#rrggbb` in the order a cell colour's `theme` index uses: lt1, dk1, lt2, dk2, accent1 to accent6, hlink, folHlink. */
  colors: string[];
  majorFont?: string;
  minorFont?: string;
}

export type XlsxSheetState = 'visible' | 'hidden' | 'veryHidden';

/** A part of the package a relationship points at. */
export interface XlsxRelationship {
  id: string;
  /** The type URI as written. */
  type: string;
  /** Last segment of the type: `worksheet`, `comments`, `hyperlink`, `threadedComment`, ... as written. */
  kind: string;
  /** Resolved part name (no leading slash) for an internal target; the target text as written for an external one. */
  target: string;
  external: boolean;
}

/** The comment parts of a sheet: located, not read. */
export interface XlsxCommentParts {
  /** Legacy notes (`comments`). */
  comments?: string;
  /** The VML drawing that anchors them. */
  vmlDrawing?: string;
  /** Threaded comments (Excel 365). */
  threadedComments: string[];
}

export interface XlsxSheetRef {
  name: string;
  sheetId: number;
  state: XlsxSheetState;
  /** The relationship id as written (`r:id`). */
  relId: string;
  /** `worksheet`, `chartsheet`, `dialogsheet`, `macrosheet`, or the relationship kind of anything else. */
  kind: string;
  /** The part name, absent when the relationship is missing or points at a part the package does not hold. */
  part?: string;
  /** Every relationship of the sheet part (empty when it has none). */
  rels: XlsxRelationship[];
  commentParts: XlsxCommentParts;
}

export interface XlsxDefinedName {
  name: string;
  /** The formula text as written, no leading `=`. */
  ref: string;
  /** Index into `sheets` when the name is local to a sheet. */
  localSheetId?: number;
  hidden?: boolean;
  comment?: string;
  /** Built-in names such as `_xlnm.Print_Area`. */
  builtin: boolean;
}

export interface XlsxCalcProperties {
  calcId?: number;
  calcMode?: 'auto' | 'manual' | 'autoNoTable';
  fullCalcOnLoad?: boolean;
  refMode?: 'A1' | 'R1C1';
  iterate?: boolean;
  iterateCount?: number;
  iterateDelta?: number;
  fullPrecision?: boolean;
  calcOnSave?: boolean;
  concurrentCalc?: boolean;
  forceFullCalc?: boolean;
}

export interface XlsxWorkbookInfo {
  /** The workbook part name. */
  part: string;
  date1904: boolean;
  codeName?: string;
  /** `appName` of `fileVersion` (`xl`, `Calc`, ...), when written. */
  appName?: string;
  activeTab?: number;
  firstSheet?: number;
  calc: XlsxCalcProperties;
  sheets: XlsxSheetRef[];
  definedNames: XlsxDefinedName[];
  /** The shared-strings, styles, theme and persons (threaded-comment authors) parts, when present. */
  parts: { sharedStrings?: string; styles?: string; theme?: string; persons?: string };
}
