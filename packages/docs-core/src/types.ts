// Core types for Spindle Docs

// ============================================================================
// Page Configuration
// ============================================================================

/** Page sizes in pixels (96 DPI) */
export const PAGE_SIZES = {
  LETTER: { w: 816, h: 1056 },      // 8.5" x 11"
  LEGAL: { w: 816, h: 1344 },       // 8.5" x 14"
  A4: { w: 794, h: 1123 },          // 210mm x 297mm
  A5: { w: 559, h: 794 },           // 148mm x 210mm
  TABLOID: { w: 1056, h: 1632 },    // 11" x 17"
  EXECUTIVE: { w: 696, h: 1008 },   // 7.25" x 10.5"
  B5: { w: 708, h: 1001 },          // 182mm x 257mm
} as const;

export type PageSizeKey = keyof typeof PAGE_SIZES;

export interface PageSize {
  w: number;
  h: number;
}

export interface PageMargins {
  top: number;
  right: number;
  bottom: number;
  left: number;
  header?: number;
  footer?: number;
}

export interface PageConfig {
  size: PageSize;
  margins: PageMargins;
  orientation: 'portrait' | 'landscape';
}

/** Default page configuration (Letter, portrait, 1-inch margins) */
export const DEFAULT_PAGE_CONFIG: PageConfig = {
  size: PAGE_SIZES.LETTER,
  margins: {
    top: 96,      // 1 inch
    right: 96,    // 1 inch
    bottom: 96,   // 1 inch
    left: 96,     // 1 inch
    header: 48,   // 0.5 inch
    footer: 48,   // 0.5 inch
  },
  orientation: 'portrait',
};

// ============================================================================
// Dynamic Fields (for headers/footers)
// ============================================================================

/** Types of dynamic fields that can be inserted in headers/footers */
export type DynamicFieldType = 'pageNumber' | 'totalPages' | 'date' | 'time' | 'title';

/** A dynamic field run that renders contextual content */
export interface DynamicFieldRun {
  type: 'dynamicField';
  fieldType: DynamicFieldType;
  /** Optional format string (e.g., "MM/DD/YYYY" for date) */
  format?: string;
}

// ============================================================================
// Text Styling
// ============================================================================

export interface TextStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  fontFamily?: string;
  fontSize?: number;
  color?: string;
  backgroundColor?: string;
  superscript?: boolean;
  subscript?: boolean;
  smallCaps?: boolean;
  /** All capitals (DOCX `caps`). */
  allCaps?: boolean;
  /** Hidden text (DOCX `vanish`): kept and exported; the editor still shows it, dimmed. */
  hidden?: boolean;
  /** East Asian and complex-script fonts, when they differ from `fontFamily` (the Latin font). */
  fontFamilyEastAsia?: string;
  fontFamilyComplex?: string;
}

export const DEFAULT_TEXT_STYLE: TextStyle = {
  bold: false,
  italic: false,
  underline: false,
  strikethrough: false,
  fontFamily: 'Arial',
  fontSize: 11,
  color: '#000000',
};

// ============================================================================
// Paragraph Styling
// ============================================================================

export type TextAlignment = 'left' | 'center' | 'right' | 'justify';
export type LineSpacingType = 'single' | 'onePointFive' | 'double' | 'atLeast' | 'exactly' | 'multiple';

export interface LineSpacing {
  type: LineSpacingType;
  value?: number; // Used for 'atLeast', 'exactly', 'multiple'
}

export interface ParagraphStyle {
  alignment?: TextAlignment;
  lineSpacing?: LineSpacing;
  spaceBefore?: number;
  spaceAfter?: number;
  firstLineIndent?: number;
  leftIndent?: number;
  rightIndent?: number;
  keepWithNext?: boolean;
  keepTogether?: boolean;
  pageBreakBefore?: boolean;
}

export const DEFAULT_PARAGRAPH_STYLE: ParagraphStyle = {
  alignment: 'left',
  lineSpacing: { type: 'single' },
  spaceBefore: 0,
  spaceAfter: 8,
  firstLineIndent: 0,
  leftIndent: 0,
  rightIndent: 0,
};

// ============================================================================
// Text Runs and Inline Content
// ============================================================================

export interface TextRun {
  type: 'text';
  text: string;
  styleId?: string; // Reference to shared style in StylePool
  // Inline styles (when not using styleId)
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  superscript?: boolean;
  subscript?: boolean;
  fontFamily?: string;
  fontSize?: number;
  color?: string;
  backgroundColor?: string;
  smallCaps?: boolean;
  allCaps?: boolean;
  hidden?: boolean;
  fontFamilyEastAsia?: string;
  fontFamilyComplex?: string;
  /** Thread ID when this run is covered by a comment. */
  commentThreadId?: string;
  /** Named character style of the run (a key of `DocumentData.namedStyles`); distinct from the pool `styleId`. */
  namedStyleId?: string;
}

export interface InlineImage {
  type: 'image';
  src: string;
  width: number;
  height: number;
  alt?: string;
}

export interface InlineLink extends TextStyle {
  type: 'link';
  text: string;
  href: string;
  title?: string;
  styleId?: string;
  /** Thread ID when this link is covered by a comment. */
  commentThreadId?: string;
  /** Named character style of the run (a key of `DocumentData.namedStyles`). */
  namedStyleId?: string;
}

/**
 * A field the editor does not evaluate (REF, PAGEREF, SEQ, STYLEREF, AUTHOR, MERGEFIELD, IF, ...). It displays its
 * cached result, which is only recomputed by the application that opens the exported file.
 */
export interface InlineField {
  type: 'field';
  /** The instruction as written, e.g. `REF _Ref123 \\h`. */
  instruction: string;
  /** The field type: the first word of the instruction, upper case (`REF`, `SEQ`, ...). */
  kind: string;
  /** The cached result: plain text, or runs when the result carries formatting or images. */
  result: string | Array<TextRun | InlineImage>;
  /** Flagged for the consuming application to recompute when the file is opened (`w:dirty`). */
  dirty?: boolean;
}

/**
 * A named position in the text (a Word bookmark): the target of internal links (an `InlineLink` whose `href` is
 * `#name`). It takes no space but counts as one position in the editor. A bookmark that covers text is kept as a
 * point at its start.
 */
export interface InlineBookmark {
  type: 'bookmark';
  name: string;
}

export type InlineContent = TextRun | InlineImage | InlineLink | DynamicFieldRun | InlineField | InlineBookmark;

// ============================================================================
// Block Types
// ============================================================================

export type BlockType = 'paragraph' | 'heading' | 'list-item' | 'table' | 'image' | 'horizontal-rule' | 'page-break';
export type HeadingLevel = 1 | 2 | 3 | 4 | 5 | 6;
export type ListType = 'bullet' | 'numbered';

export interface BaseBlock {
  id: string;
  type: BlockType;
}

/**
 * Paragraph formatting carried on the block itself (not through the paragraph style pool), so it lives in the
 * editor document and is shared in collaboration. Distances are pixels (96 dpi).
 */
export interface ParagraphFormat {
  alignment?: TextAlignment;
  lineSpacing?: LineSpacing;
  spaceBefore?: number;
  spaceAfter?: number;
  /** Left indent in pixels. */
  leftIndent?: number;
  /** Right indent in pixels. */
  rightIndent?: number;
  /** First-line indent in pixels; negative for a hanging indent. */
  firstLineIndent?: number;
  keepWithNext?: boolean;
  /** Keep the paragraph's lines on one page. */
  keepTogether?: boolean;
  pageBreakBefore?: boolean;
  /** Named paragraph style of the block (a key of `DocumentData.namedStyles`); distinct from the pool `styleId`. */
  namedStyleId?: string;
}

export interface ParagraphBlock extends BaseBlock, ParagraphFormat {
  type: 'paragraph';
  content: InlineContent[];
  styleId?: string; // Reference to paragraph style in pool
  // Inline paragraph styles (when not using styleId)
  /** Indent level set by the indent/outdent commands (24 px each), on top of `leftIndent`. */
  indent?: number;
}

export interface HeadingBlock extends BaseBlock, ParagraphFormat {
  type: 'heading';
  level: HeadingLevel;
  content: InlineContent[];
  styleId?: string;
}

export interface ListItemBlock extends BaseBlock, ParagraphFormat {
  type: 'list-item';
  listType: ListType;
  level: number; // Nesting level (0-based)
  /** First number of a numbered list (set on the list's first item when not 1; a `start` on a later item begins a new list). */
  start?: number;
  content: InlineContent[];
  styleId?: string;
}

/** One border line. `widthPx` defaults to 1; `color` is a hex colour, absent for automatic (black). */
export interface TableBorder {
  style: 'single' | 'double' | 'dashed' | 'dotted' | 'none';
  widthPx?: number;
  color?: string;
}

/**
 * Borders of a table or of one cell. On a table, a side that is not listed has no line (`{}` is a table without
 * borders); a table without `borders` shows the editor's default grid. On a cell, a side that is not listed
 * takes the table's line there.
 */
export interface TableBorders {
  top?: TableBorder;
  right?: TableBorder;
  bottom?: TableBorder;
  left?: TableBorder;
  /** Lines between rows (tables only). */
  insideH?: TableBorder;
  /** Lines between columns (tables only). */
  insideV?: TableBorder;
}

/** Which parts of a table style apply to a table (DOCX `w:tblLook`). */
export interface TableLook {
  firstRow?: boolean;
  lastRow?: boolean;
  firstColumn?: boolean;
  lastColumn?: boolean;
  noHBand?: boolean;
  noVBand?: boolean;
}

export type CellVerticalAlign = 'top' | 'center' | 'bottom';

export interface TableCell {
  id: string;
  content: InlineContent[];
  colspan?: number;
  rowspan?: number;
  styleId?: string;
  backgroundColor?: string;
  /** The cell's own borders (with a table style's conditional borders already applied). */
  borders?: TableBorders;
  verticalAlign?: CellVerticalAlign;
  /** True for a header cell (`<th>`). */
  isHeader?: boolean;
  /** Per-column pixel widths for the columns this cell spans. */
  colwidth?: number[];
}

export interface TableRow {
  id: string;
  cells: TableCell[];
  height?: number;
  /** Repeat the row at the top of each page the table continues on (DOCX `w:tblHeader`). Kept and exported; tables are not split across pages yet. */
  repeatHeader?: boolean;
  /** Keep the row on one page (DOCX `w:cantSplit`). */
  cantSplit?: boolean;
}

export interface TableBlock extends BaseBlock {
  type: 'table';
  rows: TableRow[];
  colWidths?: number[];
  styleId?: string;
  /** Borders of the table (see `TableBorders`). */
  borders?: TableBorders;
  /** Which parts of the table style apply (DOCX `w:tblLook`). */
  look?: TableLook;
  /** Id of the table's named table style, a key of `namedStyles`. The cells hold the formatting the style gives them. */
  namedStyleId?: string;
}

export interface ImageBlock extends BaseBlock {
  type: 'image';
  src: string;
  width: number;
  height: number;
  alt?: string;
  title?: string;
  alignment?: 'left' | 'center' | 'right';
  caption?: InlineContent[];
}

export interface HorizontalRuleBlock extends BaseBlock {
  type: 'horizontal-rule';
}

export interface PageBreakBlock extends BaseBlock {
  type: 'page-break';
}

export type Block =
  | ParagraphBlock
  | HeadingBlock
  | ListItemBlock
  | TableBlock
  | ImageBlock
  | HorizontalRuleBlock
  | PageBreakBlock;

// ============================================================================
// Headers and Footers
// ============================================================================

/** Text run for header/footer content */
export interface HeaderFooterTextRun {
  type: 'text';
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fontSize?: number;
  fontFamily?: string;
  color?: string;
}

/** Image run for header/footer (e.g., company logo) */
export interface HeaderFooterImageRun {
  type: 'image';
  src: string;
  alt?: string;
  width?: number;
  height?: number;
}

/** Inline content for header/footer paragraphs */
export type HeaderFooterInlineContent = HeaderFooterTextRun | DynamicFieldRun | HeaderFooterImageRun;

/** A paragraph in header/footer */
export interface HeaderFooterParagraph {
  type: 'paragraph';
  content: HeaderFooterInlineContent[];
  alignment?: 'left' | 'center' | 'right';
}

/** Header/footer content configuration */
export interface HeaderFooterContent {
  /** Content blocks to render */
  blocks: HeaderFooterParagraph[];
  /** Whether first page uses different content */
  differentFirstPage?: boolean;
  /** Content for first page (if differentFirstPage is true) */
  firstPageBlocks?: HeaderFooterParagraph[];
}

/** @deprecated Use HeaderFooterContent instead */
export interface HeaderFooter {
  id: string;
  /** Content blocks for the default header/footer */
  content: Block[];
  /** Whether first page uses different content */
  differentFirstPage?: boolean;
  /** Content blocks for first page (used when differentFirstPage is true) */
  firstPageContent?: Block[];
  /** Whether odd/even pages use different content */
  differentOddEven?: boolean;
  /** Content blocks for odd pages (used when differentOddEven is true) */
  oddPageContent?: Block[];
  /** Content blocks for even pages (used when differentOddEven is true) */
  evenPageContent?: Block[];
}

// ============================================================================
// Section
// ============================================================================

export interface Section {
  id: string;
  pageConfig: PageConfig;
  blocks: Block[];
  header?: HeaderFooterContent;
  footer?: HeaderFooterContent;
}

// ============================================================================
// Document
// ============================================================================

export interface Document {
  id: string;
  title: string;
  sections: Section[];
  defaultPageConfig: PageConfig;
  createdAt?: string;
  updatedAt?: string;
  namedStyles?: Record<string, NamedStyle>;
}

// ============================================================================
// Selection and Cursor
// ============================================================================

export interface CursorPosition {
  blockId: string;
  offset: number;
  /** For inline content, which run index */
  runIndex?: number;
}

export interface TextSelection {
  anchor: CursorPosition;
  focus: CursorPosition;
  isCollapsed: boolean;
}

// ============================================================================
// Style Pools
// ============================================================================

export interface TextStylePool {
  styles: Map<string, TextStyle>;
  getOrCreate(style: TextStyle): string;
  get(styleId: string): TextStyle | undefined;
  getAllStyles(): Map<string, TextStyle>;
}

export interface ParagraphStylePool {
  styles: Map<string, ParagraphStyle>;
  getOrCreate(style: ParagraphStyle): string;
  get(styleId: string): ParagraphStyle | undefined;
  getAllStyles(): Map<string, ParagraphStyle>;
}

// ============================================================================
// Events
// ============================================================================

export type DocumentEventType =
  | 'documentChange'
  | 'blockChange'
  | 'blockAdd'
  | 'blockDelete'
  | 'selectionChange'
  | 'pageConfigChange'
  | 'historyChange'
  | 'commentChange'
  | 'commentEvent';

export interface DocumentEventData {
  type: DocumentEventType;
  payload: unknown;
}

export type DocumentEventHandler = (data: DocumentEventData) => void;

// ============================================================================
// History/Undo
// ============================================================================

export interface HistoryEntry {
  timestamp: number;
  snapshot: DocumentSnapshot;
  description?: string;
}

export interface DocumentSnapshot {
  sections: Section[];
  selection?: TextSelection;
}

// ============================================================================
// Serialization
// ============================================================================

/**
 * A named style of the document (DOCX `w:style`): a name and the style's own properties. Blocks and runs refer to
 * it by `namedStyleId`. Their own fields hold the effective formatting, so a style's properties are what is written
 * to `styles.xml` (the editor has no style gallery yet).
 */
export interface NamedStyle {
  /** Display name ("Callout"). */
  name: string;
  kind: 'paragraph' | 'character' | 'table';
  basedOn?: string;
  next?: string;
  link?: string;
  /** The default style of its kind. */
  isDefault?: boolean;
  /** Outline level 1 to 9 the style gives its paragraphs; 0 is body text (cancels an inherited level). */
  outlineLevel?: number;
  text?: TextStyle;
  paragraph?: ParagraphFormat;
  /** The table-specific part of a table style (`kind: 'table'`). */
  table?: TableStyleData;
}

/** Formatting a table style gives one part of a table. */
export interface TableStylePart {
  text?: TextStyle;
  backgroundColor?: string;
  borders?: TableBorders;
  verticalAlign?: CellVerticalAlign;
}

/**
 * The table part of a table style: the table's borders, the formatting of every cell (`whole`) and of the parts
 * a table's `look` selects, keyed by DOCX `w:tblStylePr` type (`firstRow`, `lastRow`, `firstCol`, `lastCol`,
 * `band1Horz`, `band2Horz`, `band1Vert`, `band2Vert`, `nwCell`, `neCell`, `swCell`, `seCell`).
 */
export interface TableStyleData {
  borders?: TableBorders;
  whole?: TableStylePart;
  conditions?: Record<string, TableStylePart>;
  rowBandSize?: number;
  colBandSize?: number;
}

export interface DocumentData {
  id: string;
  title: string;
  sections: SectionData[];
  defaultPageConfig: PageConfig;
  textStylePool: Record<string, TextStyle>;
  paragraphStylePool: Record<string, ParagraphStyle>;
  createdAt?: string;
  updatedAt?: string;
  /** Comment threads anchored to spans of the document body. */
  threads?: import('./comments').DocsCommentThread[];
  /** Named styles by style id (see `NamedStyle`). Synced in collaboration as the `namedStyles` Y map. */
  namedStyles?: Record<string, NamedStyle>;
}

export interface SectionData {
  id: string;
  pageConfig: PageConfig;
  blocks: Block[];
  header?: HeaderFooterContent;
  footer?: HeaderFooterContent;
}

