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
  /** Thread ID when this run is covered by a comment. */
  commentThreadId?: string;
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
}

export type InlineContent = TextRun | InlineImage | InlineLink | DynamicFieldRun;

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

export interface ParagraphBlock extends BaseBlock {
  type: 'paragraph';
  content: InlineContent[];
  styleId?: string; // Reference to paragraph style in pool
  // Inline paragraph styles (when not using styleId)
  alignment?: TextAlignment;
  indent?: number;
  lineSpacing?: LineSpacing;
  spaceBefore?: number;
  spaceAfter?: number;
}

export interface HeadingBlock extends BaseBlock {
  type: 'heading';
  level: HeadingLevel;
  content: InlineContent[];
  styleId?: string;
  // Inline paragraph styles (when not using styleId)
  alignment?: TextAlignment;
}

export interface ListItemBlock extends BaseBlock {
  type: 'list-item';
  listType: ListType;
  level: number; // Nesting level (0-based)
  /** First number of a numbered list (set on the list's first item when not 1). */
  start?: number;
  content: InlineContent[];
  styleId?: string;
}

export interface TableCell {
  id: string;
  content: InlineContent[];
  colspan?: number;
  rowspan?: number;
  styleId?: string;
  backgroundColor?: string;
  /** True for a header cell (`<th>`). */
  isHeader?: boolean;
  /** Per-column pixel widths for the columns this cell spans. */
  colwidth?: number[];
}

export interface TableRow {
  id: string;
  cells: TableCell[];
  height?: number;
}

export interface TableBlock extends BaseBlock {
  type: 'table';
  rows: TableRow[];
  colWidths?: number[];
  styleId?: string;
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
}

export interface SectionData {
  id: string;
  pageConfig: PageConfig;
  blocks: Block[];
  header?: HeaderFooterContent;
  footer?: HeaderFooterContent;
}

