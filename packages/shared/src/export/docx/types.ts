// Generic docx interchange shapes. Deliberately independent of docs-core's
// Block/InlineContent model — shared sits below the editor packages in the
// dependency graph and can't import their types. Callers resolve their own
// pooled styles down to concrete values (and text-style-pool refs to plain
// booleans/colors) before calling exportDocx.

export interface DocRunBase {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  superscript?: boolean;
  subscript?: boolean;
  smallCaps?: boolean;
  fontFamily?: string;
  /** Points. */
  fontSize?: number;
  /** #rrggbb */
  color?: string;
  /** #rrggbb (rendered as a highlight — docx has no free-form run background). */
  backgroundColor?: string;
  /** Id of the comment thread (see `DocCommentThread`) whose anchored range covers this run. */
  commentThreadId?: string;
}

export interface DocTextRun extends DocRunBase {
  type: 'text';
  text: string;
}

export interface DocLinkRun extends DocRunBase {
  type: 'link';
  text: string;
  href: string;
}

export interface DocImageRun {
  type: 'image';
  /** Only `data:` URIs are embedded; anything else is dropped with a warning. */
  src: string;
  /** Pixels (96dpi), same convention as the rest of this interchange format. */
  width: number;
  height: number;
  alt?: string;
}

/** A field that resolves to the current page number when the document is paginated. */
export interface DocPageNumberRun {
  type: 'pageNumber';
}

/** A field that resolves to the document's total page count. */
export interface DocTotalPagesRun {
  type: 'totalPages';
}

export type DocInline = DocTextRun | DocLinkRun | DocImageRun | DocPageNumberRun | DocTotalPagesRun;

export interface DocParagraphStyle {
  alignment?: 'left' | 'center' | 'right' | 'justify';
  /** Points. */
  spaceBeforePt?: number;
  /** Points. */
  spaceAfterPt?: number;
  /** Points. */
  leftIndentPt?: number;
  /** Points. */
  firstLineIndentPt?: number;
  lineSpacing?: 'single' | 'onePointFive' | 'double';
}

export interface DocTableCell {
  content: DocInline[];
  colspan?: number;
  rowspan?: number;
}

export interface DocTableRow {
  cells: DocTableCell[];
}

export type DocBlock =
  | ({ type: 'paragraph'; content: DocInline[] } & DocParagraphStyle)
  | ({ type: 'heading'; level: 1 | 2 | 3 | 4 | 5 | 6; content: DocInline[] } & DocParagraphStyle)
  | ({ type: 'listItem'; listType: 'bullet' | 'numbered'; level: number; content: DocInline[] } & DocParagraphStyle)
  | { type: 'table'; rows: DocTableRow[]; colWidthsPx?: number[] }
  | { type: 'image'; src: string; width: number; height: number; alt?: string; alignment?: 'left' | 'center' | 'right' }
  | { type: 'horizontalRule' }
  | { type: 'pageBreak' };

export interface DocHeaderFooterParagraph {
  content: DocInline[];
  alignment?: 'left' | 'center' | 'right';
}

export interface DocHeaderFooter {
  paragraphs: DocHeaderFooterParagraph[];
  differentFirstPage?: boolean;
  firstPageParagraphs?: DocHeaderFooterParagraph[];
}

export interface DocSection {
  /** Pixels (96dpi). */
  pageWidthPx: number;
  pageHeightPx: number;
  marginTopPx: number;
  marginRightPx: number;
  marginBottomPx: number;
  marginLeftPx: number;
  orientation: 'portrait' | 'landscape';
  blocks: DocBlock[];
  header?: DocHeaderFooter;
  footer?: DocHeaderFooter;
}

/** One comment of a thread: the root comment or a reply. */
export interface DocComment {
  /** Unique within the document (the source file's `w:id` on import). */
  id: string;
  /** Display name. */
  author: string;
  /** Stable identity of the author when the source records one (otherwise the display name). */
  authorId?: string;
  initials?: string;
  /** ISO-8601 timestamp as written in the source (`w:date`). */
  date?: string;
  /** Plain text; paragraph breaks are `\n`. */
  text: string;
}

/** A comment thread: `comments[0]` is the root, the rest are replies in order. */
export interface DocCommentThread {
  id: string;
  /** `done` in Word's terms. */
  resolved?: boolean;
  /** Plain text of the anchored range (empty when the comment has no surviving anchor). */
  quote: string;
  comments: DocComment[];
}

export interface DocumentExportInput {
  title: string;
  sections: DocSection[];
  /** Comment threads, anchored to the runs that carry the matching `commentThreadId` (a thread with no such run is written at the start of the document). */
  comments?: DocCommentThread[];
}

export interface DocumentExportResult {
  bytes: Uint8Array;
  /** Non-fatal issues (e.g. an image whose src isn't an embeddable data: URI). */
  warnings: string[];
}

export interface DocumentImportResult {
  title: string;
  sections: DocSection[];
  /** Comment threads; the runs they are anchored to carry the matching `commentThreadId`. Absent when the file has none. */
  comments?: DocCommentThread[];
  /** Non-fatal issues (unsupported feature, unparseable structure, etc.) — surfaced rather than silently dropped. */
  warnings: string[];
}
