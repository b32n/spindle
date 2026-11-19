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
  /** All capitals (`w:caps`). */
  allCaps?: boolean;
  /** Hidden text (`w:vanish`). */
  hidden?: boolean;
  /** East Asian and complex-script fonts (`w:rFonts` `eastAsia` / `cs`); `fontFamily` is the Latin font. */
  fontFamilyEastAsia?: string;
  fontFamilyComplex?: string;
  /** Id of the comment thread (see `DocCommentThread`) whose anchored range covers this run. */
  commentThreadId?: string;
  /** Id of the named character style (`w:rStyle`) of the run, a key of `namedStyles`. The run's own fields still hold the effective formatting. */
  namedStyleId?: string;
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

/**
 * A field the editor does not evaluate (REF, PAGEREF, SEQ, STYLEREF, AUTHOR, MERGEFIELD, IF, ...): its instruction and
 * the result Word last computed. The result is what is displayed; it is written back as the field's cached result so
 * Word can refresh it.
 */
export interface DocFieldRun {
  type: 'field';
  /** The instruction as written, e.g. `REF _Ref123 \\h`. */
  instruction: string;
  /** The field type: the first word of the instruction, upper case (`REF`, `SEQ`, ...). */
  kind: string;
  /** The cached result: plain text, or runs when the result carries formatting or images. */
  result: string | Array<DocTextRun | DocImageRun>;
  /** Set when the field is flagged for Word to recompute when the file is opened (`w:dirty`). */
  dirty?: boolean;
}

/**
 * A named position in the text (a Word bookmark), the target of internal links (`href: '#name'`) and of REF and
 * PAGEREF fields. It takes no space. A bookmark that covers text is kept as a point at its start.
 */
export interface DocBookmarkRun {
  type: 'bookmark';
  name: string;
}

export type DocInline = DocTextRun | DocLinkRun | DocImageRun | DocPageNumberRun | DocTotalPagesRun | DocFieldRun | DocBookmarkRun;

export interface DocParagraphStyle {
  alignment?: 'left' | 'center' | 'right' | 'justify';
  /** Points. */
  spaceBeforePt?: number;
  /** Points. */
  spaceAfterPt?: number;
  /** Points. */
  leftIndentPt?: number;
  /** Points; negative for a hanging indent. */
  firstLineIndentPt?: number;
  /** Points. */
  rightIndentPt?: number;
  lineSpacing?: 'single' | 'onePointFive' | 'double';
  /** Line spacing other than the three keywords: a `multiple` of single spacing, or `atLeast` / `exactly` a height in points. Takes precedence over `lineSpacing`. */
  lineSpacingCustom?: { rule: 'multiple' | 'atLeast' | 'exactly'; value: number };
  keepWithNext?: boolean;
  /** Keep the lines of the paragraph together on one page (`w:keepLines`). */
  keepTogether?: boolean;
  pageBreakBefore?: boolean;
  /** Id of the named paragraph style (`w:pStyle`) of the block, a key of `namedStyles`. The block's own fields still hold the effective formatting. */
  namedStyleId?: string;
}

/**
 * A named style of the document (a `w:style` of `styles.xml`). `text` and `paragraph` are the style's own
 * properties, not the result of its `basedOn` chain.
 */
export interface DocNamedStyle {
  /** Display name ("Callout", "heading 1"). */
  name: string;
  kind: 'paragraph' | 'character' | 'table';
  /** Style id this style is based on. */
  basedOn?: string;
  /** Style id of the paragraph that follows one in this style when the user presses Enter. */
  next?: string;
  /** Style id of the linked paragraph/character counterpart. */
  link?: string;
  /** The default style of its kind (`w:default`). */
  isDefault?: boolean;
  /** Outline level 1 to 9 (`w:outlineLvl`) the style gives its paragraphs; 0 is body text (`w:outlineLvl` 9), which cancels an outline level inherited from the style it is based on. */
  outlineLevel?: number;
  text?: DocRunBase;
  paragraph?: DocParagraphStyle;
}

export interface DocTableCell {
  content: DocInline[];
  colspan?: number;
  rowspan?: number;
}

export interface DocTableRow {
  cells: DocTableCell[];
}

/**
 * `listItem.start` is the number of the first item of a numbered list. Set it on the first item of each list
 * when it is not 1; a `start` on a later item begins a new list (a restart), even when it is 1.
 */
export type DocBlock =
  | ({ type: 'paragraph'; content: DocInline[] } & DocParagraphStyle)
  | ({ type: 'heading'; level: 1 | 2 | 3 | 4 | 5 | 6; content: DocInline[] } & DocParagraphStyle)
  | ({ type: 'listItem'; listType: 'bullet' | 'numbered'; level: number; content: DocInline[]; start?: number } & DocParagraphStyle)
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
  /** Named styles by style id, referenced by `namedStyleId` of blocks and runs. */
  namedStyles?: Record<string, DocNamedStyle>;
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
  /** Named styles the document uses (and the styles they are based on, linked to or followed by), by style id. Absent when none. */
  namedStyles?: Record<string, DocNamedStyle>;
  /** Non-fatal issues (unsupported feature, unparseable structure, etc.) — surfaced rather than silently dropped. */
  warnings: string[];
}
