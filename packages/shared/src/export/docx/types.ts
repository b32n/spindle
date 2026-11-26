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

/** The kind of a content control (a `w:sdt`): the element of its `w:sdtPr` that gives it its behaviour. */
export type DocContentControlType = 'richText' | 'text' | 'dropDown' | 'comboBox' | 'date' | 'checkbox' | 'picture' | 'docPart';

/** One choice of a drop-down list or combo box. */
export interface DocContentControlItem {
  text: string;
  value: string;
}

/**
 * The properties of a content control (the `w:sdtPr` of a `w:sdt`). The content it wraps is the control's `content`
 * (inline) or `blocks`; a checkbox shows its state as a glyph in that content as well as in `checked`.
 */
export interface DocContentControl {
  type: DocContentControlType;
  /** `w:tag`: the programmatic name. */
  tag?: string;
  /** `w:alias`: the title shown to the user. */
  alias?: string;
  /** `w:lock`: the control cannot be deleted (`sdtLocked`), its content cannot be edited (`contentLocked`), or both. */
  lock?: 'sdtLocked' | 'contentLocked' | 'sdtContentLocked';
  /** Name of the placeholder building block (`w:placeholder`); the text of the placeholder is the control's content while `showingPlaceholder`. */
  placeholder?: string;
  /** The content is the placeholder text, not a value (`w:showingPlcHdr`). */
  showingPlaceholder?: boolean;
  /** Choices of a `dropDown` or `comboBox`. */
  items?: DocContentControlItem[];
  /** Display format of a `date` (`w:dateFormat`, e.g. `dd/MM/yyyy`). */
  dateFormat?: string;
  /** The date a `date` holds (`w:fullDate`), as written. */
  dateValue?: string;
  /** State of a `checkbox` (`w14:checked`). */
  checked?: boolean;
  /** A `text` control that allows line breaks (`w:text w:multiLine`). */
  multiline?: boolean;
  /** Gallery of a `docPart` control (`w:docPartGallery`, e.g. `Cover Pages`, `Table of Contents`). */
  gallery?: string;
}

/** Content wrapped by a content control inside a paragraph (a run-level `w:sdt`). Controls nest. */
export interface DocContentControlRun {
  type: 'contentControl';
  control: DocContentControl;
  content: DocInline[];
}

/**
 * A reference to a footnote or endnote, with the note's text. The reference takes no text of its own: where it is shown
 * and how it is numbered is up to the application (Word numbers notes in document order and writes the number itself).
 * Only the first paragraph of a note is a paragraph of its own; further ones are joined to it with a line break
 * (`\n`), and a table, list or image in a note is not kept.
 */
export interface DocNoteRefRun {
  type: 'noteRef';
  kind: 'footnote' | 'endnote';
  /** The note's id in the source file (`w:id`); unique per kind. The export renumbers notes in document order. */
  id: string;
  /** The text of the note, without the reference mark Word puts at its start. */
  content: DocInline[];
  /** The mark that replaces the automatic number (`w:customMarkFollows`). */
  customMark?: string;
}

export type DocInline =
  | DocTextRun
  | DocLinkRun
  | DocImageRun
  | DocPageNumberRun
  | DocTotalPagesRun
  | DocFieldRun
  | DocBookmarkRun
  | DocContentControlRun
  | DocNoteRefRun;

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
  /** The table-specific part of a table style (`kind: 'table'`). */
  table?: DocTableStyle;
}

/** One border line. `widthPx` defaults to 1; `color` is a hex colour, absent for automatic (black). */
export interface DocBorder {
  style: 'single' | 'double' | 'dashed' | 'dotted' | 'none';
  widthPx?: number;
  color?: string;
}

/** Borders of a table (with the lines between its rows and columns) or of one cell (the four sides). */
export interface DocBorders {
  top?: DocBorder;
  right?: DocBorder;
  bottom?: DocBorder;
  left?: DocBorder;
  /** Lines between rows (tables only). */
  insideH?: DocBorder;
  /** Lines between columns (tables only). */
  insideV?: DocBorder;
}

/** Which parts of a table style apply (`w:tblLook`). */
export interface DocTableLook {
  firstRow?: boolean;
  lastRow?: boolean;
  firstColumn?: boolean;
  lastColumn?: boolean;
  noHBand?: boolean;
  noVBand?: boolean;
}

export type DocCellVerticalAlign = 'top' | 'center' | 'bottom';

/** Direction of the text in a cell when it is not horizontal (`w:textDirection`): bottom to top, or top to bottom. */
export type DocCellTextDirection = 'btLr' | 'tbRl';

/** Formatting a table style applies to one part of a table (a `w:tblStylePr`, or the style's whole-table cell properties). */
export interface DocTableStylePart {
  text?: DocRunBase;
  /** Cell shading, a hex colour. */
  backgroundColor?: string;
  /** Cell borders. */
  borders?: DocBorders;
  verticalAlign?: DocCellVerticalAlign;
}

/**
 * The table part of a table style: the borders of the whole table, the formatting of every cell (`whole`) and
 * of the parts selected by `tblLook`, keyed by `w:tblStylePr` type (`firstRow`, `lastRow`, `firstCol`, `lastCol`,
 * `band1Horz`, `band2Horz`, `band1Vert`, `band2Vert`, `nwCell`, `neCell`, `swCell`, `seCell`).
 */
export interface DocTableStyle {
  borders?: DocBorders;
  whole?: DocTableStylePart;
  conditions?: Record<string, DocTableStylePart>;
  rowBandSize?: number;
  colBandSize?: number;
}

export interface DocTableCell {
  /** The cell's first paragraph with content (or its first paragraph), as inline content. `blocks` holds the whole cell when it is more than one plain paragraph. */
  content: DocInline[];
  /** The whole content of a cell that has several blocks, a nested table, a list or a heading (paragraphs, lists, images and tables). Absent for a cell of one plain paragraph. */
  blocks?: DocBlock[];
  colspan?: number;
  rowspan?: number;
  /** Shading, a hex colour (`w:shd` fill), with a table style's conditional shading already applied. */
  backgroundColor?: string;
  /** The cell's own borders, with a table style's conditional borders already applied. Sides not given come from the table's. */
  borders?: DocBorders;
  verticalAlign?: DocCellVerticalAlign;
  /** Text direction when the text is not horizontal. Kept and written back; the editor shows the text horizontally. */
  textDirection?: DocCellTextDirection;
  /** Widths in pixels of the grid columns the cell spans. */
  colwidthPx?: number[];
}

export interface DocTableRow {
  cells: DocTableCell[];
  /** Repeat the row at the top of each page the table continues on (`w:tblHeader`). */
  repeatHeader?: boolean;
  /** Keep the row on one page (`w:cantSplit`). */
  cantSplit?: boolean;
}

/**
 * `listItem.start` is the number of the first item of a numbered list. Set it on the first item of each list
 * when it is not 1; a `start` on a later item begins a new list (a restart), even when it is 1.
 */
export type DocBlock =
  | ({ type: 'paragraph'; content: DocInline[] } & DocParagraphStyle)
  | ({ type: 'heading'; level: 1 | 2 | 3 | 4 | 5 | 6; content: DocInline[] } & DocParagraphStyle)
  | ({ type: 'listItem'; listType: 'bullet' | 'numbered'; level: number; content: DocInline[]; start?: number } & DocParagraphStyle)
  | { type: 'table'; rows: DocTableRow[]; colWidthsPx?: number[]; borders?: DocBorders; look?: DocTableLook; namedStyleId?: string }
  | { type: 'image'; src: string; width: number; height: number; alt?: string; alignment?: 'left' | 'center' | 'right' }
  | { type: 'horizontalRule' }
  | { type: 'pageBreak' }
  /** Blocks wrapped by a content control (a block-level `w:sdt`). Controls nest. */
  | { type: 'contentControl'; control: DocContentControl; blocks: DocBlock[] }
  /**
   * A table of contents (a `TOC` field): the instruction (`TOC \\o "1-3" \\h \\z \\u`) and the entries it showed when it was last
   * updated, one paragraph each. In a file the field starts in the first entry and ends after the last; the entries link to
   * `_Toc` bookmarks and give their page numbers with `PAGEREF` fields, which a reader that does not update the table keeps as
   * they were.
   */
  | { type: 'toc'; instruction: string; blocks: DocBlock[] };

export interface DocHeaderFooterParagraph {
  content: DocInline[];
  alignment?: 'left' | 'center' | 'right';
}

export interface DocHeaderFooter {
  /** The header/footer of the pages that have no variant of their own (`w:type="default"`; the odd pages when `differentOddEven`). */
  paragraphs: DocHeaderFooterParagraph[];
  /** The first page shows `firstPageParagraphs` (`w:titlePg`; the first page is blank when there are none). */
  differentFirstPage?: boolean;
  firstPageParagraphs?: DocHeaderFooterParagraph[];
  /** Even pages show `evenPageParagraphs` (the document's `w:evenAndOddHeaders`; the even pages are blank when there are none). */
  differentOddEven?: boolean;
  evenPageParagraphs?: DocHeaderFooterParagraph[];
}

/** Text columns of a section (`w:cols`). Kept and written back; the editor lays the text out in one column. */
export interface DocColumns {
  count: number;
  /** Space between columns, in pixels. */
  spacingPx?: number;
  /** A line between the columns. */
  separator?: boolean;
  /** Widths in pixels of columns of different widths (`w:equalWidth="0"`); absent when the columns are equal. */
  widthsPx?: number[];
}

/** Page numbering of a section (`w:pgNumType`). Kept and written back; page numbers shown by the editor do not follow it yet. */
export interface DocPageNumbering {
  /** The number of the section's first page (`w:start`), when it restarts or continues from another number. */
  start?: number;
  /** DOCX number format name of the page numbers (`w:fmt`, e.g. `lowerRoman`); absent for decimal. */
  format?: string;
}

/** Where the text sits between the top and bottom margins of the section's pages (`w:vAlign`; `top` is the default and is not kept). */
export type DocSectionVerticalAlign = 'center' | 'bottom' | 'both';

/**
 * How a section starts relative to the previous one (`w:type` of its `w:sectPr`). The default, a new page, is not kept; the
 * first section has no previous one, so its value is ignored.
 */
export type DocSectionBreakType = 'continuous' | 'evenPage' | 'oddPage';

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
  columns?: DocColumns;
  pageNumbering?: DocPageNumbering;
  verticalAlign?: DocSectionVerticalAlign;
  breakType?: DocSectionBreakType;
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
