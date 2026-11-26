/**
 * FlowBlock Data Structures
 * 
 * Abstract representation of document content that can be measured,
 * laid out across pages, and rendered independently per page.
 * 
 * Content is converted to this intermediate representation before
 * layout and rendering.
 */

// ============================================================================
// Text Runs - Inline content within blocks
// ============================================================================

/** Which note: its kind and its number among the notes of that kind, in document order (1-based). */
export interface NoteKey {
  kind: 'footnote' | 'endnote';
  ordinal: number;
}

/** The text a note is found by in the page (`data-note-ref`, `data-note-target`). */
export function noteKeyString(key: NoteKey): string {
  return `${key.kind}:${key.ordinal}`;
}

/**
 * The content control (a Word structured document tag) a run is inside, innermost one. A control takes two editor
 * positions around its content, which the converter draws as zero-width runs so offsets and positions stay in step.
 */
export interface RunControl {
  /** Alias, tag or kind of the control: its tooltip. */
  title: string;
  /** The content is placeholder text. */
  placeholder?: boolean;
  /** A checkbox control: clicking its glyph toggles it. */
  checkbox?: boolean;
}

/**
 * A text run with formatting
 */
export interface TextRun {
  kind: 'text';
  text: string;
  /** The content control the run is in, if any. */
  control?: RunControl;
  /** The run is the number of a note reference (or the number that starts a note in the notes area): clicking it goes to the other end. */
  noteRef?: NoteKey;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  fontSize?: number;
  fontFamily?: string;
  color?: string;
  backgroundColor?: string;
  superscript?: boolean;
  subscript?: boolean;
  smallCaps?: boolean;
  /** Rendered in capitals (the text itself is unchanged). */
  allCaps?: boolean;
  /** Hidden text (DOCX `vanish`): kept in the flow so every character stays reachable, drawn dimmed. */
  hidden?: boolean;
  /** Thread ID when this run is covered by a comment. */
  commentThreadId?: string;
}

/**
 * A line break within a paragraph
 */
export interface LineBreakRun {
  kind: 'lineBreak';
}

/**
 * An inline image
 */
export interface ImageRun {
  kind: 'image';
  src: string;
  width: number;
  height: number;
  alt?: string;
}

/**
 * A hyperlink
 */
export interface LinkRun {
  kind: 'link';
  text: string;
  href: string;
  /** Drawn as the text around it, with no link colour or underline (the entries of a table of contents). */
  plain?: boolean;
  /** The content control the run is in, if any. */
  control?: RunControl;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: string;
  /** Thread ID when this link is covered by a comment. */
  commentThreadId?: string;
}

/**
 * Union type for all inline content
 */
export type Run = TextRun | LineBreakRun | ImageRun | LinkRun;

// ============================================================================
// Block Types - Top-level content elements
// ============================================================================

/**
 * Text alignment options
 */
export type TextAlignment = 'left' | 'center' | 'right' | 'justify';

/**
 * Paragraph attributes
 */
export interface ParagraphAttrs {
  /** A tab in the text is filled with a dotted leader up to the text after it, which is right-aligned (the entries of a table of contents). */
  tabLeader?: 'dot';
  alignment?: TextAlignment;
  lineHeight?: number;
  spaceBefore?: number;
  spaceAfter?: number;
  firstLineIndent?: number;
  indent?: number;       // Indent level (0, 1, 2, etc.) - each level = 24px
  leftIndent?: number;   // Additional left padding in pixels
  rightIndent?: number;  // Additional right padding in pixels
}

/**
 * A paragraph block
 */
export interface ParagraphBlock {
  kind: 'paragraph';
  /** Titles of the content controls that start with this block, outermost first (drawn as chips in the left margin). */
  controlChips?: string[];
  id: string;
  runs: Run[];
  attrs?: ParagraphAttrs;
  /** ProseMirror position of the paragraph node (set for blocks inside table cells) */
  pmStart?: number;
  pmEnd?: number;
  /** A block of the notes area: it has no node in the document, and clicking it goes to the note's reference. */
  noteTarget?: NoteKey;
  /** The block is drawn from other content (the notes area), not from a node of its own. */
  synthetic?: boolean;
}

/**
 * A heading block
 */
export interface HeadingBlock {
  kind: 'heading';
  /** Titles of the content controls that start with this block, outermost first (drawn as chips in the left margin). */
  controlChips?: string[];
  id: string;
  level: 1 | 2 | 3 | 4 | 5 | 6;
  runs: Run[];
  attrs?: ParagraphAttrs;
  /** ProseMirror position of the heading node (set for blocks inside table cells) */
  pmStart?: number;
  pmEnd?: number;
}

/**
 * A list item block
 */
export interface ListItemBlock {
  kind: 'listItem';
  /** Titles of the content controls that start with this block, outermost first (drawn as chips in the left margin). */
  controlChips?: string[];
  id: string;
  listType: 'bullet' | 'ordered';
  level: number;
  runs: Run[];
  attrs?: ParagraphAttrs;
  /** Index within the list (1-based) for ordered lists */
  listIndex?: number;
  /** ProseMirror position of the item's paragraph (set for blocks inside table cells) */
  pmStart?: number;
  pmEnd?: number;
}

/** One border line of a table or cell (see docs-core `TableBorder`). */
export interface TableBorder {
  style: 'single' | 'double' | 'dashed' | 'dotted' | 'none';
  widthPx?: number;
  color?: string;
}

/** Borders of a table or cell (see docs-core `TableBorders`). */
export interface TableBorders {
  top?: TableBorder;
  right?: TableBorder;
  bottom?: TableBorder;
  left?: TableBorder;
  insideH?: TableBorder;
  insideV?: TableBorder;
}

/**
 * A table cell
 */
export interface TableCell {
  id: string;
  blocks: FlowBlock[];
  colspan?: number;
  rowspan?: number;
  /** Column widths array (for cells that span multiple columns) */
  colwidth?: number[];
  /** ProseMirror position start (set during conversion) */
  pmStart?: number;
  /** ProseMirror position end (set during conversion) */
  pmEnd?: number;
  /** Background color (e.g., "#ff0000" or "rgba(255,0,0,0.5)") */
  backgroundColor?: string;
  /** The cell's own borders; sides not listed take the table's line. */
  borders?: TableBorders;
  verticalAlign?: 'top' | 'center' | 'bottom';
}

/**
 * A table row
 */
export interface TableRow {
  id: string;
  cells: TableCell[];
  height?: number;
  repeatHeader?: boolean;
  cantSplit?: boolean;
}

/**
 * A table block
 */
export interface TableBlock {
  kind: 'table';
  /** Titles of the content controls that start with this block, outermost first (drawn as chips in the left margin). */
  controlChips?: string[];
  id: string;
  rows: TableRow[];
  colWidths?: number[];
  /** Borders of the table; absent for the default grid, `{}` for no lines. */
  borders?: TableBorders;
  /** ProseMirror position of the table node (set for a table nested in a cell) */
  pmStart?: number;
  pmEnd?: number;
}

/**
 * A standalone image block
 */
export interface ImageBlock {
  kind: 'image';
  /** Titles of the content controls that start with this block, outermost first (drawn as chips in the left margin). */
  controlChips?: string[];
  id: string;
  src: string;
  width: number;
  height: number;
  alt?: string;
  alignment?: 'left' | 'center' | 'right';
  /** ProseMirror start position (optional, populated during conversion) */
  pmStart?: number;
  /** ProseMirror end position (optional, populated during conversion) */
  pmEnd?: number;
}

/**
 * A horizontal rule
 */
export interface HorizontalRuleBlock {
  kind: 'horizontalRule';
  /** Titles of the content controls that start with this block, outermost first (drawn as chips in the left margin). */
  controlChips?: string[];
  id: string;
}

/**
 * A page break
 */
export interface PageBreakBlock {
  kind: 'pageBreak';
  /** Titles of the content controls that start with this block, outermost first (drawn as chips in the left margin). */
  controlChips?: string[];
  id: string;
}

/**
 * Union type for all block types
 */
export type FlowBlock =
  | ParagraphBlock
  | HeadingBlock
  | ListItemBlock
  | TableBlock
  | ImageBlock
  | HorizontalRuleBlock
  | PageBreakBlock;

// ============================================================================
// Helper Types
// ============================================================================

/**
 * A block that contains runs (text content)
 */
export type RunContainingBlock = ParagraphBlock | HeadingBlock | ListItemBlock;

/**
 * Type guard for run-containing blocks
 */
export function isRunContainingBlock(block: FlowBlock): block is RunContainingBlock {
  return block.kind === 'paragraph' || block.kind === 'heading' || block.kind === 'listItem';
}

/**
 * Type guard for text runs
 */
export function isTextRun(run: Run): run is TextRun {
  return run.kind === 'text';
}

/**
 * Type guard for link runs
 */
export function isLinkRun(run: Run): run is LinkRun {
  return run.kind === 'link';
}

// ============================================================================
// FlowBlock Factory Functions
// ============================================================================

let blockIdCounter = 0;

/**
 * Generate a unique block ID
 */
export function generateBlockId(): string {
  return `block-${++blockIdCounter}`;
}

/**
 * Reset the block ID counter (useful for testing)
 */
export function resetBlockIdCounter(): void {
  blockIdCounter = 0;
}

/**
 * Create a paragraph block
 */
export function createParagraphBlock(
  runs: Run[],
  attrs?: ParagraphAttrs,
  id?: string
): ParagraphBlock {
  return {
    kind: 'paragraph',
    id: id ?? generateBlockId(),
    runs,
    attrs,
  };
}

/**
 * Create a heading block
 */
export function createHeadingBlock(
  level: 1 | 2 | 3 | 4 | 5 | 6,
  runs: Run[],
  attrs?: ParagraphAttrs,
  id?: string
): HeadingBlock {
  return {
    kind: 'heading',
    id: id ?? generateBlockId(),
    level,
    runs,
    attrs,
  };
}

/**
 * Create a list item block
 */
export function createListItemBlock(
  listType: 'bullet' | 'ordered',
  level: number,
  runs: Run[],
  attrs?: ParagraphAttrs,
  id?: string
): ListItemBlock {
  return {
    kind: 'listItem',
    id: id ?? generateBlockId(),
    listType,
    level,
    runs,
    attrs,
  };
}

/**
 * Create a text run
 */
export function createTextRun(
  text: string,
  formatting?: Omit<TextRun, 'kind' | 'text'>
): TextRun {
  return {
    kind: 'text',
    text,
    ...formatting,
  };
}

/**
 * Create a line break run
 */
export function createLineBreakRun(): LineBreakRun {
  return { kind: 'lineBreak' };
}

/**
 * Get the plain text content of a block
 */
export function getBlockText(block: FlowBlock): string {
  if (!isRunContainingBlock(block)) {
    return '';
  }
  
  return block.runs
    .map(run => {
      if (run.kind === 'text') return run.text;
      if (run.kind === 'link') return run.text;
      if (run.kind === 'lineBreak') return '\n';
      return '';
    })
    .join('');
}

/**
 * Check if a block is empty (no text content)
 */
export function isBlockEmpty(block: FlowBlock): boolean {
  if (!isRunContainingBlock(block)) {
    return false; // Non-text blocks are never "empty"
  }
  
  const text = getBlockText(block);
  return text.trim() === '';
}

