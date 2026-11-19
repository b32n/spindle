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

/**
 * A text run with formatting
 */
export interface TextRun {
  kind: 'text';
  text: string;
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
  id: string;
  runs: Run[];
  attrs?: ParagraphAttrs;
}

/**
 * A heading block
 */
export interface HeadingBlock {
  kind: 'heading';
  id: string;
  level: 1 | 2 | 3 | 4 | 5 | 6;
  runs: Run[];
  attrs?: ParagraphAttrs;
}

/**
 * A list item block
 */
export interface ListItemBlock {
  kind: 'listItem';
  id: string;
  listType: 'bullet' | 'ordered';
  level: number;
  runs: Run[];
  attrs?: ParagraphAttrs;
  /** Index within the list (1-based) for ordered lists */
  listIndex?: number;
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
}

/**
 * A table row
 */
export interface TableRow {
  id: string;
  cells: TableCell[];
  height?: number;
}

/**
 * A table block
 */
export interface TableBlock {
  kind: 'table';
  id: string;
  rows: TableRow[];
  colWidths?: number[];
}

/**
 * A standalone image block
 */
export interface ImageBlock {
  kind: 'image';
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
  id: string;
}

/**
 * A page break
 */
export interface PageBreakBlock {
  kind: 'pageBreak';
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

