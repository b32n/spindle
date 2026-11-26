/**
 * ProseMirror to FlowBlocks Converter
 * 
 * Converts a ProseMirror document to FlowBlocks for layout and rendering.
 * This is the bridge between the editing layer (ProseMirror) and the
 * presentation layer (layout engine + painter).
 */

import { assignAuthorColors, authorColor, documentRevisionAuthors, readBlockRevision } from '@b32nio/spindle-docs-core';
import { Node as PmNode, Mark } from 'prosemirror-model';
import { sanitizeHref } from './sanitize';

/** The displayed text of a field's cached result (a string, or text/image runs). */
function fieldResultText(result: unknown): string {
  if (typeof result === 'string') return result;
  if (!Array.isArray(result)) return '';
  return result.map((r) => (typeof (r as { text?: unknown })?.text === 'string' ? (r as { text: string }).text : '')).join('');
}
import {
  FlowBlock,
  NoteKey,
  Run,
  RunControl,
  TextRun,
  LinkRun,
  ParagraphBlock,
  HeadingBlock,
  ListItemBlock,
  TableBlock,
  TableRow,
  TableCell,
  ImageBlock,
  HorizontalRuleBlock,
  PageBreakBlock,
  ParagraphAttrs,
  ChangeBar,
  RevisionInfo,
  TextAlignment,
  generateBlockId,
  resetBlockIdCounter,
} from './flow-blocks';

/**
 * Options for conversion
 */
export interface ConversionOptions {
  /** Reset block ID counter before conversion (useful for consistent IDs) */
  resetIds?: boolean;
  /**
   * Leave the footnotes out of the notes area: their text is returned separately (see `SectionedFlowBlocks.footnotes`)
   * for the layout to put at the bottom of the pages their references are on. Only the sectioned conversion supports it.
   */
  footnotesAtPageBottom?: boolean;
}

// ---------------------------------------------------------------------------
// Notes. A note reference is drawn as its number; the text of the notes is collected in an area at the end of the
// section the references are in (footnotes first, then endnotes), after the blocks the conversion made from the nodes.
// ---------------------------------------------------------------------------

interface NoteEntry {
  key: NoteKey;
  label: string;
  node: PmNode;
}

/** The references met so far by the conversion in progress, in document order, and how many there are of each kind. */
let noteEntries: NoteEntry[] = [];
const noteCounts = { footnote: 0, endnote: 0 };

/** The document being converted, and the colours of its tracked-change authors (worked out when the first change is met). */
let convertedDoc: PmNode | null = null;
let authorColors: Map<string, string> | null = null;

function authorColorOf(author: string): string {
  authorColors ??= convertedDoc ? assignAuthorColors(documentRevisionAuthors(convertedDoc)) : new Map();
  return authorColors.get(author) ?? authorColor(author);
}

function resetNotes(doc: PmNode): void {
  convertedDoc = doc;
  authorColors = null;
  noteEntries = [];
  noteCounts.footnote = 0;
  noteCounts.endnote = 0;
}

const ROMAN: Array<[number, string]> = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']];

/** The number a note is shown with: footnotes in digits, endnotes in lower-case roman numerals (Word's defaults). */
export function noteLabel(key: NoteKey): string {
  if (key.kind === 'footnote') return String(key.ordinal);
  let n = key.ordinal;
  let out = '';
  for (const [value, numeral] of ROMAN) {
    while (n >= value) {
      out += numeral;
      n -= value;
    }
  }
  return out;
}

/** The block that draws one note in the notes area (or at the bottom of a page): its number, then its text in a smaller size. */
function noteEntryBlock(entry: NoteEntry): ParagraphBlock {
  const runs: Run[] = [{ kind: 'text', text: entry.label, superscript: true, color: '#1a73e8', noteRef: entry.key }, { kind: 'text', text: ' ' }];
  const body: Run[] = [];
  collectRuns(entry.node, body);
  for (const run of body) {
    if (run.kind === 'text' || run.kind === 'link') runs.push({ ...run, fontSize: (run as TextRun).fontSize ?? 10 } as Run);
    else runs.push(run);
  }
  return { kind: 'paragraph', id: generateBlockId(), synthetic: true, noteTarget: entry.key, runs, attrs: { spaceAfter: 2 } };
}

/** Heading and entries of the notes area for the references found since `from`, or nothing when there are none. */
function notesArea(from: number, withFootnotes = true): FlowBlock[] {
  const blocks: FlowBlock[] = [];
  for (const [kind, title] of [['footnote', 'Footnotes'], ['endnote', 'Endnotes']] as const) {
    if (kind === 'footnote' && !withFootnotes) continue;
    const entries = noteEntries.slice(from).filter((e) => e.key.kind === kind);
    if (entries.length === 0) continue;
    blocks.push({
      kind: 'paragraph',
      id: generateBlockId(),
      synthetic: true,
      runs: [{ kind: 'text', text: title, bold: true, fontSize: 11, color: '#5f6368' }],
      attrs: { spaceBefore: 18, spaceAfter: 4 },
    });
    for (const entry of entries) blocks.push(noteEntryBlock(entry));
  }
  return blocks;
}

/**
 * Convert a ProseMirror document to FlowBlocks
 * 
 * @param doc - The ProseMirror document node
 * @param options - Conversion options
 * @returns Array of FlowBlocks
 */
export function proseMirrorToFlowBlocks(
  doc: PmNode,
  options: ConversionOptions = {}
): FlowBlock[] {
  if (options.resetIds) {
    resetBlockIdCounter();
  }
  resetNotes(doc);
  
  const blocks: FlowBlock[] = [];
  
  doc.forEach((node, _offset) => {
    const block = nodeToFlowBlock(node);
    if (block) {
      if (Array.isArray(block)) {
        blocks.push(...block);
      } else {
        blocks.push(block);
      }
    }
  });
  blocks.push(...notesArea(0));
  
  return blocks;
}

/** The flow blocks of a document and which blocks belong to which section. */
export interface SectionedFlowBlocks {
  blocks: FlowBlock[];
  /**
   * One entry per section in document order: `sectionId` is the id of the `section_break` that ends it (null for the
   * last section, which has none) and `blockStart`..`blockEnd` the slice of `blocks` it owns (possibly empty).
   */
  sections: Array<{ sectionId: string | null; blockStart: number; blockEnd: number }>;
  /** With `footnotesAtPageBottom`: the block that draws each footnote, in document order, and the section (index into `sections`) its reference is in. */
  footnotes: Array<{ block: ParagraphBlock; sectionIndex: number }>;
}

/**
 * Convert a ProseMirror document to FlowBlocks and split them into sections at its `section_break` nodes. A document
 * without breaks is one section holding every block, and `blocks` is what `proseMirrorToFlowBlocks` returns.
 */
export function proseMirrorToSectionedFlowBlocks(doc: PmNode, options: ConversionOptions = {}): SectionedFlowBlocks {
  if (options.resetIds) {
    resetBlockIdCounter();
  }
  resetNotes(doc);
  const blocks: FlowBlock[] = [];
  const sections: SectionedFlowBlocks['sections'] = [];
  const footnotes: SectionedFlowBlocks['footnotes'] = [];
  const atPageBottom = !!options.footnotesAtPageBottom;
  let blockStart = 0;
  let notesFrom = 0;
  /** Close the section being read: the notes of its references follow its blocks. */
  const endSection = (sectionId: string | null) => {
    blocks.push(...notesArea(notesFrom, !atPageBottom));
    if (atPageBottom) {
      for (const entry of noteEntries.slice(notesFrom)) {
        if (entry.key.kind === 'footnote') footnotes.push({ block: noteEntryBlock(entry), sectionIndex: sections.length });
      }
    }
    notesFrom = noteEntries.length;
    sections.push({ sectionId, blockStart, blockEnd: blocks.length });
    blockStart = blocks.length;
  };
  doc.forEach((node) => {
    if (node.type.name === 'section_break') {
      endSection(String(node.attrs.sectionId ?? ''));
      return;
    }
    const block = nodeToFlowBlock(node);
    if (Array.isArray(block)) blocks.push(...block);
    else if (block) blocks.push(block);
  });
  endSection(null);
  return { blocks, sections, footnotes };
}

/**
 * Convert a single ProseMirror node to a FlowBlock (or array of blocks for lists)
 */
function nodeToFlowBlock(node: PmNode): FlowBlock | FlowBlock[] | null {
  switch (node.type.name) {
    case 'paragraph':
      return convertParagraph(node);
    
    case 'heading':
      return convertHeading(node);
    
    case 'bullet_list':
      return convertList(node, 'bullet');
    
    case 'ordered_list':
      return convertList(node, 'ordered');
    
    case 'table':
      return convertTable(node);
    
    case 'image':
      return convertImage(node);
    
    case 'horizontal_rule':
      return convertHorizontalRule();
    
    case 'page_break':
      return convertPageBreak();

    case 'content_control':
      return convertContentControl(node);

    case 'toc':
      return convertToc(node);
    
    default:
      // Unknown node type - try to extract text content as paragraph
      if (node.isTextblock) {
        return convertGenericTextBlock(node);
      }
      return null;
  }
}

const CONTROL_TYPE_LABELS: Record<string, string> = {
  richText: 'Rich text',
  text: 'Text',
  dropDown: 'Drop-down list',
  comboBox: 'Combo box',
  date: 'Date',
  checkbox: 'Check box',
  picture: 'Picture',
  docPart: 'Building block',
};

/** The title a content control is shown with: its alias, else its tag, else its kind. */
function controlTitle(attrs: Record<string, unknown>): string {
  const alias = typeof attrs.alias === 'string' ? attrs.alias : '';
  const tag = typeof attrs.tag === 'string' ? attrs.tag : '';
  return alias || tag || CONTROL_TYPE_LABELS[String(attrs.controlType)] || 'Content control';
}

/**
 * A content control around blocks draws nothing of its own: its blocks flow as if it were not there, and the first one
 * carries the control's title (a chip in the margin).
 */
function convertContentControl(node: PmNode): FlowBlock[] {
  const blocks: FlowBlock[] = [];
  node.forEach((child) => {
    const block = nodeToFlowBlock(child);
    if (Array.isArray(block)) blocks.push(...block);
    else if (block) blocks.push(block);
  });
  if (blocks.length > 0) blocks[0].controlChips = [controlTitle(node.attrs), ...(blocks[0].controlChips ?? [])];
  return blocks;
}

/**
 * A table of contents flows as its entries: paragraphs whose links look like the text around them and whose tab is a dotted
 * leader up to the page number at the right.
 */
function convertToc(node: PmNode): FlowBlock[] {
  const blocks: FlowBlock[] = [];
  node.forEach((child) => {
    if (child.type.name !== 'paragraph') return;
    const entry = convertParagraph(child);
    entry.attrs = { ...entry.attrs, tabLeader: 'dot' };
    entry.runs = entry.runs.map((run) => (run.kind === 'link' ? { ...run, plain: true } : run));
    blocks.push(entry);
  });
  return blocks;
}

/**
 * Convert a paragraph node to a ParagraphBlock
 */
function convertParagraph(node: PmNode): ParagraphBlock {
  const runs = extractRuns(node);
  const attrs = extractParagraphAttrs(node);
  
  return {
    kind: 'paragraph',
    id: generateBlockId(),
    runs,
    attrs: Object.keys(attrs).length > 0 ? attrs : undefined,
  };
}

/**
 * Convert a heading node to a HeadingBlock
 */
function convertHeading(node: PmNode): HeadingBlock {
  const runs = extractRuns(node);
  const attrs = extractParagraphAttrs(node);
  const level = (node.attrs.level as 1 | 2 | 3 | 4 | 5 | 6) || 1;
  
  return {
    kind: 'heading',
    id: generateBlockId(),
    level,
    runs,
    attrs: Object.keys(attrs).length > 0 ? attrs : undefined,
  };
}

/**
 * Convert a list (bullet_list or ordered_list) to ListItemBlocks
 * 
 * Note: Lists are flattened - each list item becomes a separate block
 * with a level indicator for nesting and listIndex for ordered lists.
 */
function convertList(node: PmNode, listType: 'bullet' | 'ordered'): ListItemBlock[] {
  const items: ListItemBlock[] = [];
  // A numbered list may begin at a number other than 1.
  const start = typeof node.attrs.start === 'number' && node.attrs.start > 0 ? Math.trunc(node.attrs.start) : 1;
  let topLevelIndex = start - 1;
  
  node.forEach((listItem) => {
    topLevelIndex++;
    const nestedItems = convertListItem(listItem, listType, 0, topLevelIndex);
    items.push(...nestedItems);
  });
  
  return items;
}

/**
 * Convert a single list item, handling nested lists
 */
function convertListItem(
  node: PmNode,
  listType: 'bullet' | 'ordered',
  level: number,
  listIndex: number
): ListItemBlock[] {
  const items: ListItemBlock[] = [];
  
  // Extract runs from the first paragraph in the list item
  let runs: Run[] = [];
  const attrs: ParagraphAttrs = {};
  let nestedBulletIndex = 0;
  let nestedOrderedIndex = 0;
  
  node.forEach((child, _offset, index) => {
    if (index === 0 && (child.type.name === 'paragraph' || child.isTextblock)) {
      // First child is the content
      runs = extractRuns(child);
      Object.assign(attrs, extractParagraphAttrs(child));
    } else if (child.type.name === 'bullet_list') {
      // Nested bullet list
      child.forEach((nestedItem) => {
        nestedBulletIndex++;
        items.push(...convertListItem(nestedItem, 'bullet', level + 1, nestedBulletIndex));
      });
    } else if (child.type.name === 'ordered_list') {
      // Nested ordered list
      child.forEach((nestedItem) => {
        nestedOrderedIndex++;
        items.push(...convertListItem(nestedItem, 'ordered', level + 1, nestedOrderedIndex));
      });
    }
  });
  
  // Create the main list item block
  const mainItem: ListItemBlock = {
    kind: 'listItem',
    id: generateBlockId(),
    listType,
    level,
    runs,
    attrs: Object.keys(attrs).length > 0 ? attrs : undefined,
    listIndex: listType === 'ordered' ? listIndex : undefined,
  };
  
  // Insert main item first, then nested items
  return [mainItem, ...items];
}

/**
 * Convert a table node to a TableBlock
 * Note: PM positions are tracked separately via createBlockPositionMapWithCells
 */
function convertTable(node: PmNode): TableBlock {
  const rows: TableRow[] = [];
  
  node.forEach((rowNode) => {
    if (rowNode.type.name === 'table_row') {
      const cells: TableCell[] = [];
      
      rowNode.forEach((cellNode) => {
        if (cellNode.type.name === 'table_cell' || cellNode.type.name === 'table_header') {
          // Convert cell content to blocks
          const cellBlocks: FlowBlock[] = [];
          cellNode.forEach((contentNode) => {
            const block = nodeToFlowBlock(contentNode);
            if (block) {
              if (Array.isArray(block)) {
                cellBlocks.push(...block);
              } else {
                cellBlocks.push(block);
              }
            }
          });
          
          cells.push({
            id: generateBlockId(),
            blocks: cellBlocks,
            colspan: cellNode.attrs.colspan as number | undefined,
            rowspan: cellNode.attrs.rowspan as number | undefined,
            colwidth: cellNode.attrs.colwidth as number[] | undefined,
            backgroundColor: cellNode.attrs.backgroundColor as string | undefined,
            ...(cellNode.attrs.borders ? { borders: cellNode.attrs.borders as TableCell['borders'] } : {}),
            ...(cellNode.attrs.verticalAlign ? { verticalAlign: cellNode.attrs.verticalAlign as TableCell['verticalAlign'] } : {}),
            // pmStart and pmEnd will be set by createBlockPositionMapWithCells
          });
        }
      });
      
      rows.push({
        id: generateBlockId(),
        cells,
        ...(rowNode.attrs.repeatHeader ? { repeatHeader: true } : {}),
        ...(rowNode.attrs.cantSplit ? { cantSplit: true } : {}),
        ...rowRevision(readBlockRevision(rowNode.attrs)),
      });
    }
  });
  
  return {
    kind: 'table',
    id: generateBlockId(),
    rows,
    ...(node.attrs.borders ? { borders: node.attrs.borders as TableBlock['borders'] } : {}),
  };
}

/**
 * Convert an image node to an ImageBlock
 */
function convertImage(node: PmNode): ImageBlock {
  return {
    kind: 'image',
    id: generateBlockId(),
    src: node.attrs.src as string || '',
    width: node.attrs.width as number || 200,
    height: node.attrs.height as number || 200,
    alt: node.attrs.alt as string | undefined,
    alignment: node.attrs.alignment as 'left' | 'center' | 'right' | undefined,
  };
}

/**
 * Convert a horizontal rule to a HorizontalRuleBlock
 */
function convertHorizontalRule(): HorizontalRuleBlock {
  return {
    kind: 'horizontalRule',
    id: generateBlockId(),
  };
}

/**
 * Convert a page break to a PageBreakBlock
 */
function convertPageBreak(): PageBreakBlock {
  return {
    kind: 'pageBreak',
    id: generateBlockId(),
  };
}

/**
 * Convert an unknown textblock to a paragraph
 */
function convertGenericTextBlock(node: PmNode): ParagraphBlock {
  const runs = extractRuns(node);
  
  return {
    kind: 'paragraph',
    id: generateBlockId(),
    runs,
  };
}

/**
 * Extract runs from a text-containing node
 */
function extractRuns(node: PmNode): Run[] {
  const runs: Run[] = [];
  collectRuns(node, runs);
  return runs;
}

/** The runs of a node's inline children; `control` is the content control they are inside, if any. */
function collectRuns(node: PmNode, runs: Run[], control?: RunControl): void {
  node.forEach((child) => {
    if (child.isText) {
      const textRun = createTextRunFromNode(child);
      if (control) textRun.control = control;
      // A line break inside a tracked run is kept in the text (the break node cannot carry the change through
      // the Y binding); it still breaks the line.
      if (child.text && child.text.includes('\n')) {
        child.text.split('\n').forEach((part, i) => {
          if (i > 0) runs.push({ kind: 'lineBreak' });
          if (part) runs.push({ ...textRun, text: part });
        });
      } else {
        runs.push(textRun);
      }
    } else if (child.type.name === 'hard_break') {
      runs.push({ kind: 'lineBreak' });
    } else if (child.type.name === 'dynamic_field') {
      // Body fields are resolved at paint time only in headers/footers; in the
      // body they render as a visible placeholder so they stay selectable.
      runs.push({ kind: 'text', text: `\u2039${String(child.attrs.fieldType)}\u203A`, ...(control && { control }) });
    } else if (child.type.name === 'bookmark') {
      // Zero width on the page, but one editor position: a zero-width space keeps the
      // layout offsets and ProseMirror positions in step, and marks where to scroll to.
      runs.push({ kind: 'text', text: '\u200B', ...(control && { control }) });
    } else if (child.type.name === 'note_ref') {
      // The reference shows its number. It is one unit of the document that takes as many positions as the note's text
      // plus its two tokens, so the number is padded with zero-width spaces to keep layout offsets and positions in step.
      const kind = child.attrs.kind === 'endnote' ? 'endnote' : 'footnote';
      const key: NoteKey = { kind, ordinal: ++noteCounts[kind] };
      const custom = typeof child.attrs.customMark === 'string' && child.attrs.customMark ? child.attrs.customMark : '';
      const label = custom || noteLabel(key);
      noteEntries.push({ key, label, node: child });
      runs.push({ kind: 'text', text: label, superscript: true, color: '#1a73e8', noteRef: key, ...(control && { control }) });
      if (child.nodeSize > label.length) runs.push({ kind: 'text', text: '\u200B'.repeat(child.nodeSize - label.length), ...(control && { control }) });
    } else if (child.type.name === 'content_control_inline') {
      // The control's opening and closing tokens are two editor positions around its content, so each is a
      // zero-width space too; the content between them is tagged with the control.
      const inner: RunControl = {
        title: controlTitle(child.attrs),
        ...(child.attrs.showingPlaceholder === true && { placeholder: true }),
        ...(child.attrs.controlType === 'checkbox' && { checkbox: true }),
      };
      runs.push({ kind: 'text', text: '\u200B', control: inner });
      collectRuns(child, runs, inner);
      runs.push({ kind: 'text', text: '\u200B', control: inner });
    } else if (child.type.name === 'field') {
      // A field shows the result Word last computed; it is never recalculated here.
      runs.push({ kind: 'text', text: fieldResultText(child.attrs.result), ...(control && { control }) });
    } else if (child.type.name === 'image' || child.type.name === 'inline_image') {
      runs.push({
        kind: 'image',
        src: child.attrs.src as string || '',
        width: child.attrs.width as number || 100,
        height: child.attrs.height as number || 100,
        alt: child.attrs.alt as string | undefined,
      });
    }
  });
}

/**
 * Create a TextRun (or LinkRun) from a ProseMirror text node with marks
 */
function createTextRunFromNode(node: PmNode): TextRun | LinkRun {
  const run: TextRun = {
    kind: 'text',
    text: node.text || '',
  };

  // Apply marks. A `link` mark promotes the run to a LinkRun so an anchor is
  // actually produced downstream; its href is sanitized to block javascript:
  // and other XSS vectors before it reaches the model.
  let linkHref: string | undefined;
  if (node.marks && node.marks.length > 0) {
    for (const mark of node.marks) {
      applyMarkToRun(run, mark);
      if (mark.type.name === 'link') {
        linkHref = sanitizeHref(mark.attrs.href as string | undefined);
      }
    }
  }

  if (linkHref !== undefined) {
    return {
      kind: 'link',
      text: run.text,
      href: linkHref,
      bold: run.bold,
      italic: run.italic,
      underline: run.underline,
      color: run.color,
      commentThreadId: run.commentThreadId,
      ...(run.revision && { revision: run.revision }),
    };
  }

  return run;
}

/**
 * Apply a ProseMirror mark to a TextRun
 */
function applyMarkToRun(run: TextRun, mark: Mark): void {
  switch (mark.type.name) {
    case 'bold':
    case 'strong':
      run.bold = true;
      break;
    
    case 'italic':
    case 'em':
      run.italic = true;
      break;
    
    case 'underline':
      run.underline = true;
      break;
    
    case 'strikethrough':
    case 'strike':
      run.strikethrough = true;
      break;
    
    case 'superscript':
      run.superscript = true;
      break;
    
    case 'subscript':
      run.subscript = true;
      break;

    case 'smallCaps':
      run.smallCaps = true;
      break;

    case 'allCaps':
      run.allCaps = true;
      break;

    case 'hidden':
      run.hidden = true;
      break;
    
    case 'textStyle':
      // Handle text style attributes
      if (mark.attrs.fontSize) {
        run.fontSize = mark.attrs.fontSize as number;
      }
      if (mark.attrs.fontFamily) {
        run.fontFamily = mark.attrs.fontFamily as string;
      }
      if (mark.attrs.color) {
        run.color = mark.attrs.color as string;
      }
      if (mark.attrs.backgroundColor) {
        run.backgroundColor = mark.attrs.backgroundColor as string;
      }
      break;
    
    case 'link':
      // Link styling. The mark is promoted to a LinkRun (with a sanitized
      // href) by createTextRunFromNode; here we only carry the visual styling.
      run.color = run.color || '#1a73e8';
      run.underline = true;
      break;

    case 'comment':
      run.commentThreadId = mark.attrs.threadId as string;
      break;

    case 'insertion':
    case 'deletion':
    case 'formatChange':
      run.revision = { ...run.revision, [mark.type.name]: revisionInfo(mark) };
      break;
  }
}

/** The layout's default line-height multiplier, which a paragraph with 'single' spacing uses. */
const SINGLE_LINE_HEIGHT = 1.5;

/**
 * A paragraph's `lineSpacing` attr ('single' or a { type, value } object) as the layout's line-height
 * multiplier, scaled from the single-spacing default. `atLeast` / `exactly` heights have no multiplier
 * equivalent and keep the default.
 */
function lineHeightOf(spacing: unknown): number | undefined {
  if (!spacing || typeof spacing !== 'object') return undefined;
  const { type, value } = spacing as { type?: string; value?: number };
  const factor = type === 'onePointFive' ? 1.5 : type === 'double' ? 2 : type === 'multiple' && typeof value === 'number' ? value : undefined;
  return factor !== undefined && factor > 0 && factor <= 10 ? SINGLE_LINE_HEIGHT * factor : undefined;
}

/**
 * Extract paragraph-level attributes from a node
 */
function extractParagraphAttrs(node: PmNode): ParagraphAttrs {
  const attrs: ParagraphAttrs = {};
  
  if (node.attrs) {
    if (node.attrs.alignment) {
      attrs.alignment = node.attrs.alignment as TextAlignment;
    }
    if (node.attrs.lineHeight !== undefined) {
      attrs.lineHeight = node.attrs.lineHeight as number;
    }
    const lineHeight = lineHeightOf(node.attrs.lineSpacing);
    if (lineHeight !== undefined) attrs.lineHeight = lineHeight;
    if (node.attrs.spaceBefore !== undefined) {
      attrs.spaceBefore = node.attrs.spaceBefore as number;
    }
    if (node.attrs.spaceAfter !== undefined) {
      attrs.spaceAfter = node.attrs.spaceAfter as number;
    }
    if (node.attrs.firstLineIndent !== undefined) {
      attrs.firstLineIndent = node.attrs.firstLineIndent as number;
    }
    // Indent level (0, 1, 2, etc.) - used by indent/outdent buttons
    if (node.attrs.indent !== undefined && node.attrs.indent !== 0) {
      attrs.indent = node.attrs.indent as number;
    }
    if (node.attrs.leftIndent !== undefined) {
      attrs.leftIndent = node.attrs.leftIndent as number;
    }
    if (node.attrs.rightIndent !== undefined) {
      attrs.rightIndent = node.attrs.rightIndent as number;
    }
    const changeBar = changeBarOf(readBlockRevision(node.attrs));
    if (changeBar) attrs.changeBar = changeBar;
  }
  
  return attrs;
}

/** A tracked change of a text run as the layout carries it, with the colour of its author. */
function revisionInfo(mark: Mark): RevisionInfo {
  const moveId = mark.attrs.moveId;
  return {
    id: String(mark.attrs.revisionId),
    author: String(mark.attrs.author),
    color: authorColorOf(String(mark.attrs.author)),
    ...(typeof moveId === 'string' && moveId && { moveId }),
  };
}

type BlockRevisionAttr = { insertion?: { id?: string; author?: string }; deletion?: { id?: string; author?: string }; formatChange?: { id?: string; author?: string } } | null | undefined;

/** The bar for a paragraph with a tracked change: in the colour of the author of the last of them (deletion, insertion, formatting). */
function changeBarOf(revision: unknown): ChangeBar | undefined {
  const r = revision as BlockRevisionAttr;
  if (!r || typeof r !== 'object') return undefined;
  const kinds: ChangeBar['kinds'] = [];
  if (r.insertion) kinds.push('insertion');
  if (r.deletion) kinds.push('deletion');
  if (r.formatChange) kinds.push('format');
  if (kinds.length === 0) return undefined;
  const author = String((r.formatChange ?? r.deletion ?? r.insertion)?.author ?? '');
  return { color: authorColorOf(author), author, kinds };
}

/** The `revision` of a flow table row for a row's `revision` attr (nothing when it has none). */
function rowRevision(revision: unknown): { revision?: NonNullable<TableRow['revision']> } {
  const r = revision as BlockRevisionAttr;
  if (!r || typeof r !== 'object') return {};
  const info = (meta: { id?: string; author?: string } | undefined): RevisionInfo | undefined =>
    meta ? { id: String(meta.id ?? ''), author: String(meta.author ?? ''), color: authorColorOf(String(meta.author ?? '')) } : undefined;
  const insertion = info(r.insertion);
  const deletion = info(r.deletion);
  return insertion || deletion ? { revision: { ...(insertion && { insertion }), ...(deletion && { deletion }) } } : {};
}

/**
 * Create a mapping from FlowBlock IDs to ProseMirror positions
 * Also populates PM positions for table cells directly on the cell objects
 * 
 * This is useful for mapping layout positions back to the editor
 */
export function createBlockPositionMap(
  doc: PmNode,
  blocks: FlowBlock[]
): Map<string, { start: number; end: number }> {
  const map = new Map<string, { start: number; end: number }>();
  
  let blockIndex = 0;

  /** `nodeStart` is the position before `node`. A content control has no block of its own: its children follow in its place. */
  const visit = (node: PmNode, nodeStart: number): void => {
    // A section break draws nothing: it has no FlowBlock to map.
    if (node.type.name === 'section_break') return;
    // The notes area closing a section has no node: the next node's blocks come after it.
    while (blockIndex < blocks.length && (blocks[blockIndex] as { synthetic?: boolean }).synthetic) blockIndex++;
    if (node.type.name === 'content_control' || node.type.name === 'toc') {
      node.forEach((child, childOffset) => visit(child, nodeStart + 1 + childOffset));
      return;
    }
    const nodeEnd = nodeStart + node.nodeSize;
    
    // Handle list nodes specially - they expand to multiple blocks
    if (node.type.name === 'bullet_list' || node.type.name === 'ordered_list') {
      // Collect list item positions with proper offset calculation
      const collectListItemPositions = (listNode: PmNode, listOffset: number) => {
        let currentOffset = listOffset + 1; // +1 to enter the list node
        
        listNode.forEach((child) => {
          if (child.type.name === 'list_item') {
            const itemStart = currentOffset;
            const itemEnd = currentOffset + child.nodeSize;
            
            // Map this list item to its actual position
            if (blockIndex < blocks.length) {
              const block = blocks[blockIndex];
              map.set(block.id, { start: itemStart, end: itemEnd });
              blockIndex++;
            }
            
            // Check for nested lists inside this list item
            let nestedOffset = currentOffset + 1; // Enter list_item
            child.forEach((grandchild) => {
              if (grandchild.type.name === 'bullet_list' || grandchild.type.name === 'ordered_list') {
                collectListItemPositions(grandchild, nestedOffset);
              }
              nestedOffset += grandchild.nodeSize;
            });
            
            currentOffset += child.nodeSize;
          }
        });
      };
      
      collectListItemPositions(node, nodeStart);
    } else if (node.type.name === 'table') {
      // Handle tables - map the table block and populate cell positions
      if (blockIndex < blocks.length) {
        const block = blocks[blockIndex];
        map.set(block.id, { start: nodeStart, end: nodeEnd });
        
        // Populate cell positions on the table block
        if (block.kind === 'table') {
          populateTableCellPositions(node, block, nodeStart);
        }
        
        blockIndex++;
      }
    } else {
      // Regular block - 1:1 mapping
      if (blockIndex < blocks.length) {
        const block = blocks[blockIndex];
        map.set(block.id, { start: nodeStart, end: nodeEnd });
        blockIndex++;
      }
    }
  };

  doc.forEach((node, offset) => visit(node, offset));
  
  return map;
}

/**
 * Populate PM positions on table cells
 * Traverses the PM table structure and assigns positions to matching cells
 */
function populateTableCellPositions(
  tableNode: PmNode,
  tableBlock: TableBlock,
  tableStart: number
): void {
  let rowIndex = 0;
  let currentOffset = tableStart + 1; // +1 to enter the table node
  
  tableNode.forEach((rowNode, rowOffset) => {
    if (rowNode.type.name === 'table_row' && rowIndex < tableBlock.rows.length) {
      const row = tableBlock.rows[rowIndex];
      let cellIndex = 0;
      currentOffset = tableStart + 1 + rowOffset + 1; // Enter the row
      
      rowNode.forEach((cellNode, cellOffset) => {
        if ((cellNode.type.name === 'table_cell' || cellNode.type.name === 'table_header') 
            && cellIndex < row.cells.length) {
          const cell = row.cells[cellIndex];
          const cellStart = currentOffset + cellOffset;
          const cellEnd = cellStart + cellNode.nodeSize;
          
          // Set PM positions directly on the cell
          cell.pmStart = cellStart;
          cell.pmEnd = cellEnd;
          assignCellBlockPositions(cellNode, cell, cellStart + 1);
          
          // Copy cell formatting attributes
          if (cellNode.attrs.backgroundColor) {
            cell.backgroundColor = cellNode.attrs.backgroundColor;
          }
          if (cellNode.attrs.colwidth) {
            cell.colwidth = cellNode.attrs.colwidth;
          }
          
          cellIndex++;
        }
      });
      
      rowIndex++;
    }
  });
}


/**
 * Record where the blocks of a table cell sit in the document (paragraphs, headings, the paragraph of each list item,
 * nested tables and images), in the order `nodeToFlowBlock` produced them: a list is flattened to its items, each
 * followed by the items of its nested lists. A nested table gets its own cells' positions too.
 */
function assignCellBlockPositions(cellNode: PmNode, cell: TableCell, contentStart: number): void {
  let index = 0;
  const take = (): FlowBlock | undefined => cell.blocks[index++];
  const placeItems = (listNode: PmNode, listStart: number) => {
    let itemStart = listStart + 1;
    listNode.forEach((item) => {
      const block = take();
      const paragraph = item.firstChild;
      if (block?.kind === 'listItem' && paragraph) {
        block.pmStart = itemStart + 1;
        block.pmEnd = itemStart + 1 + paragraph.nodeSize;
      }
      let childStart = itemStart + 1;
      item.forEach((child) => {
        if (child.type.name === 'bullet_list' || child.type.name === 'ordered_list') placeItems(child, childStart);
        childStart += child.nodeSize;
      });
      itemStart += item.nodeSize;
    });
  };

  const place = (parent: PmNode, start: number): void => {
    let offset = start;
    parent.forEach((child) => {
      switch (child.type.name) {
        case 'paragraph':
        case 'heading': {
          const block = take();
          if (block && (block.kind === 'paragraph' || block.kind === 'heading')) {
            block.pmStart = offset;
            block.pmEnd = offset + child.nodeSize;
          }
          break;
        }
        case 'bullet_list':
        case 'ordered_list':
          placeItems(child, offset);
          break;
        case 'table': {
          const block = take();
          if (block?.kind === 'table') {
            block.pmStart = offset;
            block.pmEnd = offset + child.nodeSize;
            populateTableCellPositions(child, block, offset);
          }
          break;
        }
        case 'image': {
          const block = take();
          if (block?.kind === 'image') {
            block.pmStart = offset;
            block.pmEnd = offset + child.nodeSize;
          }
          break;
        }
        case 'horizontal_rule':
        case 'page_break':
          take();
          break;
        case 'content_control':
          // No block of its own: the blocks of its children follow in its place.
          place(child, offset + 1);
          break;
        default:
          if (child.isTextblock) take();
      }
      offset += child.nodeSize;
    });
  };
  place(cellNode, contentStart);
}
