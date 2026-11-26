import { Node as PmNode, type Mark } from 'prosemirror-model';
import { docsSchema } from './schema';
import type {
  Document,
  Block,
  Section,
  InlineContent,
  BlockRevision,
  RunFormat,
  RunRevision,
  TextRun,
  InlineLink,
  TextStyle,
  TextStylePool,
  ParagraphFormat,
  LineSpacing,
  ParagraphBlock,
  HeadingBlock,
  TableRow,
  TableCell,
  TableBorders,
  TableLook,
  CellTextDirection,
  CellVerticalAlign,
  ListType,
  DynamicFieldType,
  InlineField,
  ContentControl,
  ContentControlType,
} from '../types';

/**
 * Loosely-typed ProseMirror node JSON — the shape Schema.nodeFromJSON()
 * accepts. ProseMirror itself types this position as `any`.
 */
export interface PmNodeJSON {
  type: string;
  attrs?: Record<string, unknown>;
  content?: PmNodeJSON[];
  marks?: PmMarkJSON[];
  text?: string;
  /**
   * Internal marker field, set only on synthetic '__list_item__' nodes.
   * blocksToPmDoc groups consecutive markers into list nodes — the markers
   * themselves are never handed to ProseMirror.
   */
  listType?: ListType;
  /** Internal marker field: first number of a synthetic '__list_item__' list. */
  listStart?: number;
}

/** Loosely-typed ProseMirror mark JSON. */
export interface PmMarkJSON {
  type: string;
  attrs?: Record<string, unknown>;
}

/**
 * Convert our document model to ProseMirror JSON format
 */
export function documentToProseMirror(doc: Document): PmNodeJSON {
  const content = doc.sections.flatMap((section, index) => {
    const nodes = section.blocks
      .map(block => blockToPmNode(block))
      .filter((node): node is PmNodeJSON => node !== null);
    // Every section but the last ends with a break that names it.
    return index < doc.sections.length - 1 ? [...nodes, sectionBreakNode(section.id)] : nodes;
  });

  return {
    type: 'doc',
    content: content.length > 0 ? content : [{ type: 'paragraph', content: [] }],
  };
}

/** The `section_break` node that ends the section with this id. */
function sectionBreakNode(sectionId: string): PmNodeJSON {
  return { type: 'section_break', attrs: { sectionId } };
}

/**
 * Convert a section to ProseMirror JSON format
 */
export function sectionToProseMirror(section: Section): PmNodeJSON {
  const content = section.blocks
    .map(block => blockToPmNode(block))
    .filter((node): node is PmNodeJSON => node !== null);

  return {
    type: 'doc',
    content: content.length > 0 ? content : [{ type: 'paragraph', content: [] }],
  };
}

/** PM `lineSpacing` attr: 'single' (the default) or the LineSpacing object. */
function lineSpacingAttr(spacing: LineSpacing | undefined): 'single' | LineSpacing {
  return spacing && spacing.type !== 'single' ? { ...spacing } : 'single';
}

/** The indent and pagination attrs shared by paragraphs and headings (null when unset). */
function formatAttrs(block: ParagraphFormat): Record<string, unknown> {
  return {
    leftIndent: block.leftIndent ?? null,
    rightIndent: block.rightIndent ?? null,
    firstLineIndent: block.firstLineIndent ?? null,
    keepWithNext: block.keepWithNext || null,
    keepTogether: block.keepTogether || null,
    pageBreakBefore: block.pageBreakBefore || null,
    namedStyleId: block.namedStyleId ?? null,
    revision: block.revision ? cloneJson(block.revision) : null,
  };
}

/** A plain-data copy (revisions are stored in node attributes, which must not share objects with the model). */
function cloneJson<T>(value: T): T {
  return JSON.parse(JSON.stringify(value)) as T;
}

/**
 * Read paragraph formatting back from a PM node's attrs. A paragraph's attrs have defaults (spacing after 8,
 * before 0) that mean "unset"; a heading's are null when unset, so an explicit 0 survives for both.
 */
function readFormat(attrs: Record<string, unknown>, defaults: { spaceBefore: number | null; spaceAfter: number | null }): ParagraphFormat {
  const out: ParagraphFormat = {};
  if (attrs.alignment && attrs.alignment !== 'left') out.alignment = attrs.alignment as ParagraphFormat['alignment'];
  if (attrs.lineSpacing && typeof attrs.lineSpacing === 'object') out.lineSpacing = { ...(attrs.lineSpacing as LineSpacing) };
  if (typeof attrs.spaceBefore === 'number' && attrs.spaceBefore !== defaults.spaceBefore) out.spaceBefore = attrs.spaceBefore;
  if (typeof attrs.spaceAfter === 'number' && attrs.spaceAfter !== defaults.spaceAfter) out.spaceAfter = attrs.spaceAfter;
  if (typeof attrs.leftIndent === 'number') out.leftIndent = attrs.leftIndent;
  if (typeof attrs.rightIndent === 'number') out.rightIndent = attrs.rightIndent;
  if (typeof attrs.firstLineIndent === 'number') out.firstLineIndent = attrs.firstLineIndent;
  if (attrs.keepWithNext === true) out.keepWithNext = true;
  if (attrs.keepTogether === true) out.keepTogether = true;
  if (attrs.pageBreakBefore === true) out.pageBreakBefore = true;
  if (typeof attrs.namedStyleId === 'string' && attrs.namedStyleId) out.namedStyleId = attrs.namedStyleId;
  if (attrs.revision && typeof attrs.revision === 'object') out.revision = cloneJson(attrs.revision as BlockRevision);
  return out;
}

const CONTROL_TYPES: readonly ContentControlType[] = ['richText', 'text', 'dropDown', 'comboBox', 'date', 'checkbox', 'picture', 'docPart'];

/** The attrs of a `content_control` / `content_control_inline` node for a content control (null/false when unset). */
function controlAttrs(control: ContentControl): Record<string, unknown> {
  return {
    controlType: control.type,
    tag: control.tag ?? null,
    alias: control.alias ?? null,
    lock: control.lock ?? null,
    placeholder: control.placeholder ?? null,
    showingPlaceholder: control.showingPlaceholder ?? false,
    items: control.items ? control.items.map((i) => ({ ...i })) : null,
    dateFormat: control.dateFormat ?? null,
    dateValue: control.dateValue ?? null,
    checked: control.checked ?? false,
    multiline: control.multiline ?? false,
    gallery: control.gallery ?? null,
  };
}

/** The `{ text, value }` choices of a `items` attr (anything else in it is dropped). */
function controlItems(items: unknown[]): Array<{ text: string; value: string }> {
  return items.flatMap((i) => {
    const item = i as { text?: unknown; value?: unknown } | null;
    return item && typeof item.text === 'string' && typeof item.value === 'string' ? [{ text: item.text, value: item.value }] : [];
  });
}

/** The content control a `content_control` / `content_control_inline` node's attrs describe. */
function readControl(attrs: Record<string, unknown>): ContentControl {
  const type = CONTROL_TYPES.find((t) => t === attrs.controlType) ?? 'richText';
  const str = (v: unknown): string | undefined => (typeof v === 'string' && v ? v : undefined);
  const lock = attrs.lock === 'sdtLocked' || attrs.lock === 'contentLocked' || attrs.lock === 'sdtContentLocked' ? attrs.lock : undefined;
  return {
    type,
    ...(str(attrs.tag) && { tag: str(attrs.tag) }),
    ...(str(attrs.alias) && { alias: str(attrs.alias) }),
    ...(lock && { lock }),
    ...(str(attrs.placeholder) && { placeholder: str(attrs.placeholder) }),
    ...(attrs.showingPlaceholder === true && { showingPlaceholder: true }),
    ...(Array.isArray(attrs.items) && controlItems(attrs.items).length > 0 && { items: controlItems(attrs.items) }),
    ...(str(attrs.dateFormat) && { dateFormat: str(attrs.dateFormat) }),
    ...(str(attrs.dateValue) && { dateValue: str(attrs.dateValue) }),
    ...(type === 'checkbox' && { checked: attrs.checked === true }),
    ...(attrs.multiline === true && { multiline: true }),
    ...(str(attrs.gallery) && { gallery: str(attrs.gallery) }),
  } as ContentControl;
}

/**
 * Convert a single block to ProseMirror node JSON
 */
function blockToPmNode(block: Block, stylePool?: TextStylePool): PmNodeJSON | null {
  switch (block.type) {
    case 'paragraph':
      return {
        type: 'paragraph',
        attrs: {
          alignment: block.alignment || 'left',
          indent: block.indent || 0,
          lineSpacing: lineSpacingAttr(block.lineSpacing),
          spaceBefore: block.spaceBefore || 0,
          spaceAfter: block.spaceAfter ?? 8,
          ...formatAttrs(block),
        },
        content: inlineContentToPm(block.content, stylePool),
      };
      
    case 'heading':
      return {
        type: 'heading',
        attrs: {
          level: block.level,
          alignment: block.alignment || 'left',
          lineSpacing: lineSpacingAttr(block.lineSpacing),
          spaceBefore: block.spaceBefore ?? null,
          spaceAfter: block.spaceAfter ?? null,
          ...formatAttrs(block),
        },
        content: inlineContentToPm(block.content, stylePool),
      };
      
    case 'list-item':
      // Return just the list_item node - grouping into lists is handled by blocksToPmDoc
      return {
        type: '__list_item__',  // Special marker - will be grouped into lists by blocksToPmDoc
        listType: block.listType,
        listStart: block.start,
        content: [{
          type: 'list_item',
          attrs: { level: block.level ?? 0 }, // preserve nesting depth across the round-trip
          content: [{
            type: 'paragraph',
            attrs: {
              alignment: block.alignment || 'left',
              lineSpacing: lineSpacingAttr(block.lineSpacing),
              spaceBefore: block.spaceBefore || 0,
              spaceAfter: block.spaceAfter ?? 8,
              ...formatAttrs(block),
            },
            content: inlineContentToPm(block.content, stylePool),
          }],
        }],
      };
      
    case 'table':
      return {
        type: 'table',
        attrs: { colWidths: block.colWidths ?? null, styleId: block.styleId ?? null, borders: block.borders ?? null, look: block.look ?? null, namedStyleId: block.namedStyleId ?? null },
        content: block.rows.map(row => ({
          type: 'table_row',
          attrs: { height: row.height ?? null, repeatHeader: row.repeatHeader ?? false, cantSplit: row.cantSplit ?? false, revision: row.revision ? cloneJson(row.revision) : null },
          content: row.cells.map(cell => ({
            type: cell.isHeader ? 'table_header' : 'table_cell',
            attrs: {
              colspan: cell.colspan || 1,
              rowspan: cell.rowspan || 1,
              styleId: cell.styleId ?? null,
              backgroundColor: cell.backgroundColor ?? null,
              colwidth: cell.colwidth ?? null,
              borders: cell.borders ?? null,
              verticalAlign: cell.verticalAlign ?? null,
              textDirection: cell.textDirection ?? null,
            },
            content: cell.blocks && cell.blocks.length > 0
              ? blocksToPmContent(cell.blocks, stylePool)
              : [{ type: 'paragraph', content: inlineContentToPm(cell.content, stylePool) }],
          })),
        })),
      };
      
    case 'image':
      return {
        type: 'image',
        attrs: {
          src: block.src,
          alt: block.alt || '',
          title: block.title ?? null,
          width: block.width,
          height: block.height,
          alignment: block.alignment || 'center',
        },
      };
      
    case 'horizontal-rule':
      return { type: 'horizontal_rule' };
      
    case 'page-break':
      return { type: 'page_break' };

    case 'content-control': {
      const content = blocksToPmContent(block.blocks, stylePool);
      return {
        type: 'content_control',
        attrs: controlAttrs(block.control),
        content: content.length > 0 ? content : [{ type: 'paragraph', content: [] }],
      };
    }

    case 'toc':
      // The entries are paragraphs (anything else a hand-made table of contents holds is not an entry).
      return {
        type: 'toc',
        attrs: { instruction: block.instruction },
        content: blocksToPmContent(block.blocks.filter((b) => b.type === 'paragraph'), stylePool),
      };
      
    default:
      return null;
  }
}

/**
 * Convert a TextStyle to ProseMirror marks
 */
function styleToMarks(style: TextStyle): PmMarkJSON[] {
  const marks: PmMarkJSON[] = [];
  
  if (style.bold) marks.push({ type: 'bold' });
  if (style.italic) marks.push({ type: 'italic' });
  if (style.underline) marks.push({ type: 'underline' });
  if (style.strikethrough) marks.push({ type: 'strikethrough' });
  if (style.superscript) marks.push({ type: 'superscript' });
  if (style.subscript) marks.push({ type: 'subscript' });
  if (style.smallCaps) marks.push({ type: 'smallCaps' });
  if (style.allCaps) marks.push({ type: 'allCaps' });
  if (style.hidden) marks.push({ type: 'hidden' });
  
  // Add textStyle mark if any text style attributes are set
  const textStyleAttrs: Record<string, unknown> = {};
  if (style.color) textStyleAttrs.color = style.color;
  if (style.backgroundColor) textStyleAttrs.backgroundColor = style.backgroundColor;
  if (style.fontSize) textStyleAttrs.fontSize = style.fontSize;
  if (style.fontFamily) textStyleAttrs.fontFamily = style.fontFamily;
  if (style.fontFamilyEastAsia) textStyleAttrs.fontFamilyEastAsia = style.fontFamilyEastAsia;
  if (style.fontFamilyComplex) textStyleAttrs.fontFamilyComplex = style.fontFamilyComplex;
  
  if (Object.keys(textStyleAttrs).length > 0) {
    marks.push({ type: 'textStyle', attrs: textStyleAttrs });
  }
  
  return marks;
}

/** The character-formatting fields of a text run or link (same names in TextRun, InlineLink and TextStyle). */
function inlineStyleOf(item: TextStyle): TextStyle {
  const style: TextStyle = {};
  if (item.bold) style.bold = true;
  if (item.italic) style.italic = true;
  if (item.underline) style.underline = true;
  if (item.strikethrough) style.strikethrough = true;
  if (item.superscript) style.superscript = true;
  if (item.subscript) style.subscript = true;
  if (item.smallCaps) style.smallCaps = true;
  if (item.allCaps) style.allCaps = true;
  if (item.hidden) style.hidden = true;
  if (item.color) style.color = item.color;
  if (item.backgroundColor) style.backgroundColor = item.backgroundColor;
  if (item.fontSize) style.fontSize = item.fontSize;
  if (item.fontFamily) style.fontFamily = item.fontFamily;
  if (item.fontFamilyEastAsia) style.fontFamilyEastAsia = item.fontFamilyEastAsia;
  if (item.fontFamilyComplex) style.fontFamilyComplex = item.fontFamilyComplex;
  return style;
}

/**
 * Emit a text run as PM nodes. A '\n' in the model is a line break, which PM
 * represents as a `hard_break` node (PM text nodes cannot contain it).
 */
function textWithBreaks(text: string, marks: PmMarkJSON[]): PmNodeJSON[] {
  // A break inside a tracked run stays in the text node: the Y binding keeps
  // marks on text only, so a marked `hard_break` would lose its change.
  if (text.includes('\n') && marks.some((m) => m.type === 'insertion' || m.type === 'deletion')) {
    return [{ type: 'text', text, marks }];
  }
  const out: PmNodeJSON[] = [];
  text.split('\n').forEach((part, i) => {
    if (i > 0) out.push({ type: 'hard_break' });
    if (part.length > 0) {
      out.push(marks.length > 0 ? { type: 'text', text: part, marks } : { type: 'text', text: part });
    }
  });
  return out;
}

/**
 * Convert inline content array to ProseMirror content
 * @param content - The inline content array
 * @param stylePool - Optional style pool for looking up styles by ID
 */
function inlineContentToPm(content: InlineContent[], stylePool?: TextStylePool): PmNodeJSON[] {
  if (!content || content.length === 0) {
    return [];
  }
  
  return content.map(item => {
    if (item.type === 'text') {
      let marks: PmMarkJSON[] = [];
      
      // Check if using styleId (style pool reference)
      if (item.styleId && stylePool) {
        const style = stylePool.get(item.styleId);
        if (style) {
          marks = styleToMarks(style);
        }
      } else {
        // Use inline styles
        const inlineStyle = inlineStyleOf(item);
        marks = styleToMarks(inlineStyle);
      }

      if (item.namedStyleId) marks.push({ type: 'namedStyle', attrs: { id: item.namedStyleId } });
      if (item.commentThreadId) {
        marks.push({ type: 'comment', attrs: { threadId: item.commentThreadId } });
      }
      marks.push(...revisionMarks(item.revision));

      return textWithBreaks(item.text, marks);
    }
    
    if (item.type === 'link') {
      const linkAttrs: Record<string, unknown> = { href: item.href };
      if (item.title) linkAttrs.title = item.title;
      const marks: PmMarkJSON[] = [{ type: 'link', attrs: linkAttrs }];
      // Keep character formatting that was applied on top of the link.
      const poolStyle = item.styleId && stylePool ? stylePool.get(item.styleId) : undefined;
      marks.push(...styleToMarks(poolStyle ?? inlineStyleOf(item)));
      if (item.namedStyleId) marks.push({ type: 'namedStyle', attrs: { id: item.namedStyleId } });
      if (item.commentThreadId) {
        marks.push({ type: 'comment', attrs: { threadId: item.commentThreadId } });
      }
      marks.push(...revisionMarks(item.revision));
      return textWithBreaks(item.text, marks);
    }

    if (item.type === 'image') {
      return {
        type: 'inline_image',
        attrs: {
          src: item.src,
          alt: item.alt ?? '',
          width: item.width,
          height: item.height,
        },
      };
    }

    if (item.type === 'bookmark') {
      return { type: 'bookmark', attrs: { name: item.name } };
    }

    if (item.type === 'field') {
      return {
        type: 'field',
        attrs: {
          instruction: item.instruction,
          kind: item.kind,
          result: item.result,
          dirty: item.dirty ?? false,
        },
      };
    }

    if (item.type === 'dynamicField') {
      return {
        type: 'dynamic_field',
        attrs: { fieldType: item.fieldType, format: item.format ?? null },
      };
    }

    if (item.type === 'noteRef') {
      return {
        type: 'note_ref',
        attrs: { kind: item.kind, noteId: item.id, customMark: item.customMark ?? null },
        content: inlineContentToPm(item.content.filter((c) => c.type !== 'noteRef' && c.type !== 'contentControl'), stylePool),
      };
    }

    if (item.type === 'contentControl') {
      return {
        type: 'content_control_inline',
        attrs: controlAttrs(item.control),
        content: inlineContentToPm(item.content, stylePool),
      };
    }

    return null;
  }).flat().filter((node): node is PmNodeJSON => node !== null);
}

/**
 * Convert ProseMirror document back to our document model
 *
 * A document without `section_break` nodes becomes one section: `existingDoc.sections[0]` with the new blocks. With
 * breaks, each break ends the section whose id it names (looked up in `existingDoc.sections`, so page setup, header and
 * footer carry over) and the blocks after the last break form the last section, which is the last of
 * `existingDoc.sections` that no break names. A break naming an id that `existingDoc` does not have gets a section
 * that copies the page setup of the first one, and a repeated id is renamed `<id>_2`, `<id>_3`, ... so section ids
 * stay unique.
 * @param pmDoc - The ProseMirror document
 * @param existingDoc - The existing document to update
 * @param stylePool - Optional style pool for deduplicating styles (recommended for large documents)
 */
export function proseMirrorToDocument(
  pmDoc: PmNode, 
  existingDoc: Document,
  stylePool?: TextStylePool
): Document {
  const groups: { breakId: string | null; blocks: Block[] }[] = [];
  let blocks: Block[] = [];
  let blockIndex = 0;
  
  let previous: PmNode | undefined;
  pmDoc.forEach((node) => {
    if (node.type.name === 'section_break') {
      groups.push({ breakId: String(node.attrs.sectionId ?? ''), blocks });
      blocks = [];
    } else {
      const convertedBlocks = pmNodeToBlocks(node, blockIndex, stylePool, previous?.type.name === 'ordered_list');
      blocks.push(...convertedBlocks);
      blockIndex += convertedBlocks.length;
    }
    previous = node;
  });
  groups.push({ breakId: null, blocks });

  const existing = existingDoc.sections;
  if (groups.length === 1) {
    // Ensure at least one paragraph
    if (groups[0].blocks.length === 0) {
      groups[0].blocks.push({
        id: `block_${Date.now()}`,
        type: 'paragraph',
        content: [],
      });
    }
    return {
      ...existingDoc,
      sections: [{
        ...existing[0],
        blocks: groups[0].blocks,
      }],
      updatedAt: new Date().toISOString(),
    };
  }

  // Several sections (a section may hold no blocks: two breaks in a row, or a break at the very end).
  const named = new Set(groups.flatMap(g => (g.breakId === null ? [] : [g.breakId])));
  const trailing = [...existing].reverse().find(s => !named.has(s.id));
  const used = new Set<string>();
  const sections = groups.map((group): Section => {
    const base = group.breakId === null
      ? trailing
      : existing.find(s => s.id === group.breakId);
    let id = group.breakId ?? base?.id ?? `section_${groups.length}`;
    for (let n = 2; used.has(id); n++) id = `${group.breakId ?? base?.id ?? 'section'}_${n}`;
    used.add(id);
    return {
      ...(base ?? { pageConfig: existing[0].pageConfig }),
      id,
      blocks: group.blocks,
    };
  });
  return {
    ...existingDoc,
    sections,
    updatedAt: new Date().toISOString(),
  };
}

/** The paragraphs, headings and list items among blocks, looking inside content controls. */
function textBlocks(blocks: Block[]): Block[] {
  return blocks.flatMap((b) => (b.type === 'content-control' ? textBlocks(b.blocks) : b.type === 'paragraph' || b.type === 'heading' || b.type === 'list-item' ? [b] : []));
}

/** The blocks of a node's children (a table cell or a content control): a list becomes its items. */
function pmChildrenToBlocks(parent: PmNode, startIndex: number, stylePool?: TextStylePool): Block[] {
  const blocks: Block[] = [];
  let previous: PmNode | undefined;
  parent.forEach((child) => {
    blocks.push(...pmNodeToBlocks(child, startIndex + blocks.length, stylePool, previous?.type.name === 'ordered_list'));
    previous = child;
  });
  return blocks;
}

/**
 * The content of a table cell: `content` is the inline content of its first paragraph that has any (what a cell of
 * one paragraph holds), `blocks` the whole cell unless it is one paragraph.
 */
function pmCellContent(cellNode: PmNode, startIndex: number, stylePool?: TextStylePool): { content: InlineContent[]; blocks?: Block[] } {
  const blocks = pmChildrenToBlocks(cellNode, startIndex, stylePool);
  const inline = (block: Block | undefined): InlineContent[] => (block && 'content' in block ? block.content : []);
  const paragraphs = textBlocks(blocks);
  const firstIndex = paragraphs.findIndex((b) => inline(b).some((c) => c.type !== 'bookmark'));
  // Bookmarks of the paragraphs before the first one with content stay in front of it
  const content = firstIndex < 0 ? inline(paragraphs[0]) : [...paragraphs.slice(0, firstIndex).flatMap(inline), ...inline(paragraphs[firstIndex])];
  const plain = blocks.length === 1 && blocks[0].type === 'paragraph';
  return { content, ...(!plain && blocks.length > 0 ? { blocks } : {}) };
}

/**
 * Convert ProseMirror node to our block model
 * Returns an array because some nodes (like lists) expand to multiple blocks
 */
function pmNodeToBlocks(node: PmNode, startIndex: number, stylePool?: TextStylePool, followsNumberedList = false): Block[] {
  switch (node.type.name) {
    case 'paragraph': {
      const block: ParagraphBlock = {
        id: `block_${startIndex}_${Date.now()}`,
        type: 'paragraph',
        content: pmContentToInline(node, stylePool),
      };
      // Extract paragraph attributes
      if (node.attrs.indent && node.attrs.indent !== 0) {
        block.indent = node.attrs.indent;
      }
      Object.assign(block, readFormat(node.attrs, { spaceBefore: 0, spaceAfter: 8 }));
      return [block];
    }
      
    case 'heading': {
      const block: HeadingBlock = {
        id: `block_${startIndex}_${Date.now()}`,
        type: 'heading',
        level: node.attrs.level as 1 | 2 | 3 | 4 | 5 | 6,
        content: pmContentToInline(node, stylePool),
      };
      Object.assign(block, readFormat(node.attrs, { spaceBefore: null, spaceAfter: null }));
      return [block];
    }
      
    case 'bullet_list':
    case 'ordered_list': {
      const blocks: Block[] = [];
      const listType = node.type.name === 'bullet_list' ? 'bullet' : 'numbered';
      
      node.forEach((listItem, _, index) => {
        // Get the first paragraph content from the list item
        let content: InlineContent[] = [];
        let format: ParagraphFormat = {};
        listItem.forEach((child) => {
          if (child.type.name === 'paragraph') {
            content = pmContentToInline(child, stylePool);
            format = readFormat(child.attrs, { spaceBefore: 0, spaceAfter: 8 });
          }
        });
        
        // A list's first number is kept when it is not 1, or when the list follows another numbered list
        // (a restart, which would otherwise merge into the list before it).
        const start = node.attrs.start as number | undefined;
        const first = index === 0 && listType === 'numbered';
        const keepStart = first && start && (start !== 1 || followsNumberedList);
        blocks.push({
          id: `block_${startIndex + index}_${Date.now()}`,
          type: 'list-item',
          listType,
          level: (listItem.attrs.level as number) ?? 0,
          ...(keepStart ? { start } : {}),
          ...format,
          content,
        });
      });
      
      return blocks;
    }
      
    case 'table': {
      const rows: TableRow[] = [];
      node.forEach((rowNode) => {
        const cells: TableCell[] = [];
        rowNode.forEach((cellNode) => {
          const { content, blocks } = pmCellContent(cellNode, startIndex, stylePool);
          cells.push({
            id: `cell_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`,
            content,
            ...(blocks ? { blocks } : {}),
            colspan: cellNode.attrs.colspan,
            rowspan: cellNode.attrs.rowspan,
            ...(cellNode.type.name === 'table_header' ? { isHeader: true } : {}),
            ...(cellNode.attrs.styleId ? { styleId: cellNode.attrs.styleId as string } : {}),
            ...(cellNode.attrs.backgroundColor ? { backgroundColor: cellNode.attrs.backgroundColor as string } : {}),
            ...(cellNode.attrs.colwidth ? { colwidth: cellNode.attrs.colwidth as number[] } : {}),
            ...(cellNode.attrs.borders ? { borders: cellNode.attrs.borders as TableBorders } : {}),
            ...(cellNode.attrs.verticalAlign ? { verticalAlign: cellNode.attrs.verticalAlign as CellVerticalAlign } : {}),
            ...(cellNode.attrs.textDirection ? { textDirection: cellNode.attrs.textDirection as CellTextDirection } : {}),
          });
        });
        rows.push({
          id: `row_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`,
          cells,
          ...(rowNode.attrs.height != null ? { height: rowNode.attrs.height as number } : {}),
          ...(rowNode.attrs.repeatHeader ? { repeatHeader: true } : {}),
          ...(rowNode.attrs.cantSplit ? { cantSplit: true } : {}),
          ...(rowNode.attrs.revision && typeof rowNode.attrs.revision === 'object' ? { revision: cloneJson(rowNode.attrs.revision as TableRow['revision']) } : {}),
        });
      });

      return [{
        id: `block_${startIndex}_${Date.now()}`,
        type: 'table',
        rows,
        ...(node.attrs.colWidths ? { colWidths: node.attrs.colWidths as number[] } : {}),
        ...(node.attrs.styleId ? { styleId: node.attrs.styleId as string } : {}),
        ...(node.attrs.borders ? { borders: node.attrs.borders as TableBorders } : {}),
        ...(node.attrs.look ? { look: node.attrs.look as TableLook } : {}),
        ...(node.attrs.namedStyleId ? { namedStyleId: node.attrs.namedStyleId as string } : {}),
      }];
    }
      
    case 'image':
      return [{
        id: `block_${startIndex}_${Date.now()}`,
        type: 'image',
        src: node.attrs.src,
        alt: node.attrs.alt,
        ...(node.attrs.title ? { title: node.attrs.title as string } : {}),
        width: node.attrs.width || 300,
        height: node.attrs.height || 200,
        alignment: node.attrs.alignment,
      }];
      
    case 'horizontal_rule':
      return [{
        id: `block_${startIndex}_${Date.now()}`,
        type: 'horizontal-rule',
      }];
      
    case 'page_break':
      return [{
        id: `block_${startIndex}_${Date.now()}`,
        type: 'page-break',
      }];

    case 'toc':
      return [{
        id: `block_${startIndex}_${Date.now()}`,
        type: 'toc',
        instruction: String(node.attrs.instruction ?? ''),
        blocks: pmChildrenToBlocks(node, startIndex + 1, stylePool),
      }];

    case 'content_control':
      return [{
        id: `block_${startIndex}_${Date.now()}`,
        type: 'content-control',
        control: readControl(node.attrs),
        blocks: pmChildrenToBlocks(node, startIndex + 1, stylePool),
      }];
      
    default:
      return [];
  }
}

/**
 * Extract TextStyle from ProseMirror marks
 */
function extractStyleFromMarks(marks: readonly Mark[]): TextStyle | null {
  const style: TextStyle = {};
  let hasStyle = false;
  
  for (const mark of marks) {
    switch (mark.type.name) {
      case 'bold':
        style.bold = true;
        hasStyle = true;
        break;
      case 'italic':
        style.italic = true;
        hasStyle = true;
        break;
      case 'underline':
        style.underline = true;
        hasStyle = true;
        break;
      case 'strikethrough':
        style.strikethrough = true;
        hasStyle = true;
        break;
      case 'superscript':
        style.superscript = true;
        hasStyle = true;
        break;
      case 'subscript':
        style.subscript = true;
        hasStyle = true;
        break;
      case 'smallCaps':
        style.smallCaps = true;
        hasStyle = true;
        break;
      case 'allCaps':
        style.allCaps = true;
        hasStyle = true;
        break;
      case 'hidden':
        style.hidden = true;
        hasStyle = true;
        break;
      case 'textStyle':
        if (mark.attrs.color) { style.color = mark.attrs.color; hasStyle = true; }
        if (mark.attrs.backgroundColor) { style.backgroundColor = mark.attrs.backgroundColor; hasStyle = true; }
        if (mark.attrs.fontSize) { style.fontSize = mark.attrs.fontSize; hasStyle = true; }
        if (mark.attrs.fontFamily) { style.fontFamily = mark.attrs.fontFamily; hasStyle = true; }
        if (mark.attrs.fontFamilyEastAsia) { style.fontFamilyEastAsia = mark.attrs.fontFamilyEastAsia; hasStyle = true; }
        if (mark.attrs.fontFamilyComplex) { style.fontFamilyComplex = mark.attrs.fontFamilyComplex; hasStyle = true; }
        break;
    }
  }
  
  return hasStyle ? style : null;
}

/**
 * Convert ProseMirror node content to our inline content format
 * @param node - The ProseMirror node
 * @param stylePool - Optional style pool for deduplicating styles
 */
function pmContentToInline(node: PmNode, stylePool?: TextStylePool): InlineContent[] {
  const content: InlineContent[] = [];

  node.forEach((child) => {
    if (child.type.name === 'inline_image') {
      content.push({
        type: 'image',
        src: child.attrs.src as string,
        width: (child.attrs.width as number | null) ?? 100,
        height: (child.attrs.height as number | null) ?? 100,
        ...(child.attrs.alt ? { alt: child.attrs.alt as string } : {}),
      });
      return;
    }

    if (child.type.name === 'bookmark') {
      content.push({ type: 'bookmark', name: child.attrs.name as string });
      return;
    }

    if (child.type.name === 'field') {
      content.push({
        type: 'field',
        instruction: child.attrs.instruction as string,
        kind: child.attrs.kind as string,
        result: child.attrs.result as InlineField['result'],
        ...(child.attrs.dirty ? { dirty: true } : {}),
      });
      return;
    }

    if (child.type.name === 'note_ref') {
      content.push({
        type: 'noteRef',
        kind: child.attrs.kind === 'endnote' ? 'endnote' : 'footnote',
        id: String(child.attrs.noteId ?? ''),
        content: pmContentToInline(child, stylePool),
        ...(typeof child.attrs.customMark === 'string' && child.attrs.customMark ? { customMark: child.attrs.customMark } : {}),
      });
      return;
    }

    if (child.type.name === 'content_control_inline') {
      content.push({ type: 'contentControl', control: readControl(child.attrs), content: pmContentToInline(child, stylePool) });
      return;
    }

    if (child.type.name === 'dynamic_field') {
      content.push({
        type: 'dynamicField',
        fieldType: child.attrs.fieldType as DynamicFieldType,
        ...(child.attrs.format ? { format: child.attrs.format as string } : {}),
      });
      return;
    }

    // A hard break is a '\n' in the model; fold it into the previous plain
    // run when possible so a paragraph with soft breaks stays compact.
    if (child.type.name === 'hard_break') {
      const prev = content[content.length - 1];
      const breakRevision = revisionOf(child.marks);
      if (breakRevision) {
        content.push({ type: 'text', text: '\n', revision: breakRevision });
      } else if (prev && prev.type === 'text' && !prev.styleId && !hasRunFormatting(prev)) {
        prev.text += '\n';
      } else {
        content.push({ type: 'text', text: '\n' });
      }
      return;
    }

    if (!child.isText) return;

    const text = child.text || '';

    // A comment mark anchors a thread to this run.
    const commentMark = child.marks.find((m) => m.type.name === 'comment');
    const commentThreadId = commentMark ? (commentMark.attrs.threadId as string) : undefined;

    const revision = revisionOf(child.marks);

    const style = extractStyleFromMarks(child.marks);
    const linkMark = child.marks.find((m) => m.type.name === 'link');
    let item: InlineContent;
    if (linkMark) {
      const link: InlineLink = { type: 'link', text, href: linkMark.attrs.href };
      if (linkMark.attrs.title) link.title = linkMark.attrs.title as string;
      if (style && stylePool) link.styleId = stylePool.getOrCreate(style);
      else if (style) Object.assign(link, style);
      item = link;
    } else if (style && stylePool) {
      // Use style pool for deduplication (recommended for large documents)
      item = { type: 'text', text, styleId: stylePool.getOrCreate(style) } as TextRun;
    } else if (style) {
      // Inline styles (fallback when no style pool provided)
      item = { type: 'text', text, ...style } as TextRun;
    } else {
      item = { type: 'text', text } as TextRun;
    }

    const namedStyleMark = child.marks.find((m) => m.type.name === 'namedStyle');
    if (namedStyleMark && typeof namedStyleMark.attrs.id === 'string' && namedStyleMark.attrs.id) {
      (item as TextRun | InlineLink).namedStyleId = namedStyleMark.attrs.id;
    }
    if (commentThreadId) {
      (item as TextRun | InlineLink).commentThreadId = commentThreadId;
    }
    if (revision) (item as TextRun | InlineLink).revision = revision;
    // Text following a folded line break continues the same plain run.
    const prev = content[content.length - 1];
    if (
      item.type === 'text' && !hasRunFormatting(item) &&
      prev && prev.type === 'text' && !hasRunFormatting(prev) && prev.text.endsWith('\n')
    ) {
      prev.text += item.text;
    } else {
      content.push(item);
    }
  });

  return content;
}

/** The `insertion` / `deletion` marks for a run's tracked changes. */
function revisionMarks(revision: RunRevision | undefined): PmMarkJSON[] {
  const marks: PmMarkJSON[] = [];
  for (const kind of ['insertion', 'deletion'] as const) {
    const meta = revision?.[kind];
    if (meta) marks.push({ type: kind, attrs: { revisionId: meta.id, author: meta.author, ...(meta.date && { date: meta.date }), ...(meta.moveId && { moveId: meta.moveId }) } });
  }
  const change = revision?.formatChange;
  if (change) marks.push({ type: 'formatChange', attrs: { revisionId: change.id, author: change.author, ...(change.date && { date: change.date }), previous: cloneJson(change.previous) } });
  return marks;
}

/** The tracked changes a text node's marks record, if any. */
export function revisionOf(marks: readonly { type: { name: string }; attrs: Record<string, unknown> }[]): RunRevision | undefined {
  const revision: RunRevision = {};
  for (const kind of ['insertion', 'deletion'] as const) {
    const mark = marks.find((m) => m.type.name === kind);
    if (mark) {
      const date = mark.attrs.date;
      const moveId = mark.attrs.moveId;
      revision[kind] = { id: String(mark.attrs.revisionId), author: String(mark.attrs.author), ...(typeof date === 'string' && date && { date }), ...(typeof moveId === 'string' && moveId && { moveId }) };
    }
  }
  const change = marks.find((m) => m.type.name === 'formatChange');
  if (change) {
    const date = change.attrs.date;
    const previous = change.attrs.previous;
    revision.formatChange = {
      id: String(change.attrs.revisionId),
      author: String(change.attrs.author),
      ...(typeof date === 'string' && date && { date }),
      previous: previous && typeof previous === 'object' ? cloneJson(previous as RunFormat) : {},
    };
  }
  return revision.insertion || revision.deletion || revision.formatChange ? revision : undefined;
}

/** True if a text run carries any formatting or comment anchor. */
function hasRunFormatting(run: TextRun): boolean {
  const { type: _type, text: _text, ...rest } = run;
  return Object.keys(rest).length > 0;
}

/**
 * Create a ProseMirror document node from JSON
 */
export function createPmDoc(json: PmNodeJSON): PmNode {
  return docsSchema.nodeFromJSON(json);
}

/**
 * Convert an array of blocks to a ProseMirror document node
 * @param blocks - The blocks to convert
 * @param stylePool - Optional style pool for looking up styles by ID
 */
export function blocksToPmDoc(blocks: Block[], stylePool?: TextStylePool): PmNode {
  const content = blocksToPmContent(blocks, stylePool);
  const json = {
    type: 'doc',
    content: content.length > 0 ? content : [{ type: 'paragraph', content: [] }],
  };
  
  return docsSchema.nodeFromJSON(json);
}

/**
 * Convert the sections of a document to one ProseMirror document: the blocks of each section, and after every section
 * but the last a `section_break` that names it. A single section gives the same document as `blocksToPmDoc`.
 * @param sections - The sections, in document order
 * @param stylePool - Optional style pool for looking up styles by ID
 */
export function sectionsToPmDoc(sections: ReadonlyArray<{ id: string; blocks: Block[] }>, stylePool?: TextStylePool): PmNode {
  const content = sections.flatMap((section, index) => {
    const nodes = blocksToPmContent(section.blocks, stylePool);
    return index < sections.length - 1 ? [...nodes, sectionBreakNode(section.id)] : nodes;
  });
  return docsSchema.nodeFromJSON({
    type: 'doc',
    content: content.length > 0 ? content : [{ type: 'paragraph', content: [] }],
  });
}

/** Blocks as ProseMirror node JSON, with consecutive list items grouped into lists (the content of a document or of a table cell). */
function blocksToPmContent(blocks: Block[], stylePool?: TextStylePool): PmNodeJSON[] {
  const rawNodes = blocks
    .map(block => blockToPmNode(block, stylePool))
    .filter((node): node is PmNodeJSON => node !== null);

  // Group consecutive list items into single list nodes
  const content: PmNodeJSON[] = [];
  let currentList: {
    type: string;
    listType: ListType | undefined;
    start?: number;
    items: PmNodeJSON[];
  } | null = null;
  const listNode = (l: NonNullable<typeof currentList>): PmNodeJSON => ({
    type: l.type,
    ...(l.start && l.start !== 1 ? { attrs: { start: l.start } } : {}),
    content: l.items,
  });
  
  for (const node of rawNodes) {
    if (node.type === '__list_item__') {
      // This is a list item that needs to be grouped
      const listType = node.listType === 'bullet' ? 'bullet_list' : 'ordered_list';
      
      // An item that states its own first number begins a new list, even right after another list of the same kind.
      const restarts = node.listType === 'numbered' && node.listStart !== undefined && !!currentList && currentList.items.length > 0;
      if (currentList && currentList.listType === node.listType && !restarts) {
        // Add to existing list
        currentList.items.push(...(node.content ?? []));
      } else {
        // Finish previous list if any
        if (currentList) content.push(listNode(currentList));
        // Start new list
        currentList = {
          type: listType,
          listType: node.listType,
          start: node.listStart,
          items: [...(node.content ?? [])],
        };
      }
    } else {
      // Regular node - flush any pending list
      if (currentList) {
        content.push(listNode(currentList));
        currentList = null;
      }
      content.push(node);
    }
  }
  
  // Flush final list if any
  if (currentList) content.push(listNode(currentList));
  
  return content;
}

/**
 * Get plain text from a ProseMirror node. Text deleted by a tracked change is left out (the final text) unless
 * `includeDeleted` is set.
 */
export function getPmPlainText(node: PmNode, includeDeleted = false): string {
  let text = '';
  node.descendants((child) => {
    if (child.isText && (includeDeleted || !child.marks.some((m) => m.type.name === 'deletion'))) {
      text += child.text;
    }
  });
  return text;
}

/**
 * Check if a ProseMirror document is empty
 */
export function isDocEmpty(doc: PmNode): boolean {
  if (doc.childCount === 0) return true;
  if (doc.childCount === 1) {
    const child = doc.firstChild;
    if (child && child.type.name === 'paragraph' && child.content.size === 0) {
      return true;
    }
  }
  return false;
}

