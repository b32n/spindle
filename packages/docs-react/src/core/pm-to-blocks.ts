/**
 * ProseMirror to FlowBlocks Converter
 * 
 * Converts a ProseMirror document to FlowBlocks for layout and rendering.
 * This is the bridge between the editing layer (ProseMirror) and the
 * presentation layer (layout engine + painter).
 */

import { Node as PmNode, Mark } from 'prosemirror-model';
import { sanitizeHref } from './sanitize';
import {
  FlowBlock,
  Run,
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
  
  return blocks;
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
    
    default:
      // Unknown node type - try to extract text content as paragraph
      if (node.isTextblock) {
        return convertGenericTextBlock(node);
      }
      return null;
  }
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
  let topLevelIndex = 0;
  
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
            // pmStart and pmEnd will be set by createBlockPositionMapWithCells
          });
        }
      });
      
      rows.push({
        id: generateBlockId(),
        cells,
      });
    }
  });
  
  return {
    kind: 'table',
    id: generateBlockId(),
    rows,
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
  
  node.forEach((child) => {
    if (child.isText) {
      const textRun = createTextRunFromNode(child);
      runs.push(textRun);
    } else if (child.type.name === 'hard_break') {
      runs.push({ kind: 'lineBreak' });
    } else if (child.type.name === 'dynamic_field') {
      // Body fields are resolved at paint time only in headers/footers; in the
      // body they render as a visible placeholder so they stay selectable.
      runs.push({ kind: 'text', text: `\u2039${String(child.attrs.fieldType)}\u203A` });
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
  
  return runs;
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
  }
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
  }
  
  return attrs;
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
  
  doc.forEach((node, offset) => {
    const nodeStart = offset;
    const nodeEnd = offset + node.nodeSize;
    
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
  });
  
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

