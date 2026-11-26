// Block utility functions

import type { Block, InlineContent, TextRun } from '../types';

let blockIdCounter = 0;

export function generateBlockId(): string {
  return `block_${Date.now()}_${++blockIdCounter}_${Math.random().toString(36).substring(2, 7)}`;
}

export function generateCellId(): string {
  return `cell_${Date.now()}_${++blockIdCounter}_${Math.random().toString(36).substring(2, 7)}`;
}

export function generateRowId(): string {
  return `row_${Date.now()}_${++blockIdCounter}_${Math.random().toString(36).substring(2, 7)}`;
}

/**
 * Create a text run from a string
 */
export function createTextRun(text: string, styleId?: string): TextRun {
  return {
    type: 'text',
    text,
    styleId,
  };
}

/**
 * Get plain text from inline content array
 */
export function getPlainText(content: InlineContent[]): string {
  return content.reduce((text, item) => {
    if (item.type === 'text') {
      return text + item.text;
    }
    if (item.type === 'link') {
      return text + item.text;
    }
    if (item.type === 'contentControl') {
      return text + getPlainText(item.content);
    }
    return text;
  }, '');
}

/**
 * Check if a block has text content
 */
export function hasTextContent(block: Block): boolean {
  if (block.type === 'paragraph' || block.type === 'heading' || block.type === 'list-item') {
    return block.content.length > 0 && getPlainText(block.content).length > 0;
  }
  return false;
}

/**
 * Check if a block is empty
 */
export function isBlockEmpty(block: Block): boolean {
  if (block.type === 'paragraph' || block.type === 'heading' || block.type === 'list-item') {
    return block.content.length === 0 || getPlainText(block.content) === '';
  }
  if (block.type === 'table') {
    return block.rows.length === 0;
  }
  if (block.type === 'content-control' || block.type === 'toc') {
    return block.blocks.every(isBlockEmpty);
  }
  return false;
}

/**
 * Clone a block with a new ID
 */
export function cloneBlock(block: Block): Block {
  const cloned = JSON.parse(JSON.stringify(block)) as Block;
  cloned.id = generateBlockId();
  return cloned;
}

/**
 * Merge two text content arrays, combining adjacent text runs with the same style
 */
export function mergeInlineContent(content1: InlineContent[], content2: InlineContent[]): InlineContent[] {
  const merged = [...content1, ...content2];
  const result: InlineContent[] = [];
  
  for (const item of merged) {
    const last = result[result.length - 1];
    
    // Merge adjacent text runs with the same style
    if (
      last &&
      last.type === 'text' &&
      item.type === 'text' &&
      last.styleId === item.styleId
    ) {
      last.text += item.text;
    } else {
      result.push({ ...item });
    }
  }
  
  return result;
}

/**
 * Split inline content at a given offset
 */
export function splitInlineContent(
  content: InlineContent[],
  offset: number
): [InlineContent[], InlineContent[]] {
  const before: InlineContent[] = [];
  const after: InlineContent[] = [];
  let currentOffset = 0;
  let splitDone = false;
  
  for (const item of content) {
    if (splitDone) {
      after.push({ ...item });
      continue;
    }
    
    let itemLength = 0;
    if (item.type === 'text') {
      itemLength = item.text.length;
    } else if (item.type === 'link') {
      itemLength = item.text.length;
    } else if (item.type === 'image') {
      itemLength = 1;
    }
    
    if (currentOffset + itemLength <= offset) {
      before.push({ ...item });
      currentOffset += itemLength;
    } else if (currentOffset >= offset) {
      after.push({ ...item });
      splitDone = true;
    } else {
      // Split within this item
      const splitPoint = offset - currentOffset;
      
      if (item.type === 'text') {
        if (splitPoint > 0) {
          before.push({ ...item, text: item.text.substring(0, splitPoint) });
        }
        if (splitPoint < item.text.length) {
          after.push({ ...item, text: item.text.substring(splitPoint) });
        }
      } else if (item.type === 'link') {
        if (splitPoint > 0) {
          before.push({ ...item, text: item.text.substring(0, splitPoint) });
        }
        if (splitPoint < item.text.length) {
          after.push({ ...item, text: item.text.substring(splitPoint) });
        }
      } else {
        // Image - can't split, put in before or after based on offset
        if (splitPoint === 0) {
          after.push({ ...item });
        } else {
          before.push({ ...item });
        }
      }
      
      splitDone = true;
      currentOffset += itemLength;
    }
  }
  
  return [before, after];
}

