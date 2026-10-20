// Selection and cursor management

import type { CursorPosition, TextSelection, Block, InlineContent } from './types';

export function createCursorPosition(blockId: string, offset: number, runIndex?: number): CursorPosition {
  return { blockId, offset, runIndex };
}

export function createCollapsedSelection(position: CursorPosition): TextSelection {
  return {
    anchor: position,
    focus: position,
    isCollapsed: true,
  };
}

export function createSelection(anchor: CursorPosition, focus: CursorPosition): TextSelection {
  const isCollapsed = 
    anchor.blockId === focus.blockId && 
    anchor.offset === focus.offset &&
    anchor.runIndex === focus.runIndex;
  
  return {
    anchor,
    focus,
    isCollapsed,
  };
}

export function isSelectionCollapsed(selection: TextSelection): boolean {
  return selection.isCollapsed;
}

export function getSelectionStart(selection: TextSelection): CursorPosition {
  // For a proper implementation, we'd need to compare positions based on document order
  // For now, we return anchor as start
  return selection.anchor;
}

export function getSelectionEnd(selection: TextSelection): CursorPosition {
  return selection.focus;
}

/**
 * Get the total text length of a block's content
 */
export function getBlockTextLength(block: Block): number {
  if (block.type === 'horizontal-rule' || block.type === 'page-break' || block.type === 'table') {
    return 0;
  }
  
  if (block.type === 'image') {
    return 1; // Treat image as single character for selection purposes
  }

  const content = block.content as InlineContent[];
  return content.reduce((total, item) => {
    if (item.type === 'text') {
      return total + item.text.length;
    }
    if (item.type === 'link') {
      return total + item.text.length;
    }
    if (item.type === 'image' || item.type === 'dynamicField') {
      return total + 1; // Inline atoms (image, field) count as 1
    }
    return total;
  }, 0);
}

/**
 * Get text content from a block
 */
export function getBlockText(block: Block): string {
  if (block.type === 'horizontal-rule' || block.type === 'page-break' || block.type === 'table' || block.type === 'image') {
    return '';
  }

  const content = block.content as InlineContent[];
  return content.reduce((text, item) => {
    if (item.type === 'text') {
      return text + item.text;
    }
    if (item.type === 'link') {
      return text + item.text;
    }
    return text;
  }, '');
}

/**
 * Find the run index and offset within run for a given block offset
 */
export function findRunAtOffset(
  content: InlineContent[],
  offset: number
): { runIndex: number; offsetInRun: number } | null {
  if (content.length === 0) {
    return null;
  }

  const runLength = (item: InlineContent): number => {
    if (item.type === 'text') return item.text.length;
    if (item.type === 'link') return item.text.length;
    if (item.type === 'image' || item.type === 'dynamicField') return 1;
    return 0;
  };

  // Clamp negative offsets to the start of the first run.
  if (offset <= 0) {
    return { runIndex: 0, offsetInRun: 0 };
  }

  let currentOffset = 0;

  for (let i = 0; i < content.length; i++) {
    const runEnd = currentOffset + runLength(content[i]);

    // Prefer the later run at an exact boundary: a boundary offset (=== runEnd)
    // belongs to the next run so newly typed text inherits the following run's
    // style rather than the previous run's. The last run owns its own end.
    if (offset < runEnd || (offset === runEnd && i === content.length - 1)) {
      return { runIndex: i, offsetInRun: offset - currentOffset };
    }

    currentOffset = runEnd;
  }

  // Offset is past the end of all content — clamp to the end of the last run
  // (offsetInRun never exceeds the run's length).
  const lastIndex = content.length - 1;
  return { runIndex: lastIndex, offsetInRun: runLength(content[lastIndex]) };
}

/**
 * Convert run index and offset within run to block offset
 */
export function runOffsetToBlockOffset(
  content: InlineContent[],
  runIndex: number,
  offsetInRun: number
): number {
  let blockOffset = 0;
  
  for (let i = 0; i < runIndex && i < content.length; i++) {
    const item = content[i];
    if (item.type === 'text') {
      blockOffset += item.text.length;
    } else if (item.type === 'link') {
      blockOffset += item.text.length;
    } else if (item.type === 'image' || item.type === 'dynamicField') {
      blockOffset += 1;
    }
  }
  
  return blockOffset + offsetInRun;
}

