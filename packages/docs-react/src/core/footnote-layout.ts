/**
 * Footnote layout: footnotes at the bottom of the page their reference is on.
 *
 * A footnote's text is drawn by a block of its own (see `pm-to-blocks`), kept out of the flow of the body. The layout
 * engine gives each page a capacity (its content height less a reserve at the bottom); this module finds out which
 * notes each page needs, reserves that much, and lays the body out again until the reserves settle (reserves only
 * grow, so it ends). The notes of the references on a page are then placed in the reserved area under a separator.
 * A note that does not fit continues on the next page (with a longer separator); a page never gives more than
 * `maxFraction` of its content height to notes, so the body always keeps room.
 */

import type { FlowBlock, Run } from './flow-blocks';
import { getLineRangeHeight, hasLineData, type Measure } from './measurer';
import type { DocumentLayout, FootnoteHooks, PageFragment, PageLayout, PageNote, PendingNote } from './true-layout-engine';

/** Where each footnote's text is drawn and which body lines refer to which footnote. */
export interface FootnotePlan {
  /** Footnote keys (`footnote:3`) referenced by each body block (by index in the blocks array), by line, in document order. */
  refs: Map<number, Array<{ line: number; key: string }>>;
  /** The block that draws each footnote's text: its index in the blocks and measures arrays, and its id. */
  noteBlocks: Map<string, { index: number; id: string }>;
  /** Height of the separator above a page's notes, unscaled. Default 12. */
  separatorHeight?: number;
  /** The most of a page's content height the notes of a page may take. The rest continues on the next page. Default 0.5. */
  maxFraction?: number;
}

const DEFAULT_SEPARATOR = 12;
const DEFAULT_MAX_FRACTION = 0.5;

/** Footnote keys of the references in a run list, in order, with the index of the run each is in. */
function noteRuns(runs: readonly Run[]): Array<{ runIndex: number; key: string }> {
  const found: Array<{ runIndex: number; key: string }> = [];
  runs.forEach((run, runIndex) => {
    if (run.kind === 'text' && run.noteRef && run.noteRef.kind === 'footnote') found.push({ runIndex, key: `footnote:${run.noteRef.ordinal}` });
  });
  return found;
}

/** Footnote keys referenced anywhere in a block that is laid out whole (a table with its cells, nested ones too). */
function wholeBlockNotes(block: FlowBlock): string[] {
  if (block.kind === 'table') return block.rows.flatMap((row) => row.cells.flatMap((cell) => cell.blocks.flatMap(wholeBlockNotes)));
  if (block.kind === 'paragraph' || block.kind === 'heading' || block.kind === 'listItem') return noteRuns(block.runs).map((n) => n.key);
  return [];
}

/**
 * Which footnotes each of the first `bodyCount` blocks refers to, by the line the reference is on (a table or other block
 * that is not split is one line). `blocks` and `measures` are the same length.
 */
export function footnoteRefs(blocks: readonly FlowBlock[], measures: readonly Measure[], bodyCount: number): Map<number, Array<{ line: number; key: string }>> {
  const refs = new Map<number, Array<{ line: number; key: string }>>();
  for (let index = 0; index < bodyCount; index++) {
    const block = blocks[index];
    const measure = measures[index];
    if (block.kind === 'paragraph' || block.kind === 'heading' || block.kind === 'listItem') {
      const found = noteRuns(block.runs);
      if (found.length === 0) continue;
      const lineOf = (runIndex: number): number => {
        if (!hasLineData(measure)) return 0;
        const line = measure.lines.findIndex((l) => l.segments?.some((s) => s.runIndex === runIndex));
        return line < 0 ? 0 : line;
      };
      refs.set(index, found.map(({ runIndex, key }) => ({ line: lineOf(runIndex), key })));
    } else {
      const keys = wholeBlockNotes(block);
      if (keys.length > 0) refs.set(index, keys.map((key) => ({ line: 0, key })));
    }
  }
  return refs;
}

/**
 * How the layout engine asks about footnotes. `measures` are the measures of the same blocks as `plan` (the notes' blocks
 * included), `scale` the zoom the layout is made at, and `blockOffset` what to add to the engine's block indexes to get
 * the plan's (a section laid out on its own sees its blocks from 0).
 */
export function createFootnoteHooks(plan: FootnotePlan, measures: readonly Measure[], scale: number, blockOffset = 0): FootnoteHooks {
  const separator = (plan.separatorHeight ?? DEFAULT_SEPARATOR) * scale;
  const maxFraction = plan.maxFraction ?? DEFAULT_MAX_FRACTION;
  return {
    refsIn(blockIndex, fromLine, toLine) {
      const keys: string[] = [];
      for (const ref of plan.refs.get(blockIndex + blockOffset) ?? []) {
        if (ref.line >= fromLine && ref.line < toLine && !keys.includes(ref.key)) keys.push(ref.key);
      }
      return keys;
    },

    place(contentHeight, carry, keys) {
      const queue: PendingNote[] = carry.map((c) => ({ ...c }));
      for (const key of keys) if (plan.noteBlocks.has(key) && !queue.some((q) => q.key === key)) queue.push({ key, fromLine: 0 });
      const items: PageNote[] = [];
      if (queue.length === 0) return { items, height: 0, continued: false, carry: [] };
      const limit = contentHeight * maxFraction;
      let used = separator;
      while (queue.length > 0) {
        const head = queue[0];
        const note = plan.noteBlocks.get(head.key)!;
        const measure = measures[note.index];
        const total = hasLineData(measure) ? measure.lines.length : 1;
        const whole = getLineRangeHeight(measure, head.fromLine, total) * scale;
        if (used + whole <= limit) {
          items.push({ key: head.key, blockIndex: note.index, blockId: note.id, fromLine: head.fromLine, toLine: total, totalLines: total, height: whole });
          used += whole;
          queue.shift();
          continue;
        }
        // Not all of it fits: as many lines as do (at least one when nothing else is on the page), the rest goes on.
        let take = 0;
        let height = 0;
        for (let to = head.fromLine + 1; to < total; to++) {
          const h = getLineRangeHeight(measure, head.fromLine, to) * scale;
          if (used + h > limit) break;
          take = to - head.fromLine;
          height = h;
        }
        if (take === 0 && items.length === 0) {
          take = 1;
          height = getLineRangeHeight(measure, head.fromLine, head.fromLine + 1) * scale;
        }
        if (take > 0 && head.fromLine + take < total) {
          items.push({ key: head.key, blockIndex: note.index, blockId: note.id, fromLine: head.fromLine, toLine: head.fromLine + take, totalLines: total, height });
          used += height;
          head.fromLine += take;
        } else if (take > 0) {
          // A note of one line that is taller than the page may give: it is placed whole, as the only thing on the page.
          items.push({ key: head.key, blockIndex: note.index, blockId: note.id, fromLine: head.fromLine, toLine: total, totalLines: total, height: whole });
          used += whole;
          queue.shift();
        }
        break;
      }
      return { items, height: items.length > 0 ? used : 0, continued: carry.length > 0 && items.length > 0 && items[0].key === carry[0].key, carry: queue };
    },
  };
}

/** The content height of a page (scaled). */
function contentHeightOf(layout: DocumentLayout, page: PageLayout): number {
  const config = page.pageConfig ?? layout.pageConfig;
  return (config.height - config.margins.top - config.margins.bottom) * layout.scale;
}

/**
 * Put the notes of each page of a layout made with footnote hooks into its fragments (below its body, at the bottom of
 * the content area) and mark where the separator goes. The layout is returned as a copy.
 */
export function applyPageNotes(layout: DocumentLayout, plan: FootnotePlan): DocumentLayout {
  const separator = (plan.separatorHeight ?? DEFAULT_SEPARATOR) * layout.scale;
  const pages: PageLayout[] = layout.pages.map((page) => {
    const notes = page.notes;
    if (!notes || notes.items.length === 0) return page;
    const areaTop = contentHeightOf(layout, page) - notes.height;
    let y = areaTop + separator;
    const fragments: PageFragment[] = [...page.fragments];
    for (const item of notes.items) {
      fragments.push({
        blockId: item.blockId,
        blockIndex: item.blockIndex,
        fromLine: item.fromLine,
        toLine: item.toLine,
        x: 0,
        y,
        height: item.height,
        isFirstFragment: item.fromLine === 0,
        // A note cut by the page continues on the next one.
        isLastFragment: item.toLine >= item.totalLines,
      });
      y += item.height;
    }
    return { ...page, fragments, footnotes: { top: areaTop, separatorHeight: separator, continued: notes.continued } };
  });
  return { ...layout, pages };
}
