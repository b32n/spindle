import { coveredBeyondAnchor, intersects, isDisjoint, lastWins, normalized } from './merge-regions';
import { MAX_OVERLAPPING_MERGES, SheetImpl } from './sheet';
import { WorkbookImpl } from './workbook';
import type { Range } from './types';

const R = (startRow: number, startCol: number, endRow: number, endCol: number): Range => ({ startRow, startCol, endRow, endCol });

/** A fixed generator, so a failure is the same every run. */
function rng(seed: number): (n: number) => number {
  let s = seed;
  return (n) => {
    s = (Math.imul(s, 1103515245) + 12345) & 0x7fffffff;
    return (s >> 8) % n;
  };
}

describe('intersects and normalized', () => {
  it('tells overlap from touching and from apart, and puts the corners in order', () => {
    expect(intersects(R(0, 0, 2, 2), R(2, 2, 4, 4))).toBe(true);
    expect(intersects(R(0, 0, 2, 2), R(3, 0, 4, 2))).toBe(false);
    expect(intersects(R(0, 0, 2, 2), R(0, 3, 2, 4))).toBe(false);
    expect(normalized(R(5, 7, 1, 2))).toEqual(R(1, 2, 5, 7));
  });
});

describe('isDisjoint', () => {
  it('is true for nothing, one, and rects that do not meet, side by side or one under the other', () => {
    expect(isDisjoint([])).toBe(true);
    expect(isDisjoint([R(0, 0, 9, 9)])).toBe(true);
    expect(isDisjoint([R(0, 0, 1, 1), R(0, 2, 1, 3), R(2, 0, 3, 3)])).toBe(true);
    // the columns of a rect that has ended are free again, also across word boundaries
    expect(isDisjoint([R(0, 0, 0, 100), R(1, 31, 1, 33), R(1, 0, 1, 30), R(1, 34, 1, 16_383)])).toBe(true);
  });

  it('is false for rects that share a cell, however far apart their rows start', () => {
    expect(isDisjoint([R(0, 0, 5, 5), R(5, 5, 6, 6)])).toBe(false);
    expect(isDisjoint([R(0, 0, 1_000_000, 0), R(900_000, 0, 900_001, 3)])).toBe(false);
    expect(isDisjoint([R(0, 30, 0, 33), R(0, 33, 0, 40)])).toBe(false);
    expect(isDisjoint([R(0, 0, 0, 0), R(0, 0, 0, 0)])).toBe(false);
  });

  it('does not know rects outside the sheet\'s 16,384 columns to be disjoint', () => {
    expect(isDisjoint([R(0, 0, 1, 16_384)])).toBe(false);
    expect(isDisjoint([R(0, -1, 1, 3)])).toBe(false);
  });

  it('agrees with a comparison of every pair on random rects', () => {
    const rnd = rng(1);
    for (let t = 0; t < 600; t++) {
      const rects = Array.from({ length: rnd(8) }, () => {
        const r0 = rnd(20);
        const c0 = rnd(80);
        return R(r0, c0, r0 + rnd(4), c0 + rnd(40));
      });
      let expected = true;
      for (let a = 0; a < rects.length; a++) for (let b = a + 1; b < rects.length; b++) if (intersects(rects[a], rects[b])) expected = false;
      expect(isDisjoint(rects)).toBe(expected);
    }
  });
});

describe('lastWins', () => {
  it('keeps all of rects that do not overlap', () => {
    expect(lastWins([R(0, 0, 1, 1), R(2, 2, 3, 3)], 0, 10)).toEqual({ alive: [true, true], considered: 2 });
  });

  it('drops a rect that a later one overlaps, though that one is dropped itself', () => {
    // 0 overlaps 1, 1 overlaps 2, 0 and 2 apart: only the last survives of 1, 0 is dropped by 1 even though 2 then drops 1
    expect(lastWins([R(0, 0, 2, 2), R(2, 2, 4, 4), R(4, 4, 6, 6)], 0, 10).alive).toEqual([false, false, true]);
  });

  it('takes no more than `limit` rects when they overlap, never fewer than the fixed ones', () => {
    const rects = [R(0, 0, 1, 1), R(0, 0, 1, 1), R(0, 0, 1, 1), R(5, 5, 6, 6)];
    expect(lastWins(rects, 0, 2)).toEqual({ alive: [false, true], considered: 2 });
    expect(lastWins(rects, 3, 2)).toEqual({ alive: [false, false, true], considered: 3 });
  });

  it('is not limited when the rects are disjoint', () => {
    const rects = Array.from({ length: 50 }, (_, i) => R(i * 2, 0, i * 2 + 1, 1));
    expect(lastWins(rects, 0, 2).considered).toBe(50);
  });
});

describe('coveredBeyondAnchor', () => {
  it('finds the cells inside a rect but not its first cell', () => {
    const cells: Array<[number, number]> = [[0, 0], [0, 1], [1, 0], [2, 2], [5, 5], [9, 0]];
    expect(coveredBeyondAnchor(cells, [R(0, 0, 2, 1)]).sort()).toEqual([1, 2]);
  });

  it('counts a cell covered by another rect beyond its own anchor, and two rects anchored on one cell as one', () => {
    const cells: Array<[number, number]> = [[0, 0], [1, 1]];
    // (0,0) is the anchor of the first rect and inside the second, from (-1,-1)... the second starts at (0,0) too
    expect(coveredBeyondAnchor(cells, [R(0, 0, 2, 2), R(0, 0, 1, 1)]).sort()).toEqual([1]);
    expect(coveredBeyondAnchor(cells, [R(0, 0, 2, 2), R(0, 0, 0, 0)]).sort()).toEqual([1]);
    expect(coveredBeyondAnchor([[0, 0]], [R(-1, -1, 1, 1), R(0, 0, 0, 1)])).toEqual([0]);
  });

  it('gives nothing without rects or cells, and ignores rects above and below the cells', () => {
    expect(coveredBeyondAnchor([], [R(0, 0, 1, 1)])).toEqual([]);
    expect(coveredBeyondAnchor([[3, 3]], [])).toEqual([]);
    expect(coveredBeyondAnchor([[5, 1]], [R(0, 0, 1, 1), R(9, 0, 10, 1)])).toEqual([]);
  });

  it('agrees with a test of every cell against every rect on random input', () => {
    const rnd = rng(2);
    for (let t = 0; t < 400; t++) {
      const rects = Array.from({ length: rnd(7) }, () => {
        const r0 = rnd(15);
        const c0 = rnd(15);
        return R(r0, c0, r0 + rnd(6), c0 + rnd(6));
      });
      const cells: Array<[number, number]> = Array.from({ length: rnd(40) }, () => [rnd(20), rnd(20)]);
      const expected = cells.flatMap(([r, c], i) => (rects.some((x) => intersects(x, R(r, c, r, c)) && !(x.startRow === r && x.startCol === c)) ? [i] : []));
      expect(coveredBeyondAnchor(cells, rects).sort((a, b) => a - b)).toEqual(expected);
    }
  });
});

// ---------------------------------------------------------------------------
// The differential: merging many ranges at once is merging them one at a time.
// ---------------------------------------------------------------------------

/** Two sheets with the same cells (about `fill` in a hundred of a grid of `size`), the same merges to start with. */
function twins(rnd: (n: number) => number, size: number, fill: number, before: Range[]): [SheetImpl, SheetImpl] {
  const make = (): SheetImpl => new SheetImpl('s', 'S');
  const [one, bulk] = [make(), make()];
  for (let r = 0; r < size; r++)
    for (let c = 0; c < size; c++) if (rnd(100) < fill) for (const s of [one, bulk]) s.setCell(r, c, { value: r * size + c });
  // (the draw above is shared: both got the same cells)
  for (const b of before) {
    one.mergeCells(b);
    bulk.mergeCells(b);
  }
  return [one, bulk];
}

const dump = (s: SheetImpl, size: number): { merges: Range[]; cells: Array<[number, number, unknown]>; ids: string[] } => ({
  merges: s.getMergedRegions(),
  cells: Array.from({ length: size * size }, (_, i) => [Math.floor(i / size), i % size, s.getCell(Math.floor(i / size), i % size)?.value] as [number, number, unknown]).filter(([, , v]) => v !== undefined),
  // which rows and columns have ids (as a pattern: the ids themselves are random)
  ids: [...Array.from({ length: size }, (_, i) => (s.getRowId(i) ? 'r' : '-')), ...Array.from({ length: size }, (_, i) => (s.getColId(i) ? 'c' : '-'))],
});

describe('SheetImpl.mergeCellsBulk against mergeCells one by one', () => {
  /** A range near a random cell, its corners in any order, often a single cell. */
  const randomRange = (rnd: (n: number) => number, size: number, span: number): Range => {
    const r0 = rnd(size);
    const c0 = rnd(size);
    const near = (v: number): number => Math.max(0, Math.min(size - 1, v + rnd(span * 2 + 1) - span));
    return R(r0, c0, near(r0), near(c0));
  };

  it.each([
    ['small ranges, mostly apart', 14, 1, 12],
    ['ranges that overlap often', 12, 5, 25],
    ['few big ranges over many cells', 10, 9, 4],
  ])('%s: the same regions in the same order, the same cells, the same ids (random)', (_name, size, span, count) => {
    const rnd = rng(size * 31 + span);
    for (let t = 0; t < 250; t++) {
      const before = Array.from({ length: rnd(3) }, () => randomRange(rnd, size, span));
      const [one, bulk] = twins(rnd, size, 20 + rnd(80), before);
      const ranges = Array.from({ length: rnd(count) + 1 }, () => randomRange(rnd, size, span));
      for (const range of ranges) one.mergeCells(range);
      expect(bulk.mergeCellsBulk(ranges)).toBe(0);
      expect(dump(bulk, size)).toEqual(dump(one, size));
    }
  });

  it('is the same for merges that do not overlap each other (the fast path), in a grid of blocks', () => {
    const rnd = rng(77);
    for (let t = 0; t < 100; t++) {
      const blocks: Range[] = [];
      for (let r = 0; r < 12; r += 3) for (let c = 0; c < 12; c += 3) if (rnd(3) > 0) blocks.push(R(r, c, r + rnd(3), c + rnd(3)));
      const [one, bulk] = twins(rnd, 12, 70, []);
      for (const b of blocks) one.mergeCells(b);
      bulk.mergeCellsBulk(blocks);
      expect(dump(bulk, 12)).toEqual(dump(one, 12));
    }
  });

  it('notifies once, however many ranges, and not at all when nothing is merged', () => {
    const sheet = new SheetImpl('s', 'S');
    const changes: boolean[] = [];
    (sheet as unknown as { structureChangeListener: (f: boolean) => void }).structureChangeListener = (f) => changes.push(f);
    expect(sheet.mergeCellsBulk([R(0, 0, 0, 0), R(3, 3, 3, 3)])).toBe(0);
    expect(sheet.mergeCellsBulk([])).toBe(0);
    expect(changes).toEqual([]);
    sheet.mergeCellsBulk([R(0, 0, 1, 1), R(4, 4, 5, 5), R(8, 8, 9, 9)]);
    expect(changes).toHaveLength(1);
  });

  it('drops orphan regions as a merge does, and keeps an existing region the new ones do not touch (its own ids)', () => {
    const sheet = new SheetImpl('s', 'S');
    sheet.mergeCells(R(0, 0, 1, 1));
    sheet.mergeCells(R(10, 10, 11, 11));
    const kept = sheet.config.mergedRegions![1];
    // the first region's rows are removed: it is an orphan
    sheet.deleteRows(0, 2);
    sheet.mergeCellsBulk([R(3, 3, 4, 4)]);
    expect(sheet.getMergedRegions()).toEqual([R(8, 10, 9, 11), R(3, 3, 4, 4)]);
    expect(sheet.config.mergedRegions![0]).toBe(kept);
  });

  it('clears a huge merge of a sparse sheet by one pass over its cells, not over its area', () => {
    const sheet = new SheetImpl('s', 'S');
    for (let i = 0; i < 20; i++) sheet.setCell(i * 1000, i % 5, { value: i });
    // (a cell of ids the sheet no longer maps to a position is not in any merge)
    sheet.cells.set('gone-row:gone-col', { value: 'orphan' });
    sheet.mergeCellsBulk([R(0, 0, 1_048_575, 16_383)]);
    expect(sheet.cells.size).toBe(2);
    expect(sheet.getCell(0, 0)?.value).toBe(0);
    expect(sheet.getMergedRegions()).toEqual([R(0, 0, 1_048_575, 16_383)]);
  });

  it('applies ranges outside the sheet\'s columns by the exact path', () => {
    const sheet = new SheetImpl('s', 'S');
    sheet.setCell(0, 20_000, { value: 1 });
    sheet.setCell(0, 20_001, { value: 2 });
    expect(sheet.mergeCellsBulk([R(0, 20_000, 1, 20_001)])).toBe(0);
    expect(sheet.getCell(0, 20_001)).toBeUndefined();
  });

  it('applies the first MAX_OVERLAPPING_MERGES of overlapping ranges and reports the rest, existing merges counting', () => {
    const sheet = new SheetImpl('s', 'S');
    // one range at the far end overlaps the first, so they are not disjoint; every other is a block of its own
    const ranges = Array.from({ length: MAX_OVERLAPPING_MERGES + 5 }, (_, i) => R(i * 2, 0, i * 2 + 1, 1));
    ranges.push(R(0, 0, 3, 3));
    expect(sheet.mergeCellsBulk(ranges)).toBe(6);
    // the 20,000 applied are the first 20,000: the overlapped one (the first) was never dropped, as the last was not applied
    expect(sheet.getMergedRegions()).toHaveLength(MAX_OVERLAPPING_MERGES);
    expect(sheet.getMergedRegions()[0]).toEqual(R(0, 0, 1, 1));
    // with merges there already, the room is what is left of the limit
    const full = new SheetImpl('s', 'S');
    full.config.mergedRegions = Array.from({ length: MAX_OVERLAPPING_MERGES }, (_, i) => ({ startRowId: full.ensureRowId(i * 2), startColId: full.ensureColId(0), endRowId: full.ensureRowId(i * 2 + 1), endColId: full.ensureColId(1) }));
    expect(full.mergeCellsBulk([R(0, 0, 3, 3), R(100_000, 0, 100_001, 1)])).toBe(2);
    expect(full.getMergedRegions()).toHaveLength(MAX_OVERLAPPING_MERGES);
  });
});

describe('WorkbookImpl.mergeCellsBulk', () => {
  const book = (): WorkbookImpl => {
    const wb = new WorkbookImpl('wb', 'WB');
    wb.setCellValue(undefined, 0, 0, 'a');
    wb.setCellValue(undefined, 0, 1, 'b');
    return wb;
  };

  it('records one history step like mergeCells, which undo takes back', () => {
    const wb = book();
    wb.mergeCellsBulk([R(0, 0, 0, 1), R(3, 3, 4, 4)]);
    expect(wb.getSheet().getMergedRegions()).toHaveLength(2);
    expect(wb.getSheet().getCell(0, 1)).toBeUndefined();
    wb.undo();
    expect(wb.getSheet().getMergedRegions()).toHaveLength(0);
    expect(wb.getSheet().getCell(0, 1)?.value).toBe('b');
  });

  it('records no history when asked not to (a load), or during a batch', () => {
    const wb = book();
    wb.setData(wb.getData());
    wb.mergeCellsBulk([R(0, 0, 0, 1)], undefined, { history: false });
    expect(wb.canUndo()).toBe(false);
    const b = book();
    b.setData(b.getData());
    b.batch(() => b.mergeCellsBulk([R(0, 0, 0, 1)]));
    expect(b.canUndo()).toBe(true);
    b.undo();
    expect(b.canUndo()).toBe(false);
  });

  it('merges in the named sheet, and does nothing for one that is not there', () => {
    const wb = book();
    const second = wb.addSheet('Second').id;
    expect(wb.mergeCellsBulk([R(0, 0, 1, 1)], second)).toBe(0);
    expect(wb.getSheet(second).getMergedRegions()).toHaveLength(1);
    expect(wb.getSheet().getMergedRegions()).toHaveLength(0);
    expect(wb.mergeCellsBulk([R(0, 0, 1, 1)], 'nope')).toBe(0);
  });
});
