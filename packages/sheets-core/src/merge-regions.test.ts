import { coveredBeyondAnchor, intersects, isDisjoint, laterOverlaps, normalized, overlapsAny } from './merge-regions';
import { SheetImpl } from './sheet';
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

/** What `laterOverlaps` means, by a comparison of every pair. */
const naiveLater = (rects: readonly Range[]): boolean[] => rects.map((r, k) => rects.slice(k + 1).some((o) => intersects(r, o)));

/** `count` rects near random cells of a `size` square, `span` the most they reach. */
const randomRects = (rnd: (n: number) => number, count: number, size: number, span: number): Range[] =>
  Array.from({ length: count }, () => {
    const r0 = rnd(size);
    const c0 = rnd(size);
    return R(r0, c0, r0 + rnd(span), c0 + rnd(span));
  });

describe('laterOverlaps', () => {
  it('finds none in rects that do not overlap, and none in fewer than two', () => {
    expect(laterOverlaps([])).toEqual([]);
    expect(laterOverlaps([R(0, 0, 1, 1)])).toEqual([false]);
    expect(laterOverlaps([R(0, 0, 1, 1), R(2, 2, 3, 3)])).toEqual([false, false]);
  });

  it('drops a rect that a later one overlaps, though that one is dropped itself', () => {
    // 0 overlaps 1, 1 overlaps 2, 0 and 2 apart: 0 is dropped by 1 even though 2 then drops 1
    expect(laterOverlaps([R(0, 0, 2, 2), R(2, 2, 4, 4), R(4, 4, 6, 6)])).toEqual([true, true, false]);
    expect(laterOverlaps([R(0, 0, 1, 1), R(0, 0, 1, 1), R(0, 0, 1, 1), R(5, 5, 6, 6)])).toEqual([true, true, false, false]);
  });

  it.each([
    ['rects that overlap at a corner, along an edge, inside one another or across (a plus)', 1, 60, 12, 4],
    ['many small rects with few overlaps', 2, 60, 40, 3],
    ['big rects that overlap a lot', 3, 25, 12, 9],
  ])('agrees with a comparison of every pair, %s, in halves of one rect and of the default size (random)', (_name, seed, count, size, span) => {
    const rnd = rng(seed);
    for (let t = 0; t < 120; t++) {
      const rects = randomRects(rnd, rnd(count) + 2, size, span);
      const expected = naiveLater(rects);
      expect(laterOverlaps(rects, 1)).toEqual(expected);
      expect(laterOverlaps(rects)).toEqual(expected);
    }
  });

  it('agrees with a comparison of every pair for many disjoint rects with a few clusters that overlap (random)', () => {
    const rnd = rng(4);
    for (let t = 0; t < 40; t++) {
      // 2 x 2 blocks one under the other, then rects over some of them, then more blocks
      const rects: Range[] = Array.from({ length: 60 }, (_, i) => R(i * 2, 0, i * 2 + 1, 1));
      for (let k = rnd(5); k >= 0; k--) {
        const r0 = rnd(100);
        const c0 = rnd(3);
        rects.splice(rnd(rects.length), 0, R(r0, c0, r0 + rnd(100) + 5, c0 + rnd(3) + 2));
      }
      expect(laterOverlaps(rects, 4)).toEqual(naiveLater(rects));
    }
  });

  it('compares every pair for corners that are not whole numbers, or are too far apart to be counted', () => {
    const rects = [R(0, 0, 1.5, 1), R(1, 1, 2, 2), R(10, 10, 11, 11)];
    expect(laterOverlaps(rects)).toEqual([true, false, false]);
    const far = [R(0, 0, 2 ** 40, 1), R(2 ** 40, 0, 2 ** 40 + 1, 1), R(5, 5, 6, 6)];
    expect(laterOverlaps(far)).toEqual([true, false, false]);
    expect(overlapsAny([R(0, 0, 1.5, 1), R(5, 5, 6, 6)], [R(1, 1, 2, 2)])).toEqual([true, false]);
  });

  it('compares every pair for more than 4 million rects, which an index in a sort key cannot name', () => {
    // a sparse array: only its two rects are there to compare
    const many = Object.assign([R(0, 0, 1, 1), R(5, 5, 6, 6)], { length: 2 ** 22 });
    expect(overlapsAny([R(1, 1, 2, 2), R(8, 8, 9, 9)], many)).toEqual([true, false]);
  });

  it('counts from the smallest corner, so that rects beyond the sheet (negative or past column 16,383) are decided too', () => {
    expect(laterOverlaps([R(-5, -5, 0, 0), R(0, 0, 3, 3), R(10, 20_000, 11, 20_001), R(10, 20_001, 12, 20_002)])).toEqual([true, false, true, false]);
  });
});

describe('overlapsAny', () => {
  it('is false without queries or rects, and tells the sides apart (left, right, above, below, touching a corner)', () => {
    expect(overlapsAny([], [R(0, 0, 1, 1)])).toEqual([]);
    expect(overlapsAny([R(0, 0, 1, 1)], [])).toEqual([false]);
    const q = R(5, 5, 8, 8);
    for (const [name, other, hit] of [
      ['left', R(5, 0, 8, 4), false],
      ['right', R(5, 9, 8, 12), false],
      ['above', R(0, 5, 4, 8), false],
      ['below', R(9, 5, 12, 8), false],
      ['diagonal', R(9, 9, 12, 12), false],
      ['corner', R(8, 8, 12, 12), true],
      ['plus', R(0, 6, 20, 7), true],
      ['inside', R(6, 6, 7, 7), true],
    ] as Array<[string, Range, boolean]>) expect([name, overlapsAny([q], [other])[0]]).toEqual([name, hit]);
  });

  it('agrees with a comparison of every pair on random rects', () => {
    const rnd = rng(5);
    for (let t = 0; t < 200; t++) {
      const queries = randomRects(rnd, rnd(12), 20, 6);
      const rects = randomRects(rnd, rnd(12), 20, 6);
      expect(overlapsAny(queries, rects)).toEqual(queries.map((q) => rects.some((r) => intersects(q, r))));
    }
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

/** Add a region over `r` without `mergeCells`' dropping the regions it overlaps. */
function seedRegion(s: SheetImpl, r: Range): void {
  s.config.mergedRegions = [
    ...(s.config.mergedRegions ?? []),
    { startRowId: s.ensureRowId(r.startRow), startColId: s.ensureColId(r.startCol), endRowId: s.ensureRowId(r.endRow), endColId: s.ensureColId(r.endCol) },
  ];
}

/** Two sheets with the same cells (about `fill` in a hundred of a grid of `size`), the same merges to start with. */
function twins(rnd: (n: number) => number, size: number, fill: number, before: Range[], seeded: Range[] = []): [SheetImpl, SheetImpl] {
  const make = (): SheetImpl => new SheetImpl('s', 'S');
  const [one, bulk] = [make(), make()];
  for (let r = 0; r < size; r++)
    for (let c = 0; c < size; c++) if (rnd(100) < fill) for (const s of [one, bulk]) s.setCell(r, c, { value: r * size + c });
  // (the draw above is shared: both got the same cells)
  for (const b of before) {
    one.mergeCells(b);
    bulk.mergeCells(b);
  }
  // regions set as they are (as collab concurrency or `setData` can leave them), overlapping each other or not
  for (const s of [one, bulk]) for (const r of seeded) seedRegion(s, r);
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
      bulk.mergeCellsBulk(ranges);
      expect(dump(bulk, size)).toEqual(dump(one, size));
    }
  });

  it.each([
    ['ranges that overlap often', 12, 5, 25],
    ['small ranges, mostly apart', 14, 1, 12],
  ])('%s, over regions that already overlap each other: the same regions, cells and ids (random)', (_name, size, span, count) => {
    const rnd = rng(size * 17 + span);
    for (let t = 0; t < 250; t++) {
      const seeded = Array.from({ length: rnd(5) }, () => randomRange(rnd, size, span + 2)).filter((r) => r.startRow !== r.endRow || r.startCol !== r.endCol);
      const [one, bulk] = twins(rnd, size, 20 + rnd(80), Array.from({ length: rnd(2) }, () => randomRange(rnd, size, span)), seeded);
      const ranges = Array.from({ length: rnd(count) + 1 }, () => randomRange(rnd, size, span));
      for (const range of ranges) one.mergeCells(range);
      bulk.mergeCellsBulk(ranges);
      expect(dump(bulk, size)).toEqual(dump(one, size));
    }
  });

  it('keeps two existing regions that overlap each other when the merge is somewhere else, as mergeCells does', () => {
    const [one, bulk] = twins(rng(9), 12, 50, [], [R(0, 0, 3, 3), R(2, 2, 5, 5)]);
    one.mergeCells(R(9, 9, 10, 10));
    bulk.mergeCellsBulk([R(9, 9, 10, 10)]);
    expect(bulk.getMergedRegions()).toEqual([R(0, 0, 3, 3), R(2, 2, 5, 5), R(9, 9, 10, 10)]);
    expect(dump(bulk, 12)).toEqual(dump(one, 12));
    // a merge over one of them drops that one only
    one.mergeCells(R(4, 4, 6, 6));
    bulk.mergeCellsBulk([R(4, 4, 6, 6)]);
    expect(bulk.getMergedRegions()).toEqual([R(0, 0, 3, 3), R(9, 9, 10, 10), R(4, 4, 6, 6)]);
    expect(dump(bulk, 12)).toEqual(dump(one, 12));
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
    sheet.mergeCellsBulk([R(0, 0, 0, 0), R(3, 3, 3, 3)]);
    sheet.mergeCellsBulk([]);
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
    sheet.mergeCellsBulk([R(0, 20_000, 1, 20_001)]);
    expect(sheet.getCell(0, 20_001)).toBeUndefined();
  });

  it('drops no merge for another far away: one overlapping pair among 30,000 disjoint merges loses only the earlier of the pair', () => {
    const sheet = new SheetImpl('s', 'S');
    const ranges = Array.from({ length: 30_000 }, (_, i) => R(i * 2, 0, i * 2 + 1, 1));
    ranges.push(R(1, 1, 2, 2));
    sheet.mergeCellsBulk(ranges);
    const regions = sheet.getMergedRegions();
    // the first two blocks overlap the last range, which drops both (it touches rows 1 and 2, columns 1 and 2)
    expect(regions).toHaveLength(30_000 - 2 + 1);
    expect(regions[0]).toEqual(R(4, 0, 5, 1));
    expect(regions[regions.length - 1]).toEqual(R(1, 1, 2, 2));
  });

  it('is the same as one by one for many disjoint merges with a few clusters that overlap, and regions that are there already', () => {
    const rnd = rng(123);
    const ranges: Range[] = Array.from({ length: 300 }, (_, i) => R(i, (i % 7) * 40, i + rnd(2), (i % 7) * 40 + rnd(3)));
    // clusters: a few ranges over many of the others, and chains of ranges that each overlap the next
    for (let k = 0; k < 4; k++) ranges.splice(rnd(ranges.length), 0, R(rnd(250), rnd(200), rnd(250) + 20, rnd(200) + 90));
    for (let k = 0; k < 20; k++) ranges.push(R(310 + k, 0, 311 + k, 1));
    const [one, bulk] = twins(rnd, 8, 50, [], []);
    const seeded = [R(100, 0, 130, 30), R(120, 20, 150, 60)];
    for (const s of [one, bulk]) for (const r of seeded) seedRegion(s, r);
    for (const range of ranges) one.mergeCells(range);
    bulk.mergeCellsBulk(ranges);
    expect(bulk.getMergedRegions()).toEqual(one.getMergedRegions());
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
    wb.mergeCellsBulk([R(0, 0, 1, 1)], second);
    expect(wb.getSheet(second).getMergedRegions()).toHaveLength(1);
    expect(wb.getSheet().getMergedRegions()).toHaveLength(0);
    wb.mergeCellsBulk([R(0, 0, 1, 1)], 'nope');
  });
});
