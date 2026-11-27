import { countCover, indexCells, lastCover, sweepCover } from './cell-ranges';
import { MAX_COL, type RangeBounds } from './cell-ref';

function rng(seed: number): () => number {
  let s = seed >>> 0;
  return () => {
    s = (s + 0x6d2b79f5) >>> 0;
    let t = Math.imul(s ^ (s >>> 15), 1 | s);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

const range = (startRow: number, startCol: number, endRow: number, endCol: number): RangeBounds => ({ startRow, startCol, endRow, endCol });

describe('indexCells', () => {
  it('keeps the order of cells already in row order', () => {
    expect(indexCells(['0:0', '0:5', '3:1'])).toEqual({ keys: ['0:0', '0:5', '3:1'], rows: [0, 0, 3], cols: [0, 5, 1] });
  });

  it('puts cells in row order, ties in their given order', () => {
    expect(indexCells(['5:1', '0:2', '5:0', '2:2', '0:1'])).toEqual({ keys: ['0:2', '0:1', '2:2', '5:1', '5:0'], rows: [0, 0, 2, 5, 5], cols: [2, 1, 2, 1, 0] });
  });

  it('is empty for no cells', () => {
    expect(indexCells([])).toEqual({ keys: [], rows: [], cols: [] });
  });
});

describe('sweepCover', () => {
  const answers = <T>(keys: string[], ranges: RangeBounds[], cover: Parameters<typeof sweepCover<T>>[2]): Map<string, T> => {
    const cells = indexCells(keys);
    const out = new Map<string, T>();
    sweepCover(cells, ranges, cover, (cell, answer) => out.set(cells.keys[cell], answer));
    return out;
  };

  it('counts the ranges that cover each cell', () => {
    const got = answers(['0:0', '0:3', '1:1', '2:2', '4:4', '9:0'], [range(0, 0, 1, 1), range(1, 1, 2, 2), range(0, 0, 0, 5)], countCover());
    expect([...got]).toEqual([['0:0', 2], ['0:3', 1], ['1:1', 2], ['2:2', 1], ['4:4', 0], ['9:0', 0]]);
  });

  it('never lets a range that ended above the first cell it could reach cover anything', () => {
    const got = answers(['5:0', '6:0'], [range(0, 0, 2, 0), range(0, 0, 5, 0)], countCover());
    expect([...got]).toEqual([['5:0', 1], ['6:0', 0]]);
  });

  it('answers the last range of those covering a cell, or -1', () => {
    const got = answers(['0:0', '1:1', '2:2', '3:3', '0:9'], [range(0, 0, 3, 3), range(1, 1, 1, 1), range(2, 0, 2, 5)], lastCover(3));
    expect([...got]).toEqual([['0:0', 0], ['0:9', -1], ['1:1', 1], ['2:2', 2], ['3:3', 0]]);
  });

  it('reaches the first and the last column and the last row of the sheet', () => {
    const got = answers(['0:0', `1048575:${MAX_COL - 1}`], [range(0, 0, 1_048_575, MAX_COL - 1)], lastCover(1));
    expect([...got.values()]).toEqual([0, 0]);
    expect([...answers(['0:0', `1048575:${MAX_COL - 1}`], [range(0, 0, 1_048_575, MAX_COL - 1)], countCover()).values()]).toEqual([1, 1]);
  });

  it('has nothing to say without ranges or without cells', () => {
    expect([...answers(['0:0'], [], lastCover(0)).values()]).toEqual([-1]);
    expect(answers([], [range(0, 0, 1, 1)], countCover()).size).toBe(0);
  });

  it.each(Array.from({ length: 30 }, (_, i) => i + 1))('agrees with checking every range for every cell, seed %i', (seed) => {
    const next = rng(seed);
    const int = (n: number) => Math.floor(next() * n);
    const rows = 1 + int(40);
    const cols = 1 + int(40);
    const ranges = Array.from({ length: int(60) }, () => {
      const r = int(rows);
      const c = int(cols);
      return range(r, c, r + int(rows), c + int(cols));
    });
    const keys = Array.from({ length: int(150) }, () => `${int(rows + 5)}:${int(cols + 5)}`).filter((k, i, all) => all.indexOf(k) === i);
    const covers = (r: RangeBounds, key: string) => {
      const [row, col] = key.split(':').map(Number);
      return row >= r.startRow && row <= r.endRow && col >= r.startCol && col <= r.endCol;
    };
    const counts = answers(keys, ranges, countCover());
    const lasts = answers(keys, ranges, lastCover(ranges.length));
    for (const key of keys) {
      expect(counts.get(key)).toBe(ranges.filter((r) => covers(r, key)).length);
      expect(lasts.get(key)).toBe(ranges.reduce((last, r, i) => (covers(r, key) ? i : last), -1));
    }
  });
});
