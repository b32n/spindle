import type { Range } from './types';

// ---------------------------------------------------------------------------
// The arithmetic behind merging many ranges at once (`SheetImpl.mergeCellsBulk`), which must give what merging them one at a time gives
// (`mergeCells`: a new merge drops every merge it overlaps, and clears every cell it covers but its first) without that cost, which
// is a pass over all the merges and a snapshot of the order maps for each one. Two facts make it possible:
//   - a merge survives the sequence when no LATER merge overlaps it (a later one that overlaps it either drops it or finds it dropped
//     already, and one that does not overlap it leaves it), whatever the merges in between did;
//   - a cell is cleared when it is in any merge but not its first cell (clearing is never undone), so the cleared set is the union
//     of the merges less their anchors, whatever the order.
// ---------------------------------------------------------------------------

/** Whether two normalized numeric ranges (inclusive bounds) overlap. */
export const intersects = (a: Range, b: Range): boolean => a.startRow <= b.endRow && a.endRow >= b.startRow && a.startCol <= b.endCol && a.endCol >= b.startCol;

/** The range with its corners in order. */
export const normalized = (r: Range): Range => ({
  startRow: Math.min(r.startRow, r.endRow),
  endRow: Math.max(r.startRow, r.endRow),
  startCol: Math.min(r.startCol, r.endCol),
  endCol: Math.max(r.startCol, r.endCol),
});

/** A sheet has 16,384 columns: the columns of a range that fits are a bit set of 512 words. */
const COLUMNS = 16_384;

/** Call `visit(word, mask)` for each word of the bit set that holds a column of `[from, to]`, with the bits of those columns. */
function eachWord(from: number, to: number, visit: (word: number, mask: number) => void): void {
  for (let w = from >>> 5; w <= to >>> 5; w++) {
    const lo = w === from >>> 5 ? from & 31 : 0;
    const hi = w === to >>> 5 ? to & 31 : 31;
    visit(w, ((0xffffffff >>> (31 - hi)) & (0xffffffff << lo)) >>> 0);
  }
}

/**
 * Whether no two of `rects` overlap, in O(n log n + n x 512): the rects are swept down the sheet by their first row, and the columns of
 * those that reach the row are a bit set, so a rect that starts over a taken column is an overlap. A range outside the sheet's
 * columns answers `false` (it is not known to be disjoint; the caller takes the exact path).
 */
export function isDisjoint(rects: readonly Range[]): boolean {
  for (const r of rects) if (r.startCol < 0 || r.endCol >= COLUMNS) return false;
  const byStart = rects.map((_, i) => i).sort((a, b) => rects[a].startRow - rects[b].startRow);
  const byEnd = byStart.slice().sort((a, b) => rects[a].endRow - rects[b].endRow);
  const taken = new Uint32Array(COLUMNS / 32);
  let ended = 0;
  for (const id of byStart) {
    const r = rects[id];
    // a rect that ended above this row (it started before this one) gives its columns back
    while (ended < byEnd.length && rects[byEnd[ended]].endRow < r.startRow) {
      const { startCol, endCol } = rects[byEnd[ended++]];
      eachWord(startCol, endCol, (w, mask) => void (taken[w] &= ~mask));
    }
    let free = true;
    eachWord(r.startCol, r.endCol, (w, mask) => {
      if ((taken[w] & mask) !== 0) free = false;
    });
    if (!free) return false;
    eachWord(r.startCol, r.endCol, (w, mask) => void (taken[w] |= mask));
  }
  return true;
}

export interface LastWins {
  /** For the first `considered` rects: whether each is still there after the sequence. */
  alive: boolean[];
  /** How many of the rects were taken (the first ones): all of them unless the exact path was over `limit`. */
  considered: number;
}

/**
 * Which of `rects`, applied in order, survive (see above). The first `fixed` are merges that are already there (never cut). Merges that
 * do not overlap are decided in near-linear time. Overlapping ones need a comparison per pair: a file with overlapping merges (Excel
 * writes none, and repairs a file that has them) is taken as far as `limit` rects, and the rest are left out (`considered`).
 */
export function lastWins(rects: readonly Range[], fixed: number, limit: number): LastWins {
  if (isDisjoint(rects)) return { alive: rects.map(() => true), considered: rects.length };
  const considered = Math.min(rects.length, Math.max(fixed, limit));
  const alive: boolean[] = new Array(considered);
  for (let k = 0; k < considered; k++) {
    alive[k] = true;
    for (let j = k + 1; j < considered; j++) {
      if (intersects(rects[k], rects[j])) {
        alive[k] = false;
        break;
      }
    }
  }
  return { alive, considered };
}

/**
 * Which of `cells` (row, column) are in some rect of `rects` without being its first cell: a sweep down the rows in
 * O((cells + rects) log), with a count of the rects over each column (a Fenwick tree over the columns the rects name).
 * Returns indexes into `cells`.
 */
export function coveredBeyondAnchor(cells: ReadonlyArray<readonly [row: number, col: number]>, rects: readonly Range[]): number[] {
  const cols = [...new Set(rects.flatMap((r) => [r.startCol, r.endCol + 1]))].sort((a, b) => a - b);
  const slot = new Map(cols.map((c, i) => [c, i + 1]));
  const tree = new Int32Array(cols.length + 2);
  const bump = (at: number, delta: number): void => {
    for (let i = at; i < tree.length; i += i & -i) tree[i] += delta;
  };
  const change = (r: Range, delta: number): void => {
    bump(slot.get(r.startCol) as number, delta);
    bump(slot.get(r.endCol + 1) as number, -delta);
  };
  /** The number of rects over column `col`: the prefix sum up to the last named column not after it. */
  const over = (col: number): number => {
    let lo = 0;
    let hi = cols.length;
    while (lo < hi) {
      const mid = (lo + hi) >>> 1;
      if (cols[mid] <= col) lo = mid + 1;
      else hi = mid;
    }
    let sum = 0;
    for (let i = lo; i > 0; i -= i & -i) sum += tree[i];
    return sum;
  };
  // the rects anchored on a cell: it is covered beyond its own when more rects cover it than start on it
  const anchored = new Map<string, number>();
  for (const r of rects) {
    const key = `${r.startRow}:${r.startCol}`;
    anchored.set(key, (anchored.get(key) ?? 0) + 1);
  }
  const byRow = cells.map((_, i) => i).sort((a, b) => cells[a][0] - cells[b][0]);
  const byStart = rects.map((_, i) => i).sort((a, b) => rects[a].startRow - rects[b].startRow);
  const byEnd = byStart.slice().sort((a, b) => rects[a].endRow - rects[b].endRow);
  const active = new Uint8Array(rects.length);
  let started = 0;
  let ended = 0;
  const out: number[] = [];
  for (const i of byRow) {
    const [row, col] = cells[i];
    while (started < byStart.length && rects[byStart[started]].startRow <= row) {
      const id = byStart[started++];
      // a rect that ended above this row never enters
      if (rects[id].endRow >= row) {
        active[id] = 1;
        change(rects[id], 1);
      }
    }
    while (ended < byEnd.length && rects[byEnd[ended]].endRow < row) {
      const id = byEnd[ended++];
      if (active[id] === 1) {
        active[id] = 0;
        change(rects[id], -1);
      }
    }
    if (over(col) > (anchored.get(`${row}:${col}`) ?? 0)) out.push(i);
  }
  return out;
}
