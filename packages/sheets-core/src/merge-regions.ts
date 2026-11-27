import type { Range } from './types';

// ---------------------------------------------------------------------------
// The arithmetic behind merging many ranges at once (`SheetImpl.mergeCellsBulk`), which must give what merging them one at a time gives
// (`mergeCells`: a new merge drops every merge it overlaps, and clears every cell it covers but its first) without that cost, which
// is a pass over all the merges and a snapshot of the order maps for each one. Three facts make it possible:
//   - a region that is there already survives when no merge of the sequence overlaps it (regions already there that overlap each
//     other, as collab concurrency can leave them, do not drop one another: only a merge drops);
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

/** The corners of many rects, apart: rows and columns as 32-bit integers counted from the smallest, so that they can be sorted as numbers. */
interface Corners {
  sr: Int32Array;
  er: Int32Array;
  sc: Int32Array;
  ec: Int32Array;
}

/** Larger than any corner of a `Corners` (a coordinate is below 2^30), so that `FLIP - v` turns "greater than" into "less than". */
const FLIP = 2 ** 30 - 1;
/** An index is packed into the low 22 bits of a sort key, and a coordinate over it, so that sorting the keys as numbers sorts the indexes. */
const PACK = 2 ** 22;

/** The corners of `rects` counted from (`row0`, `col0`), or `undefined` when they are not whole numbers that fit. */
function cornersOf(rects: readonly Range[], row0: number, col0: number): Corners | undefined {
  const n = rects.length;
  const out = { sr: new Int32Array(n), er: new Int32Array(n), sc: new Int32Array(n), ec: new Int32Array(n) };
  for (let i = 0; i < n; i++) {
    const { startRow, endRow, startCol, endCol } = rects[i];
    const sr = startRow - row0;
    const er = endRow - row0;
    const sc = startCol - col0;
    const ec = endCol - col0;
    if (!Number.isInteger(sr) || !Number.isInteger(er) || !Number.isInteger(sc) || !Number.isInteger(ec)) return undefined;
    if (Math.max(er, ec) > FLIP || Math.min(sr, sc) < 0) return undefined;
    out.sr[i] = sr;
    out.er[i] = er;
    out.sc[i] = sc;
    out.ec[i] = ec;
  }
  return out;
}

/** The corners of each of the lists, on one scale (`undefined` for all when one does not fit, or when one has 4 million rects). */
function cornersOfAll(...lists: Array<readonly Range[]>): Corners[] | undefined {
  // (an index must fit the low bits of a sort key)
  if (lists.some((list) => list.length >= PACK)) return undefined;
  let row0 = Infinity;
  let col0 = Infinity;
  for (const list of lists) {
    for (const r of list) {
      row0 = Math.min(row0, r.startRow, r.endRow);
      col0 = Math.min(col0, r.startCol, r.endCol);
    }
  }
  const out: Corners[] = [];
  for (const list of lists) {
    const corners = cornersOf(list, row0, col0);
    if (corners === undefined) return undefined;
    out.push(corners);
  }
  return out;
}

/** How many of the sorted `values` are below `v`. */
function below(values: Int32Array, v: number): number {
  let lo = 0;
  let hi = values.length;
  while (lo < hi) {
    const mid = (lo + hi) >>> 1;
    if (values[mid] < v) lo = mid + 1;
    else hi = mid;
  }
  return lo;
}

/** The indexes of `values` in the order of the values (as keys: a value times `PACK` plus its index). */
function sortedKeys(values: Int32Array): Float64Array {
  const keys = new Float64Array(values.length);
  for (let i = 0; i < keys.length; i++) keys[i] = values[i] * PACK + i;
  return keys.sort();
}

/** For each query: how many of the rects have `ra < qa` and `rb < qb`, by a sweep over `a` with a Fenwick tree over the ranks of `b`. */
function dominance(qa: Int32Array, qb: Int32Array, ra: Int32Array, rb: Int32Array): Int32Array {
  const out = new Int32Array(qa.length);
  const m = ra.length;
  const byR = sortedKeys(ra);
  const byQ = sortedKeys(qa);
  const ranks = Int32Array.from(rb).sort();
  const tree = new Int32Array(m + 1);
  let added = 0;
  for (const key of byQ) {
    const q = key % PACK;
    while (added < m && ra[byR[added] % PACK] < qa[q]) {
      for (let i = below(ranks, rb[byR[added] % PACK]) + 1; i <= m; i += i & -i) tree[i]++;
      added++;
    }
    let count = 0;
    for (let i = below(ranks, qb[q]); i > 0; i -= i & -i) count += tree[i];
    out[q] = count;
  }
  return out;
}

const flip = (a: Int32Array): Int32Array => a.map((v) => FLIP - v);

/**
 * For each query: how many of `rects` it overlaps. A rect misses a query by lying left of it, right of it, above it or below it;
 * it cannot be left and right (or above and below) of it, so the misses are counted by inclusion-exclusion over the four sides and
 * the four corner pairs (left and above...), which are counts of "less than in both" (`dominance`). O((q + r) log).
 */
function overlapCounts(queries: Corners, rects: Corners): Int32Array {
  const m = rects.sr.length;
  const counts = new Int32Array(queries.sr.length).fill(m);
  const ec = Int32Array.from(rects.ec).sort();
  const sc = Int32Array.from(rects.sc).sort();
  const er = Int32Array.from(rects.er).sort();
  const sr = Int32Array.from(rects.sr).sort();
  for (let q = 0; q < counts.length; q++) {
    counts[q] -= below(ec, queries.sc[q]) + (m - below(sc, queries.ec[q] + 1)) + below(er, queries.sr[q]) + (m - below(sr, queries.er[q] + 1));
  }
  const corner = (qa: Int32Array, qb: Int32Array, ra: Int32Array, rb: Int32Array): void => {
    const pairs = dominance(qa, qb, ra, rb);
    for (let q = 0; q < counts.length; q++) counts[q] += pairs[q];
  };
  corner(queries.sc, queries.sr, rects.ec, rects.er);
  corner(queries.sc, flip(queries.er), rects.ec, flip(rects.sr));
  corner(flip(queries.ec), queries.sr, flip(rects.sc), rects.er);
  corner(flip(queries.ec), flip(queries.er), flip(rects.sc), flip(rects.sr));
  return counts;
}

/** The rects of `c` at the indexes `at`, and those from index `from` up to `to` (a view). */
const pick = (c: Corners, at: readonly number[]): Corners => ({ sr: Int32Array.from(at, (i) => c.sr[i]), er: Int32Array.from(at, (i) => c.er[i]), sc: Int32Array.from(at, (i) => c.sc[i]), ec: Int32Array.from(at, (i) => c.ec[i]) });
const slice = (c: Corners, from: number, to: number): Corners => ({ sr: c.sr.subarray(from, to), er: c.er.subarray(from, to), sc: c.sc.subarray(from, to), ec: c.ec.subarray(from, to) });

/**
 * For each of `queries`: whether it overlaps any of `rects`. O((q + r) log), but a comparison of every pair (O(q x r)) when a coordinate is
 * not a whole number, when coordinates are 2^30 or more apart, or for 4M rects or more (`cornersOfAll` gives up, as for `laterOverlaps`).
 * The importer cannot reach those: its ranges are whole numbers within a sheet's limits, and at most 1,048,576 merges are read.
 */
export function overlapsAny(queries: readonly Range[], rects: readonly Range[]): boolean[] {
  if (queries.length === 0 || rects.length === 0) return queries.map(() => false);
  const corners = cornersOfAll(queries, rects);
  if (corners === undefined) return queries.map((q) => rects.some((r) => intersects(q, r)));
  return Array.from(overlapCounts(corners[0], corners[1]), (n) => n > 0);
}

/**
 * For each of `rects`: whether a LATER one overlaps it, which is when it does not survive merging them in order (a merge overlapped by a
 * later one is dropped by it or found dropped already). Rects that overlap nothing are found in O(n log n + n x 512) (`isDisjoint`, or
 * a count of the overlaps of each, `overlapCounts`, when some do overlap), so that only the rects in an overlap are decided in order:
 * by splitting the order in half, deciding each half, and asking of the first half's rects that are not decided yet whether the second
 * half's overlap them. O(n log^2 n) at worst; whole numbers within 2^30 of each other are needed for that (anything else compares
 * every pair). Halves of `leaf` rects or fewer are compared pair by pair.
 */
export function laterOverlaps(rects: readonly Range[], leaf = 32): boolean[] {
  const hit = new Array<boolean>(rects.length).fill(false);
  if (rects.length < 2 || isDisjoint(rects)) return hit;
  const all = cornersOfAll(rects)?.[0];
  if (all === undefined) {
    for (let k = 0; k < rects.length; k++) hit[k] = rects.slice(k + 1).some((r) => intersects(rects[k], r));
    return hit;
  }
  // the rects in an overlap (each counts itself)
  const counts = overlapCounts(all, all);
  const involved: number[] = [];
  counts.forEach((n, i) => {
    if (n > 1) involved.push(i);
  });
  const c = pick(all, involved);
  const decided = new Array<boolean>(involved.length).fill(false);
  const solve = (lo: number, hi: number): void => {
    if (hi - lo <= leaf) {
      for (let k = lo; k < hi; k++) {
        for (let j = k + 1; j < hi && !decided[k]; j++) {
          decided[k] = c.sr[k] <= c.er[j] && c.er[k] >= c.sr[j] && c.sc[k] <= c.ec[j] && c.ec[k] >= c.sc[j];
        }
      }
      return;
    }
    const mid = (lo + hi) >>> 1;
    solve(lo, mid);
    solve(mid, hi);
    // (the last of the first half is never decided yet: nothing in its half is after it)
    const open: number[] = [];
    for (let k = lo; k < mid; k++) if (!decided[k]) open.push(k);
    const found = overlapCounts(pick(c, open), slice(c, mid, hi));
    found.forEach((n, i) => {
      if (n > 0) decided[open[i]] = true;
    });
  };
  solve(0, involved.length);
  involved.forEach((at, i) => {
    hit[at] = decided[i];
  });
  return hit;
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
