import { MAX_COL, type RangeBounds } from './cell-ref';

// ---------------------------------------------------------------------------
// Which ranges cover which cells, for all cells and all ranges at once, in O((cells + ranges) log) instead of cells x ranges.
// The cells are put in row order and swept down the sheet; the ranges that reach the current row are held in a small structure
// over the columns (`Cover`), which answers for one column. Merges ask "how many ranges cover this cell" and hyperlinks "which
// range, the last one written, covers it", so a range of a million cells costs no more than a single cell.
// ---------------------------------------------------------------------------

/** Cells as parallel arrays, in row order. */
export interface CellIndex {
  keys: string[];
  rows: number[];
  cols: number[];
}

/** The cells of a `row:col`-keyed map, in row order (ties keep the map's order; a map already in row order costs one pass). */
export function indexCells(keys: Iterable<string>): CellIndex {
  const all: Array<{ key: string; row: number; col: number }> = [];
  let sorted = true;
  for (const key of keys) {
    const i = key.indexOf(':');
    const row = Number(key.slice(0, i));
    if (all.length > 0 && row < all[all.length - 1].row) sorted = false;
    all.push({ key, row, col: Number(key.slice(i + 1)) });
  }
  if (!sorted) all.sort((a, b) => a.row - b.row);
  return { keys: all.map((c) => c.key), rows: all.map((c) => c.row), cols: all.map((c) => c.col) };
}

/** The ranges that reach the row being swept, over the columns. */
export interface Cover<T> {
  add(id: number, range: RangeBounds): void;
  remove(id: number, range: RangeBounds): void;
  /** The answer for one column. */
  at(col: number): T;
}

/** Sweep `ranges` over `cells` and call `visit(cell index, answer of the cover at the cell)` for every cell. */
export function sweepCover<T>(cells: CellIndex, ranges: readonly RangeBounds[], cover: Cover<T>, visit: (cell: number, answer: T) => void): void {
  const byStart = ranges.map((_, i) => i).sort((a, b) => ranges[a].startRow - ranges[b].startRow);
  const byEnd = ranges.map((_, i) => i).sort((a, b) => ranges[a].endRow - ranges[b].endRow);
  const active = new Uint8Array(ranges.length);
  let started = 0;
  let ended = 0;
  for (let k = 0; k < cells.rows.length; k++) {
    const row = cells.rows[k];
    while (started < byStart.length && ranges[byStart[started]].startRow <= row) {
      const id = byStart[started++];
      // a range that ended above this row never enters
      if (ranges[id].endRow >= row) {
        active[id] = 1;
        cover.add(id, ranges[id]);
      }
    }
    while (ended < byEnd.length && ranges[byEnd[ended]].endRow < row) {
      const id = byEnd[ended++];
      if (active[id] === 1) {
        active[id] = 0;
        cover.remove(id, ranges[id]);
      }
    }
    visit(k, cover.at(cells.cols[k]));
  }
}

/**
 * How many ranges cover a column: a Fenwick tree over the columns (a range adds 1 to its columns; the answer for a column is a
 * prefix sum).
 */
export function countCover(): Cover<number> {
  const tree = new Int32Array(MAX_COL + 2);
  const bump = (from: number, delta: number): void => {
    for (let i = from + 1; i < tree.length; i += i & -i) tree[i] += delta;
  };
  const change = (range: RangeBounds, delta: number): void => {
    bump(range.startCol, delta);
    bump(range.endCol + 1, -delta);
  };
  return {
    add: (_id, range) => change(range, 1),
    remove: (_id, range) => change(range, -1),
    at(col) {
      let sum = 0;
      for (let i = col + 1; i > 0; i -= i & -i) sum += tree[i];
      return sum;
    },
  };
}

/**
 * The highest id of the ranges that cover a column (`-1` for none): a segment tree over the columns, each node holding the ids
 * of the ranges that cover the whole node as a max-heap. A range leaves lazily: its id is flagged and dropped when it reaches
 * the top of a heap a query looks at.
 */
export function lastCover(rangeCount: number): Cover<number> {
  const leaves = 1 << Math.ceil(Math.log2(MAX_COL));
  const heaps: Array<number[] | undefined> = new Array<number[] | undefined>(2 * leaves);
  const gone = new Uint8Array(rangeCount);
  const nodesOf = (range: RangeBounds, visit: (node: number) => void): void => {
    for (let l = range.startCol + leaves, r = range.endCol + leaves + 1; l < r; l >>= 1, r >>= 1) {
      if (l & 1) visit(l++);
      if (r & 1) visit(--r);
    }
  };
  const push = (heap: number[], id: number): void => {
    let i = heap.push(id) - 1;
    while (i > 0 && heap[(i - 1) >> 1] < id) {
      heap[i] = heap[(i - 1) >> 1];
      i = (i - 1) >> 1;
    }
    heap[i] = id;
  };
  const pop = (heap: number[]): void => {
    const last = heap.pop() as number;
    if (heap.length === 0) return;
    let i = 0;
    for (;;) {
      let child = 2 * i + 1;
      if (child >= heap.length) break;
      if (child + 1 < heap.length && heap[child + 1] > heap[child]) child++;
      if (heap[child] <= last) break;
      heap[i] = heap[child];
      i = child;
    }
    heap[i] = last;
  };
  return {
    add(id, range) {
      nodesOf(range, (node) => push((heaps[node] ??= []), id));
    },
    remove(id) {
      gone[id] = 1;
    },
    at(col) {
      let best = -1;
      for (let node = col + leaves; node >= 1; node >>= 1) {
        const heap = heaps[node];
        if (heap === undefined) continue;
        while (heap.length > 0 && gone[heap[0]] === 1) pop(heap);
        if (heap.length > 0 && heap[0] > best) best = heap[0];
      }
      return best;
    },
  };
}
