// Array shaping functions, batch B: TRANSPOSE, CHOOSECOLS, CHOOSEROWS, TAKE, DROP, HSTACK, VSTACK, EXPAND,
// TOCOL, TOROW, WRAPROWS and WRAPCOLS.
//
// Each returns a 2D array that the spill engine writes into the cells below and to the right. An empty cell that
// passes through (TRANSPOSE(A1:A3) over a blank A2) becomes 0, as in Excel. Arguments left blank (`TAKE(a,,2)`) arrive as
// `undefined` (see omittable.ts). Large arrays are never spread into calls.
//
// LibreOffice 25.2 has no TAKE, DROP, HSTACK, VSTACK, EXPAND, CHOOSECOLS, CHOOSEROWS, TOCOL, TOROW, WRAPROWS or WRAPCOLS,
// so their conformance rows are Excel-documented and tagged unverified.

import type { EagerFn } from './helpers';
import { checkArity, flatten, isErrorValue, toBoolean, toMatrix, toNum } from './helpers';

/** Cells a result may have (the engine's own bound for SEQUENCE). */
const MAX_CELLS = 1_000_000;

/** A copy of the cells where an empty one reads 0. */
function filled(matrix: unknown[][]): unknown[][] {
  return matrix.map((row) => row.map((v) => (v === null || v === undefined ? 0 : v)));
}

/** The columns count of a matrix (rows may be ragged only when built by hand: take the widest). */
function widthOf(matrix: unknown[][]): number {
  let width = 0;
  for (const row of matrix) if (row.length > width) width = row.length;
  return width;
}

/** An integer argument; blank (`undefined`) is `fallback`. */
function intArg(v: unknown, fallback: number): number {
  return v === undefined ? fallback : Math.trunc(toNum(v));
}

/** The 1-based index list of CHOOSECOLS / CHOOSEROWS, negative counting from the end; each must be 1..size in magnitude. */
function chosenIndexes(args: unknown[], size: number): number[] {
  const out: number[] = [];
  for (const v of flatten(args)) {
    const n = Math.trunc(toNum(v));
    if (n === 0 || Math.abs(n) > size) throw new Error('#VALUE!');
    out.push(n > 0 ? n - 1 : size + n);
  }
  return out;
}

/** Rows (or columns) a TAKE / DROP keeps: `[from, to)` out of `size`, for a count that is negative from the end. */
function window(count: number, size: number, take: boolean): [number, number] {
  const n = Math.min(Math.abs(count), size);
  if (take) return count >= 0 ? [0, n] : [size - n, size];
  return count >= 0 ? [n, size] : [0, size - n];
}

/** TOCOL / TOROW's `ignore` argument: 0 keep everything, 1 skip blanks, 2 skip errors, 3 skip both. */
function ignoring(arg: unknown): (v: unknown) => boolean {
  const mode = intArg(arg, 0);
  if (mode < 0 || mode > 3) throw new Error('#VALUE!');
  return (v) =>
    ((mode & 1) !== 0 && (v === null || v === undefined)) || ((mode & 2) !== 0 && typeof v === 'string' && isErrorValue(v));
}

/** The elements of an array in row order, or column order with `byColumn`, minus those `skip` rejects. */
function scan(matrix: unknown[][], byColumn: boolean, skip: (v: unknown) => boolean): unknown[] {
  const out: unknown[] = [];
  const width = widthOf(matrix);
  const take = (v: unknown): void => {
    if (!skip(v)) out.push(v === null || v === undefined ? 0 : v);
  };
  if (byColumn) {
    for (let c = 0; c < width; c++) for (const row of matrix) take(row[c]);
  } else {
    for (const row of matrix) for (let c = 0; c < width; c++) take(row[c]);
  }
  if (out.length === 0) throw new Error('#CALC!');
  return out;
}

/** An argument that must be a vector (one row or one column): its elements in order. */
function vectorOf(arg: unknown): unknown[] {
  const m = toMatrix(arg);
  if (m.length > 1 && widthOf(m) > 1) throw new Error('#VALUE!');
  return m.length > 1 ? m.map((row) => row[0]) : m[0];
}

/** Lay a vector out in `wrap` per row (or per column), padding the last line. */
function wrapped(arg: unknown, wrapArg: unknown, padArg: unknown, byRow: boolean): unknown[][] {
  const items = filled([vectorOf(arg)])[0];
  const wrap = Math.trunc(toNum(wrapArg));
  if (wrap < 1) throw new Error('#NUM!');
  const pad = padArg === undefined ? '#N/A' : padArg;
  const lines = Math.ceil(items.length / wrap);
  if (lines * wrap > MAX_CELLS) throw new Error('#NUM!');
  const out: unknown[][] = [];
  for (let i = 0; i < lines; i++) {
    const line: unknown[] = [];
    for (let j = 0; j < wrap; j++) line.push(i * wrap + j < items.length ? items[i * wrap + j] : pad);
    out.push(line);
  }
  return byRow ? out : transposed(out);
}

function transposed(matrix: unknown[][]): unknown[][] {
  const width = widthOf(matrix);
  const out: unknown[][] = [];
  for (let c = 0; c < width; c++) out.push(matrix.map((row) => row[c]));
  return out;
}

/** HSTACK / VSTACK: join the arguments side by side (or one above the other), padding the shorter ones with #N/A. */
function stack(args: unknown[], horizontal: boolean): unknown[][] {
  const parts = args.map((a) => filled(toMatrix(a)));
  const size = Math.max(...parts.map((p) => (horizontal ? p.length : widthOf(p))));
  const out: unknown[][] = horizontal ? Array.from({ length: size }, () => []) : [];
  let cells = 0;
  for (const p of parts) {
    const rows = p.length;
    const width = widthOf(p);
    cells += horizontal ? size * width : rows * size;
    if (cells > MAX_CELLS) throw new Error('#NUM!');
    if (horizontal) {
      for (let r = 0; r < size; r++) for (let c = 0; c < width; c++) out[r].push(r < rows && c < p[r].length ? p[r][c] : '#N/A');
    } else {
      for (const row of p) out.push(Array.from({ length: size }, (_, c) => (c < row.length ? row[c] : '#N/A')));
    }
  }
  return out;
}

export const arrayExtFunctions: Record<string, EagerFn> = {
  TRANSPOSE: (args) => {
    checkArity(args, 1);
    return transposed(filled(toMatrix(args[0])));
  },

  CHOOSECOLS: (args) => {
    checkArity(args, 2, 255);
    const m = filled(toMatrix(args[0]));
    const picks = chosenIndexes(args.slice(1), widthOf(m));
    return m.map((row) => picks.map((c) => row[c]));
  },

  CHOOSEROWS: (args) => {
    checkArity(args, 2, 255);
    const m = filled(toMatrix(args[0]));
    return chosenIndexes(args.slice(1), m.length).map((r) => m[r].slice()); // a copy: a row picked twice is two arrays
  },

  TAKE: (args) => {
    checkArity(args, 2, 3);
    const m = filled(toMatrix(args[0]));
    const rows = intArg(args[1], m.length);
    const cols = intArg(args[2], widthOf(m));
    if (args[1] === undefined && args[2] === undefined) throw new Error('#VALUE!');
    if (rows === 0 || cols === 0) throw new Error('#CALC!');
    const [r0, r1] = window(rows, m.length, true);
    const [c0, c1] = window(cols, widthOf(m), true);
    return m.slice(r0, r1).map((row) => row.slice(c0, c1));
  },

  DROP: (args) => {
    checkArity(args, 2, 3);
    const m = filled(toMatrix(args[0]));
    const [r0, r1] = window(intArg(args[1], 0), m.length, false);
    const [c0, c1] = window(intArg(args[2], 0), widthOf(m), false);
    if (r0 >= r1 || c0 >= c1) throw new Error('#CALC!');
    return m.slice(r0, r1).map((row) => row.slice(c0, c1));
  },

  HSTACK: (args) => {
    checkArity(args, 1, 254);
    return stack(args, true);
  },

  VSTACK: (args) => {
    checkArity(args, 1, 254);
    return stack(args, false);
  },

  EXPAND: (args) => {
    checkArity(args, 2, 4);
    const m = filled(toMatrix(args[0]));
    const width = widthOf(m);
    const rows = intArg(args[1], m.length);
    const cols = intArg(args[2], width);
    const pad = args[3] === undefined ? '#N/A' : args[3];
    if (rows < m.length || cols < width) throw new Error('#VALUE!');
    if (rows * cols > MAX_CELLS) throw new Error('#NUM!');
    return Array.from({ length: rows }, (_, r) =>
      Array.from({ length: cols }, (_, c) => (r < m.length && c < m[r].length ? m[r][c] : pad))
    );
  },

  TOCOL: (args) => {
    checkArity(args, 1, 3);
    const items = scan(toMatrix(args[0]), args[2] !== undefined && toBoolean(args[2]), ignoring(args[1]));
    return items.map((v) => [v]);
  },

  TOROW: (args) => {
    checkArity(args, 1, 3);
    return [scan(toMatrix(args[0]), args[2] !== undefined && toBoolean(args[2]), ignoring(args[1]))];
  },

  WRAPROWS: (args) => {
    checkArity(args, 2, 3);
    return wrapped(args[0], args[1], args[2], true);
  },

  WRAPCOLS: (args) => {
    checkArity(args, 2, 3);
    return wrapped(args[0], args[1], args[2], false);
  },
};
