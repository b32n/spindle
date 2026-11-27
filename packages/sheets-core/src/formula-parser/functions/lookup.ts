// Lookup and reference functions.
//
// These all operate on already-evaluated values. The reference-aware
// functions (ROW, COLUMN, OFFSET, INDIRECT) are not here — they need the
// argument's reference rather than its value, which is a separate engine
// change.

import type { EagerFn } from './helpers';
import { checkArity, flatten, toNum, toText, toBoolean, looseEquals, wildcardEquals } from './helpers';
import { columnIndexToLabel } from '../cell-reference';

/** Normalise a value into a 2D matrix so table functions can index it. */
function toMatrix(v: unknown): unknown[][] {
  if (Array.isArray(v)) {
    if (v.length > 0 && Array.isArray(v[0])) return v as unknown[][];
    return [v as unknown[]];
  }
  return [[v]];
}

/** Three-way comparison for ordered (approximate-match) lookups. */
function compareValues(a: unknown, b: unknown): number {
  if (typeof a === 'number' && typeof b === 'number') {
    return a < b ? -1 : a > b ? 1 : 0;
  }
  const na = Number(a);
  const nb = Number(b);
  if (!isNaN(na) && !isNaN(nb) && a != null && b != null && a !== '' && b !== '') {
    return na < nb ? -1 : na > nb ? 1 : 0;
  }
  const sa = toText(a).toLowerCase();
  const sb = toText(b).toLowerCase();
  return sa < sb ? -1 : sa > sb ? 1 : 0;
}

/** Exact match used by VLOOKUP/HLOOKUP/MATCH — honours wildcards in a text key. */
function exactMatch(lookup: unknown, candidate: unknown): boolean {
  if (typeof lookup === 'string' && /[*?~]/.test(lookup)) {
    return wildcardEquals(toText(candidate), lookup);
  }
  return looseEquals(lookup, candidate);
}

/** Shared index resolver for XLOOKUP / XMATCH. Returns -1 when not found. */
function findIndex(
  lookup: unknown,
  arr: unknown[],
  matchMode: number,
  searchMode: number
): number {
  const order: number[] = [];
  for (let i = 0; i < arr.length; i++) {
    order.push(searchMode === -1 ? arr.length - 1 - i : i);
  }

  for (const i of order) {
    const hit =
      matchMode === 2
        ? wildcardEquals(toText(arr[i]), toText(lookup))
        : looseEquals(lookup, arr[i]);
    if (hit) return i;
  }

  if (matchMode === -1 || matchMode === 1) {
    let best = -1;
    for (let i = 0; i < arr.length; i++) {
      const cmp = compareValues(arr[i], lookup);
      if (matchMode === -1 && cmp < 0) {
        if (best === -1 || compareValues(arr[i], arr[best]) > 0) best = i;
      } else if (matchMode === 1 && cmp > 0) {
        if (best === -1 || compareValues(arr[i], arr[best]) < 0) best = i;
      }
    }
    return best;
  }

  return -1;
}

/** The part of an array INDEX selects, as half-open row and column bounds (0-based). */
export interface IndexWindow {
  rowStart: number;
  rowEnd: number;
  colStart: number;
  colEnd: number;
}

/**
 * Resolve INDEX's `(array, row, column, area)` against an array of `rows` x `cols`: a row or column of 0 (or one left
 * blank, `INDEX(a,,2)`) selects the whole column or row, and a lone index into a table selects the whole row.
 * `values` holds the evaluated arguments, `values[0]` being the array itself.
 */
export function indexWindow(rows: number, cols: number, values: unknown[]): IndexWindow {
  const asIndex = (v: unknown): number => (v === undefined ? 0 : Math.trunc(toNum(v)));
  let r: number;
  let c: number;
  if (values.length >= 3) {
    if (values.length === 4 && values[3] !== undefined && asIndex(values[3]) !== 1) throw new Error('#REF!'); // only one area exists
    r = asIndex(values[1]);
    c = asIndex(values[2]);
  } else if (values.length === 1) {
    if (rows !== 1 && cols !== 1) throw new Error('#REF!');
    r = 1;
    c = 1;
  } else if (rows === 1) {
    // Single-axis form: INDEX(row_vector, n) is the n-th column, INDEX(col_vector, n) the n-th row,
    // and INDEX(table, n) the whole n-th row.
    r = 1;
    c = asIndex(values[1]);
  } else if (cols === 1) {
    r = asIndex(values[1]);
    c = 1;
  } else {
    r = asIndex(values[1]);
    c = 0;
  }
  if (r < 0 || c < 0) throw new Error('#VALUE!');
  if (r > rows || c > cols) throw new Error('#REF!');
  return {
    rowStart: Math.max(r - 1, 0),
    rowEnd: r === 0 ? rows : r,
    colStart: Math.max(c - 1, 0),
    colEnd: c === 0 ? cols : c,
  };
}

export const lookupFunctions: Record<string, EagerFn> = {
  VLOOKUP: (args) => {
    const lookup = args[0];
    const table = toMatrix(args[1]);
    const colIndex = toNum(args[2]);
    const approximate = args[3] === undefined ? true : toBoolean(args[3]);
    if (colIndex < 1 || colIndex > (table[0]?.length ?? 0)) throw new Error('#REF!');

    if (approximate) {
      let matchRow = -1;
      for (let i = 0; i < table.length; i++) {
        if (compareValues(table[i][0], lookup) <= 0) matchRow = i;
        else break;
      }
      if (matchRow < 0) throw new Error('#N/A');
      return table[matchRow][colIndex - 1];
    }

    for (const row of table) {
      if (exactMatch(lookup, row[0])) return row[colIndex - 1];
    }
    throw new Error('#N/A');
  },

  HLOOKUP: (args) => {
    const lookup = args[0];
    const table = toMatrix(args[1]);
    const rowIndex = toNum(args[2]);
    const approximate = args[3] === undefined ? true : toBoolean(args[3]);
    if (rowIndex < 1 || rowIndex > table.length) throw new Error('#REF!');
    const header = table[0] ?? [];

    if (approximate) {
      let matchCol = -1;
      for (let i = 0; i < header.length; i++) {
        if (compareValues(header[i], lookup) <= 0) matchCol = i;
        else break;
      }
      if (matchCol < 0) throw new Error('#N/A');
      return table[rowIndex - 1][matchCol];
    }

    for (let i = 0; i < header.length; i++) {
      if (exactMatch(lookup, header[i])) return table[rowIndex - 1][i];
    }
    throw new Error('#N/A');
  },

  XLOOKUP: (args) => {
    const lookup = args[0];
    const lookupArray = flatten([args[1]]);
    const returnArray = flatten([args[2]]);
    const matchMode = args[4] !== undefined ? toNum(args[4]) : 0;
    const searchMode = args[5] !== undefined ? toNum(args[5]) : 1;
    const idx = findIndex(lookup, lookupArray, matchMode, searchMode);
    if (idx >= 0 && idx < returnArray.length) return returnArray[idx];
    if (args[3] !== undefined) return args[3];
    throw new Error('#N/A');
  },

  MATCH: (args) => {
    const lookup = args[0];
    const array = flatten([args[1]]);
    const matchType = args[2] !== undefined ? toNum(args[2]) : 1;

    if (matchType === 0) {
      for (let i = 0; i < array.length; i++) {
        if (exactMatch(lookup, array[i])) return i + 1;
      }
      throw new Error('#N/A');
    }

    let result = -1;
    for (let i = 0; i < array.length; i++) {
      const cmp = compareValues(array[i], lookup);
      if (matchType === 1 ? cmp <= 0 : cmp >= 0) result = i;
      else break;
    }
    if (result < 0) throw new Error('#N/A');
    return result + 1;
  },

  XMATCH: (args) => {
    const lookup = args[0];
    const array = flatten([args[1]]);
    const matchMode = args[2] !== undefined ? toNum(args[2]) : 0;
    const searchMode = args[3] !== undefined ? toNum(args[3]) : 1;
    const idx = findIndex(lookup, array, matchMode, searchMode);
    if (idx < 0) throw new Error('#N/A');
    return idx + 1;
  },

  INDEX: (args) => {
    checkArity(args, 1, 4);
    const matrix = toMatrix(args[0]);
    const w = indexWindow(matrix.length, matrix[0]?.length ?? 0, args);
    if (w.rowEnd - w.rowStart === 1 && w.colEnd - w.colStart === 1) return matrix[w.rowStart][w.colStart];
    return matrix.slice(w.rowStart, w.rowEnd).map((row) => row.slice(w.colStart, w.colEnd));
  },

  LOOKUP: (args) => {
    const lookup = args[0];
    const lookupVector = flatten([args[1]]);
    const resultVector = args[2] !== undefined ? flatten([args[2]]) : lookupVector;
    let idx = -1;
    for (let i = 0; i < lookupVector.length; i++) {
      if (compareValues(lookupVector[i], lookup) <= 0) idx = i;
      else break;
    }
    if (idx < 0 || idx >= resultVector.length) throw new Error('#N/A');
    return resultVector[idx];
  },

  CHOOSE: (args) => {
    const index = Math.trunc(toNum(args[0])); // Excel truncates a fractional index
    if (index < 1 || index >= args.length) throw new Error('#VALUE!');
    return args[index];
  },

  ROWS: (args) => {
    const v = args[0];
    if (Array.isArray(v)) {
      return v.length > 0 && Array.isArray(v[0]) ? v.length : 1;
    }
    return 1;
  },

  COLUMNS: (args) => {
    const v = args[0];
    if (Array.isArray(v)) {
      if (v.length > 0 && Array.isArray(v[0])) return (v[0] as unknown[]).length;
      return v.length;
    }
    return 1;
  },

  ADDRESS: (args) => {
    const row = toNum(args[0]);
    const col = toNum(args[1]);
    const absNum = args[2] !== undefined ? toNum(args[2]) : 1;
    const a1Style = args[3] === undefined ? true : toBoolean(args[3]);
    const sheet = args[4] !== undefined ? toText(args[4]) : '';
    if (row < 1 || col < 1) throw new Error('#VALUE!');

    let ref: string;
    if (a1Style) {
      const colMark = absNum === 1 || absNum === 3 ? '$' : '';
      const rowMark = absNum === 1 || absNum === 2 ? '$' : '';
      ref = `${colMark}${columnIndexToLabel(col - 1)}${rowMark}${row}`;
    } else {
      const rowPart = absNum === 1 || absNum === 2 ? `R${row}` : `R[${row}]`;
      const colPart = absNum === 1 || absNum === 3 ? `C${col}` : `C[${col}]`;
      ref = rowPart + colPart;
    }
    return sheet ? `${sheet}!${ref}` : ref;
  },

  HYPERLINK: (args) => (args[1] !== undefined ? toText(args[1]) : toText(args[0])),
};
