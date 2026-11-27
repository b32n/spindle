// Array functions, batch B: TRANSPOSE.
//
// Returns a 2D array that the spill engine writes into the cells below and to the right. An empty cell that passes
// through (TRANSPOSE(A1:A3) over a blank A2) becomes 0, as in Excel.

import type { EagerFn } from './helpers';
import { checkArity, toMatrix } from './helpers';

/** A copy of the cells where an empty one reads 0. */
function filled(matrix: unknown[][]): unknown[][] {
  return matrix.map((row) => row.map((v) => (v === null || v === undefined ? 0 : v)));
}

/** Swap rows and columns (rows may be ragged only when built by hand: the widest sets the width). */
function transposed(matrix: unknown[][]): unknown[][] {
  let width = 0;
  for (const row of matrix) if (row.length > width) width = row.length;
  const out: unknown[][] = [];
  for (let c = 0; c < width; c++) out.push(matrix.map((row) => row[c]));
  return out;
}

export const arrayExtFunctions: Record<string, EagerFn> = {
  TRANSPOSE: (args) => {
    checkArity(args, 1);
    return transposed(filled(toMatrix(args[0])));
  },
};
