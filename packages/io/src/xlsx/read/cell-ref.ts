// ---------------------------------------------------------------------------
// A1 references: `B7`, `$B$7`, `A1:C3`, and the column letters in between. 0-based on the way out.
// ---------------------------------------------------------------------------

export const MAX_ROW = 1_048_576;
export const MAX_COL = 16_384;

/** 0-based column of `A`..`XFD` (any case), or `undefined` for anything else. */
export function columnIndex(letters: string): number | undefined {
  if (letters.length === 0 || letters.length > 3) return undefined;
  let col = 0;
  for (let i = 0; i < letters.length; i++) {
    const c = letters.charCodeAt(i) & ~32;
    if (c < 65 || c > 90) return undefined;
    col = col * 26 + (c - 64);
  }
  return col <= MAX_COL ? col - 1 : undefined;
}

/** Column letters of a 0-based column index. */
export function columnLetters(index: number): string {
  let out = '';
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) out = String.fromCharCode(65 + ((n - 1) % 26)) + out;
  return out;
}

const CELL = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/;

export interface CellPosition {
  row: number;
  col: number;
}

/** A cell reference as 0-based row and column, or `undefined` when it is not one (or lies outside the sheet limits). */
export function parseCellRef(ref: string): CellPosition | undefined {
  const m = CELL.exec(ref);
  if (m === null) return undefined;
  const col = columnIndex(m[1]);
  const row = Number(m[2]);
  return col === undefined || row < 1 || row > MAX_ROW ? undefined : { row: row - 1, col };
}

export interface RangeBounds {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}

/** Longest text read as a range reference (blanks around it included). */
export const MAX_REF_CHARS = 64;

/** `A1` or `A1:C3` as bounds with the corners in order, or `undefined` for anything else. */
export function parseRangeRef(ref: string): RangeBounds | undefined {
  // the longest range is `$XFD$1048576:$XFD$1048576` (25 characters); a longer text is not split (a hostile attribute is megabytes of colons)
  if (ref.length > MAX_REF_CHARS) return undefined;
  const [a, b, ...rest] = ref.trim().split(':');
  if (rest.length > 0) return undefined;
  const s = parseCellRef(a);
  const e = b === undefined ? s : parseCellRef(b);
  if (s === undefined || e === undefined) return undefined;
  return { startRow: Math.min(s.row, e.row), startCol: Math.min(s.col, e.col), endRow: Math.max(s.row, e.row), endCol: Math.max(s.col, e.col) };
}
