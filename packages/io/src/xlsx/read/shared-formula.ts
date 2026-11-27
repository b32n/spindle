import { MAX_COL, MAX_ROW, columnIndex, columnLetters } from './cell-ref';

// ---------------------------------------------------------------------------
// Shared formulas: the master cell holds the formula text (`<f t="shared" ref="B2:B9" si="0">A2*2</f>`), every other cell of
// the group only `si`, and means the master's formula moved by its distance from the master. Moving rewrites the relative
// parts of each cell reference (`$` pins a part), and leaves everything else alone:
//
//   - text in double quotes (a doubled quote is an escape) and quoted sheet names ('My Sheet'!A1: the name is kept, the
//     reference after it is moved)
//   - function calls and names that merely look like a cell (`LOG10(`, `Tax2024`, `A1B`, `_xlfn.X1`): a reference must not be
//     preceded or followed by a name character, and a call is followed by `(`
//   - bracketed groups (structured references `Table1[[#This Row],[Col1]]`, external workbook indexes `[1]Sheet1!A1`)
//   - error literals (`#REF!`)
//
// Whole-column (`A:A`, `$B:C`) and whole-row (`2:2`) references move too. A reference moved off the sheet becomes `#REF!`,
// as Excel does.
// ---------------------------------------------------------------------------

// letters and digits of any script (`Größe1`, `Налог2`), `_`, `.` and `\`; a surrogate half is a name character too, so a name with
// a letter outside the BMP is not cut in the middle (the formula is read one UTF-16 unit at a time)
const NAME_CHAR = /[\p{L}\p{N}_.\\\ud800-\udfff]/u;

/** `formula` as it reads `rows` down and `cols` right of the cell it was written for. */
export function shiftFormula(formula: string, rows: number, cols: number): string {
  if (rows === 0 && cols === 0) return formula;
  let out = '';
  const n = formula.length;
  let i = 0;
  while (i < n) {
    const ch = formula[i];
    if (ch === '"' || ch === "'") {
      // a quoted run, copied as is ("" or '' inside it is an escaped quote)
      let j = i + 1;
      while (j < n && !(formula[j] === ch && formula[j + 1] !== ch)) j += formula[j] === ch ? 2 : 1;
      out += formula.slice(i, j + 1);
      i = j + 1;
    } else if (ch === '[') {
      let depth = 0;
      let j = i;
      for (; j < n; j++) {
        if (formula[j] === '[') depth++;
        else if (formula[j] === ']' && --depth === 0) break;
      }
      out += formula.slice(i, j + 1);
      i = j + 1;
    } else if (ch === '#') {
      const m = /^#[A-Za-z0-9/]+[!?]?/.exec(formula.slice(i, i + 12));
      const len = m === null ? 1 : m[0].length;
      out += formula.slice(i, i + len);
      i += len;
    } else if (ch === '$' || NAME_CHAR.test(ch)) {
      // one token of name characters (and `$`): decide whether it is a reference
      let j = i;
      while (j < n && (formula[j] === '$' || NAME_CHAR.test(formula[j]))) j++;
      const token = formula.slice(i, j);
      // a token glued to a name character before it (`_xlfn.X1`) is part of a name, as is one that names a sheet (`Q1!A1`)
      const named = (i > 0 && NAME_CHAR.test(formula[i - 1])) || formula[j] === '!';
      out += named ? token : moveToken(token, formula, i, j, rows, cols);
      i = j;
    } else {
      out += ch;
      i++;
    }
  }
  return out;
}

const CELL_TOKEN = /^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})$/;
const COL_TOKEN = /^(\$?)([A-Za-z]{1,3})$/;
const ROW_TOKEN = /^(\$?)(\d{1,7})$/;

/** The move of one name-character token: a cell, or the end of a `A:A` / `1:1` range; anything else is returned unchanged. */
function moveToken(token: string, formula: string, start: number, end: number, rows: number, cols: number): string {
  if (formula[end] === '(') return token;
  const cell = CELL_TOKEN.exec(token);
  if (cell !== null) {
    const col = columnIndex(cell[2]);
    if (col === undefined) return token;
    const row = Number(cell[4]);
    const newCol = cell[1] === '$' ? col : col + cols;
    const newRow = cell[3] === '$' ? row : row + rows;
    if (newCol < 0 || newCol >= MAX_COL || newRow < 1 || newRow > MAX_ROW) return '#REF!';
    return `${cell[1]}${columnLetters(newCol)}${cell[3]}${newRow}`;
  }
  // A:A and 1:1 only count next to a colon (otherwise `A` and `1` are a name and a number)
  const touchesColon = formula[start - 1] === ':' || formula[end] === ':';
  if (!touchesColon) return token;
  const colOnly = COL_TOKEN.exec(token);
  if (colOnly !== null) {
    const col = columnIndex(colOnly[2]);
    if (col === undefined) return token;
    if (colOnly[1] === '$') return token;
    const newCol = col + cols;
    return newCol < 0 || newCol >= MAX_COL ? '#REF!' : columnLetters(newCol);
  }
  const rowOnly = ROW_TOKEN.exec(token);
  if (rowOnly !== null) {
    if (rowOnly[1] === '$') return token;
    const newRow = Number(rowOnly[2]) + rows;
    return newRow < 1 || newRow > MAX_ROW ? '#REF!' : String(newRow);
  }
  return token;
}
