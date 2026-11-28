// Shift the references inside a formula's TEXT for an insert/delete of rows or columns.
//
// A cached-only formula (one the engine cannot evaluate) has no stable AST, so nothing re-renders its references when the sheet
// changes shape. This rewrites them in the text instead, by the same rule as the stable-id path of a normal formula (`shiftSpan`: as
// Excel does, a range keeps its surviving rows/columns and is `#REF!` only when all of them are deleted; a deleted cell is `#REF!`),
// whole-column/row references only react to their own axis, and references to other sheets only react to edits of that sheet. 3D
// references and references to sheets that do not exist are positional in the engine, and so are left alone here too.
//
// It runs on the engine tokenizer in lenient mode, so only reference tokens are rewritten: every other byte (whitespace, case, string
// literals, function names, operators, structured references, fragments the grammar rejects) is copied as written. One pass over the
// tokens: linear in the text.

import { columnIndexToLabel } from './cell-reference';
import { tokenizeLenient, type Token } from './tokenizer';
import { MAX_SHEET_COLS, MAX_SHEET_ROWS } from './types';

export interface StructuralEdit {
  axis: 'row' | 'col';
  kind: 'insert' | 'delete';
  /** First row/column (0-based) inserted before, or deleted. */
  start: number;
  count: number;
}

/** Where a 0-based row/column goes: its new position, or `undefined` when the edit deletes it (or pushes it off the sheet). */
export function shiftIndex(index: number, edit: StructuralEdit): number | undefined {
  const limit = edit.axis === 'row' ? MAX_SHEET_ROWS : MAX_SHEET_COLS;
  let shifted = index;
  if (edit.kind === 'insert') {
    if (index >= edit.start) shifted = index + edit.count;
  } else if (index >= edit.start + edit.count) {
    shifted = index - edit.count;
  } else if (index >= edit.start) {
    return undefined;
  }
  return shifted < limit ? shifted : undefined;
}

/**
 * Where the rows/columns `lo..hi` (0-based, inclusive) of a range go, as Excel moves them. The rule both reference paths share (this
 * text shifter and the engine's stable-id adjustment of a delete):
 * - insert: a boundary at or after the insertion point moves, so inserting before the first row shifts the range, inserting after
 *   the last row leaves it, and inserting inside it extends it;
 * - delete: the rows that survive stay, a deleted boundary moves to the nearest survivor (`A2:A5` without row 5 is `A2:A4`, without
 *   row 2 is `A2:A4`, `2:4` without row 2 is `2:3`), and the range is gone (`undefined`, `#REF!`) only when all of its rows are deleted.
 * A boundary pushed off the sheet also gives `undefined`. A single cell is the range `i..i`.
 */
export function shiftSpan(lo: number, hi: number, edit: StructuralEdit): { lo: number; hi: number } | undefined {
  const { start, count } = edit;
  if (edit.kind === 'insert') {
    const span = { lo: lo >= start ? lo + count : lo, hi: hi >= start ? hi + count : hi };
    const limit = edit.axis === 'row' ? MAX_SHEET_ROWS : MAX_SHEET_COLS;
    return span.hi < limit ? span : undefined;
  }
  const end = start + count;
  if (lo >= start && hi < end) return undefined;
  return {
    lo: lo < start ? lo : lo < end ? start : lo - count,
    hi: hi < start ? hi : hi < end ? start - 1 : hi - count,
  };
}

/** `shiftSpan` for the two corners of a range in the order written (either may be the lower one); they come back in that order. */
export function shiftPair(a: number, b: number, edit: StructuralEdit): [number, number] | undefined {
  const span = shiftSpan(Math.min(a, b), Math.max(a, b), edit);
  if (!span) return undefined;
  return a <= b ? [span.lo, span.hi] : [span.hi, span.lo];
}

/** The label for `index`, lower case when the label it replaces is (`a1` stays `a1`-style). */
function labelLike(old: string, index: number): string {
  const label = columnIndexToLabel(index);
  return old === old.toLowerCase() ? label.toLowerCase() : label;
}

/** A corner's text `[$]LETTERS[$]DIGITS` with the edited coordinate replaced (the case of the letters and the `$`s are kept). */
function cellCornerText(text: string, axis: StructuralEdit['axis'], next: number): string {
  const letters = text.startsWith('$') ? 1 : 0;
  let digits = letters;
  while (text[digits] < '0' || text[digits] > '9') digits++; // the tokenizer matched `$?letters$?digits`: a digit follows
  const lettersEnd = text[digits - 1] === '$' ? digits - 1 : digits;
  return axis === 'row'
    ? text.slice(0, digits) + String(next + 1)
    : text.slice(0, letters) + labelLike(text.slice(letters, lettersEnd), next) + text.slice(lettersEnd);
}

/** A whole-column corner `[$]LETTERS` or a whole-row corner `[$]DIGITS` with its coordinate replaced. */
function wholeCornerText(text: string, isColumn: boolean, next: number): string {
  const dollar = text.startsWith('$') ? '$' : '';
  return dollar + (isColumn ? labelLike(text.slice(dollar.length), next) : String(next + 1));
}

interface Corner {
  /** Replace text[from, to) with the corner's new text. */
  from: number;
  to: number;
  /** Its row/column (0-based) on the edited axis. */
  index: number;
  /** Whole-column/row corner (`$A`, `3`) rather than a cell's. */
  whole: boolean;
}

interface Span {
  /** Replace text[from, to) with `text`. */
  from: number;
  to: number;
  text: string;
}

/**
 * `formula` with its references adjusted to `edit`. `ownerEdited`: the formula sits on the edited sheet (its unprefixed references
 * point there); `isEditedSheet(name)`: a sheet prefix names the edited sheet. Returns `formula` itself when nothing changes.
 */
export function shiftFormulaRefs(
  formula: string,
  edit: StructuralEdit,
  ownerEdited: boolean,
  isEditedSheet: (name: string) => boolean,
): string {
  const base = formula.startsWith('=') ? 1 : 0;
  const text = base ? formula.slice(1) : formula;
  const tokens = tokenizeLenient(text); // ends with the `eof` token, so a token has a successor
  const spans: Span[] = [];

  /** Where a token's own text ends: before the whitespace that precedes the next token. */
  const endOf = (k: number): number => {
    let end = tokens[k + 1].pos;
    while (/\s/.test(text[end - 1])) end--;
    return end;
  };
  const isOp = (k: number, op: string): boolean => {
    const t = tokens[k];
    return t?.kind === 'op' && t.text === op;
  };

  /** The corners of the reference at tokens[ref..last] that the edit's axis addresses: their text span and current index. */
  const cornersOf = (ref: number, last: number): Corner[] => {
    const t = tokens[ref] as Extract<Token, { kind: 'cell' | 'colrange' | 'rowrange' }>;
    const index = (c: { row: number; col: number }): number => (edit.axis === 'row' ? c.row : c.col);
    if (t.kind === 'cell') {
      const corner = (k: number): Corner => {
        const c = tokens[k] as Extract<Token, { kind: 'cell' }>;
        return { from: c.pos, to: endOf(k), index: index(c), whole: false };
      };
      return last === ref ? [corner(ref)] : [corner(ref), corner(last)];
    }
    if ((t.kind === 'colrange') !== (edit.axis === 'col')) return []; // a whole column only reacts to columns, a whole row to rows
    const end = endOf(ref);
    const colon = text.indexOf(':', t.pos);
    const [i1, i2] = t.kind === 'colrange' ? [t.c1, t.c2] : [t.r1, t.r2];
    return [
      { from: t.pos, to: colon, index: i1, whole: true },
      { from: colon + 1, to: end, index: i2, whole: true },
    ];
  };

  for (let k = 0; k < tokens.length - 1; k++) {
    const t = tokens[k];
    const prefixed = t.kind === 'sheet';
    const ref = prefixed ? k + 1 : k;
    const refKind = tokens[ref].kind;
    if (refKind !== 'cell' && refKind !== 'colrange' && refKind !== 'rowrange') continue;

    // Which sheet the reference points at, and whether the edit touches it. A 3D span and an external workbook's sheet
    // (`[1]Sheet1!A1`) are not the edited sheet's.
    let applies = ownerEdited;
    if (t.kind === 'sheet') {
      const external = k > 0 && tokens[k - 1].kind === 'raw' && (tokens[k - 1] as { end: number }).end === t.pos;
      applies = !external && t.end === undefined && isEditedSheet(t.start);
    }

    // `A1:B2`: two cells joined by `:` are one range, unless the left one is further along a chain (`X:A1:B2`: an operator node). A
    // right corner may repeat the left one's own sheet prefix; a different one makes two references.
    let last = ref;
    if (refKind === 'cell' && isOp(ref + 1, ':') && !isOp(k - 1, ':')) {
      let r = ref + 2;
      const rightSheet = tokens[r];
      if (rightSheet.kind === 'sheet') {
        r = t.kind === 'sheet' && t.end === undefined && rightSheet.end === undefined && rightSheet.start === t.start ? r + 1 : -1;
      }
      const rightCell = tokens[r];
      if (rightCell?.kind === 'cell' && !(isOp(r + 1, '#') && !tokens[r + 1].space)) last = r;
    }

    if (applies) {
      const corners = cornersOf(ref, last);
      const next = corners.length === 0 ? [] : shiftPair(corners[0].index, corners[corners.length - 1].index, edit);
      if (next === undefined) {
        // A range whose rows/columns are all deleted, or a deleted cell, is `#REF!` whole, sheet prefix and right corner included.
        spans.push({ from: t.pos, to: endOf(last), text: '#REF!' });
      } else {
        corners.forEach((corner, n) => {
          if (next[n] === corner.index) return;
          const own = text.slice(corner.from, corner.to);
          const replaced = corner.whole ? wholeCornerText(own, edit.axis === 'col', next[n]) : cellCornerText(own, edit.axis, next[n]);
          spans.push({ from: corner.from, to: corner.to, text: replaced });
        });
      }
    }
    k = last;
  }

  if (spans.length === 0) return formula;
  let out = base ? '=' : '';
  let at = 0;
  for (const span of spans) {
    out += text.slice(at, span.from) + span.text;
    at = span.to;
  }
  return out + text.slice(at);
}
