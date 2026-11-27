// Shift the references inside a formula's TEXT for an insert/delete of rows or columns.
//
// A cached-only formula (one the engine cannot evaluate) has no stable AST, so nothing re-renders its references when the sheet
// changes shape. This rewrites them in the text instead, with the same outcome as the stable-id path of a normal formula: a
// reference moves with its row/column, a reference into deleted rows/columns becomes `#REF!` (a range with one deleted corner
// becomes `#REF!` as a whole, as the engine does for its own formulas), whole-column/row references only react to their own axis,
// and references to other sheets only react to edits of that sheet. 3D references and references to sheets that do not exist are
// positional in the engine, and so are left alone here too.
//
// It runs on the engine tokenizer in lenient mode, so only reference tokens are rewritten: every other byte (whitespace, case, string
// literals, function names, operators, structured references, fragments the grammar rejects) is copied as written. One pass over the
// tokens: linear in the text.

import { columnIndexToLabel } from './cell-reference';
import { tokenizeLenient } from './tokenizer';
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

/** The label for `index`, lower case when the label it replaces is (`a1` stays `a1`-style). */
function labelLike(old: string, index: number): string {
  const label = columnIndexToLabel(index);
  return old === old.toLowerCase() ? label.toLowerCase() : label;
}

/** A corner's text `[$]LETTERS[$]DIGITS` with the edited coordinate replaced (the case of the letters and the `$`s are kept). */
function shiftCellText(text: string, edit: StructuralEdit, row: number, col: number): string | undefined {
  const letters = text.startsWith('$') ? 1 : 0;
  let digits = letters;
  while (text[digits] < '0' || text[digits] > '9') digits++; // the tokenizer matched `$?letters$?digits`: a digit follows
  const lettersEnd = text[digits - 1] === '$' ? digits - 1 : digits;
  const index = edit.axis === 'row' ? row : col;
  const next = shiftIndex(index, edit);
  if (next === undefined) return undefined;
  if (next === index) return text;
  return edit.axis === 'row'
    ? text.slice(0, digits) + String(next + 1)
    : text.slice(0, letters) + labelLike(text.slice(letters, lettersEnd), next) + text.slice(lettersEnd);
}

/** A whole-column corner `[$]LETTERS` or a whole-row corner `[$]DIGITS`; it only reacts to an edit of its own axis. */
function shiftWholeText(text: string, edit: StructuralEdit, index: number, isColumn: boolean): string | undefined {
  if (isColumn !== (edit.axis === 'col')) return text;
  const next = shiftIndex(index, edit);
  if (next === undefined) return undefined;
  if (next === index) return text;
  const dollar = text.startsWith('$') ? '$' : '';
  return dollar + (isColumn ? labelLike(text.slice(dollar.length), next) : String(next + 1));
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

  /** The text a reference token (`A1`, `$A:$C`, `1:3`) becomes, undefined when the edit deletes what it points at. */
  const shifted = (k: number): string | undefined => {
    const t = tokens[k];
    const own = text.slice(t.pos, endOf(k));
    if (t.kind === 'cell') return shiftCellText(own, edit, t.row, t.col);
    if (t.kind === 'colrange' || t.kind === 'rowrange') {
      const isColumn = t.kind === 'colrange';
      const colon = own.indexOf(':');
      const first = shiftWholeText(own.slice(0, colon), edit, t.kind === 'colrange' ? t.c1 : t.r1, isColumn);
      const second = shiftWholeText(own.slice(colon + 1), edit, t.kind === 'colrange' ? t.c2 : t.r2, isColumn);
      return first === undefined || second === undefined ? undefined : `${first}:${second}`;
    }
    return undefined;
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
      const left = shifted(ref);
      const right = last === ref ? left : shifted(last);
      const to = endOf(last);
      if (left === undefined || right === undefined) {
        // A deleted reference is `#REF!` whole, sheet prefix and right corner included.
        spans.push({ from: t.pos, to, text: '#REF!' });
      } else {
        const leftTo = endOf(ref);
        if (left !== text.slice(tokens[ref].pos, leftTo)) spans.push({ from: tokens[ref].pos, to: leftTo, text: left });
        if (last !== ref && right !== text.slice(tokens[last].pos, to)) spans.push({ from: tokens[last].pos, to, text: right });
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
