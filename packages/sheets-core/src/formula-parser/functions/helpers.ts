// Shared helpers for built-in formula functions.

import type { EvaluationContext, ParsedFormulaNode } from '../types';

/** An eager function receives already-evaluated argument values. */
export type EagerFn = (args: unknown[], ctx: EvaluationContext) => unknown;

/**
 * A lazy function receives thunks instead of values, so it can decide which
 * arguments to evaluate (and catch errors thrown while evaluating them).
 * Required by short-circuiting functions like IF, IFS and IFERROR.
 */
export type LazyFn = (argThunks: Array<() => unknown>, ctx: EvaluationContext) => unknown;

/**
 * What a reference function receives: the raw argument AST (so it can read an
 * argument's *reference* rather than its value), the current cell position,
 * and an `evaluate` callback for arguments that are plain values.
 */
export interface RefFnContext {
  args: ParsedFormulaNode[];
  ctx: EvaluationContext;
  currentRow: number;
  currentCol: number;
  evaluate: (node: ParsedFormulaNode) => unknown;
}

/**
 * A reference function works with an argument's reference — needed by
 * ROW/COLUMN/OFFSET/INDIRECT/ISREF/ISFORMULA/CELL.
 */
export type RefFn = (rc: RefFnContext) => unknown;

// Includes the engine-internal #CIRCULAR! so it propagates through aggregations
// and is recognized by ISERROR/IFERROR like any other error token.
const ERROR_PATTERN = /^#(DIV\/0!|N\/A|NAME\?|NULL!|NUM!|REF!|VALUE!|ERROR!|SPILL!|CALC!|CIRCULAR!)$/;

/** True when a value is one of the standard spreadsheet error strings. */
export function isErrorValue(v: unknown): boolean {
  return typeof v === 'string' && ERROR_PATTERN.test(v);
}

/** Recursively flatten range values (`unknown[][]`) and nested args into a flat list. */
export function flatten(args: unknown[]): unknown[] {
  const out: unknown[] = [];
  for (const a of args) {
    if (Array.isArray(a)) out.push(...flatten(a as unknown[]));
    else out.push(a);
  }
  return out;
}

/**
 * Coerce a single value to a number, the way a math function expects.
 * Blank counts as 0; booleans as 1/0; an error value re-throws that error.
 */
export function toNum(v: unknown): number {
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (v == null || v === '') return 0;
  if (typeof v === 'string') {
    if (isErrorValue(v)) throw new Error(v);
    const t = v.trim();
    if (t === '') return 0;
    const n = Number(t);
    if (!isNaN(n)) return n;
  }
  throw new Error('#VALUE!');
}

/**
 * Coerce for SUM-style aggregation over a range: an error value propagates
 * (throws), while text and blanks count as 0 (Excel SUMPRODUCT/SUMIF semantics).
 */
export function toSumNum(v: unknown): number {
  if (typeof v === 'string' && isErrorValue(v)) throw new Error(v);
  const n = Number(v);
  return isNaN(n) ? 0 : n;
}

/** Coerce a value to text. Booleans render as TRUE/FALSE; blank as ''. */
export function toText(v: unknown): string {
  if (v == null) return '';
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'string' && isErrorValue(v)) throw new Error(v);
  return String(v);
}

/** Coerce a value to a boolean using spreadsheet rules. */
export function toBoolean(v: unknown): boolean {
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  if (v == null || v === '') return false;
  if (typeof v === 'string') {
    if (isErrorValue(v)) throw new Error(v); // AND/OR/… propagate an error input
    const s = v.trim().toUpperCase();
    if (s === 'TRUE') return true;
    if (s === 'FALSE') return false;
    const n = Number(v);
    if (!isNaN(n) && v.trim() !== '') return n !== 0;
  }
  return Boolean(v);
}

/**
 * Pull the genuinely-numeric values out of a set of args, the way COUNT-style
 * aggregations do: actual numbers and numeric strings count; blanks, text and
 * booleans are ignored.
 */
export function strictNumbers(args: unknown[]): number[] {
  const out: number[] = [];
  for (const v of flatten(args)) {
    if (typeof v === 'number') {
      if (!isNaN(v)) out.push(v);
    } else if (typeof v === 'string') {
      // An error value anywhere in the input propagates (Excel: SUM/AVERAGE/
      // MAX/MIN/PRODUCT/MEDIAN/… over a range containing #DIV/0! return the
      // error, not a number computed from the rest).
      if (isErrorValue(v)) throw new Error(v);
      const t = v.trim();
      if (t !== '' && !isNaN(Number(t))) out.push(Number(t));
    }
  }
  return out;
}

/** Excel-style equality used by SWITCH and lookup matching. */
export function looseEquals(a: unknown, b: unknown): boolean {
  if (a === b) return true;
  if (typeof a === 'string' && typeof b === 'string') {
    return a.toLowerCase() === b.toLowerCase();
  }
  const na = Number(a);
  const nb = Number(b);
  return !isNaN(na) && !isNaN(nb) && a !== '' && b !== '' && a != null && b != null && na === nb;
}

// Spreadsheet wildcard matching (`*` = any run, `?` = any one char, `~` escapes
// the next). Implemented as an iterative two-pointer glob match rather than a
// translated regex: `*`→`.*` in a regex backtracks catastrophically (ReDoS) on a
// `*`-heavy pattern against a long non-match, and this path is reachable from
// COUNTIF/SUMIF/MATCH/VLOOKUP/SEARCH with attacker-controlled criteria in a
// shared workbook. The two-pointer scan is O(text·pattern), never exponential.

type GlobToken = { star: true } | { any: true } | { lit: string };

/** Parse a wildcard pattern into tokens, honoring `~` escapes. */
function parseGlob(pattern: string): GlobToken[] {
  const toks: GlobToken[] = [];
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern[i];
    if (ch === '~' && i + 1 < pattern.length) {
      toks.push({ lit: pattern[i + 1] });
      i++;
    } else if (ch === '*') {
      toks.push({ star: true });
    } else if (ch === '?') {
      toks.push({ any: true });
    } else {
      toks.push({ lit: ch });
    }
  }
  return toks;
}

/**
 * Match `toks` against `text` starting at `t0`. With `anchorEnd`, the whole of
 * `text` must be consumed (full match); otherwise a prefix match suffices (used
 * for substring SEARCH). Iterative with single-star backtracking — linear, no
 * catastrophic backtracking. `text` and literal tokens must already be lowered
 * if a case-insensitive compare is wanted.
 */
function globMatchFrom(toks: GlobToken[], text: string, t0: number, anchorEnd: boolean): boolean {
  let t = t0;
  let p = 0;
  let star = -1;
  let tStar = t0;
  for (;;) {
    if (p === toks.length) {
      if (!anchorEnd) return true; // pattern consumed → prefix match
      if (t === text.length) return true; // consumed all of text → full match
    } else {
      const tok = toks[p];
      if ('star' in tok) {
        star = p;
        tStar = t;
        p++;
        continue;
      }
      if (t < text.length && ('any' in tok || tok.lit === text[t])) {
        t++;
        p++;
        continue;
      }
    }
    // Mismatch: retry the last `*`, letting it consume one more character.
    if (star !== -1 && tStar < text.length) {
      tStar++;
      t = tStar;
      p = star + 1;
      continue;
    }
    return false;
  }
}

/** Compare text against a criterion pattern that may contain `*` / `?` wildcards. */
export function wildcardEquals(text: string, pattern: string): boolean {
  if (!/[*?~]/.test(pattern)) {
    return text.toLowerCase() === pattern.toLowerCase();
  }
  return globMatchFrom(parseGlob(pattern.toLowerCase()), text.toLowerCase(), 0, true);
}

/**
 * Earliest 0-based index in `text` where the wildcard `pattern` matches a
 * substring, or -1. Case-insensitive. Used by SEARCH; linear per start position,
 * so O(text²·pattern) worst case but never exponential (no ReDoS).
 */
export function wildcardSearch(text: string, pattern: string): number {
  const hay = text.toLowerCase();
  const toks = parseGlob(pattern.toLowerCase());
  for (let i = 0; i <= hay.length; i++) {
    if (globMatchFrom(toks, hay, i, false)) return i;
  }
  return -1;
}

/**
 * Test a value against a SUMIF/COUNTIF-style criterion: a bare value, or a
 * string carrying a comparison operator (`">5"`, `"<>x"`, `">=2"`).
 */
export function matchesCriterion(value: unknown, criterion: unknown): boolean {
  let op = '=';
  let operand: unknown = criterion;

  if (typeof criterion === 'string') {
    const m = criterion.match(/^(<=|>=|<>|=|<|>)([\s\S]*)$/);
    if (m) {
      op = m[1];
      operand = m[2];
    }
  }

  const valNum =
    typeof value === 'number'
      ? value
      : typeof value === 'string' && value.trim() !== '' && !isNaN(Number(value))
        ? Number(value)
        : null;
  const opNum =
    typeof operand === 'number'
      ? operand
      : typeof operand === 'string' && operand.trim() !== '' && !isNaN(Number(operand))
        ? Number(operand)
        : null;

  if (valNum !== null && opNum !== null) {
    switch (op) {
      case '=':
        return valNum === opNum;
      case '<>':
        return valNum !== opNum;
      case '>':
        return valNum > opNum;
      case '<':
        return valNum < opNum;
      case '>=':
        return valNum >= opNum;
      case '<=':
        return valNum <= opNum;
    }
  }

  const valStr = value == null ? '' : String(value);
  const opStr = operand == null ? '' : String(operand);
  switch (op) {
    case '=':
      return wildcardEquals(valStr, opStr);
    case '<>':
      return !wildcardEquals(valStr, opStr);
    case '>':
      return valStr.toLowerCase() > opStr.toLowerCase();
    case '<':
      return valStr.toLowerCase() < opStr.toLowerCase();
    case '>=':
      return valStr.toLowerCase() >= opStr.toLowerCase();
    case '<=':
      return valStr.toLowerCase() <= opStr.toLowerCase();
    default:
      return false;
  }
}

/** Normalise a value into a 2D matrix so table/array functions can index it. */
export function toMatrix(v: unknown): unknown[][] {
  if (Array.isArray(v)) {
    if (v.length > 0 && Array.isArray(v[0])) return v as unknown[][];
    return [v as unknown[]];
  }
  return [[v]];
}

/** Three-way comparison: numeric when both sides are numbers, else lexical. */
export function compareValues(a: unknown, b: unknown): number {
  // An empty cell takes the zero-value of the other operand (blank = 0, "" or FALSE).
  if (a == null) a = typeof b === 'number' ? 0 : typeof b === 'boolean' ? false : '';
  if (b == null) b = typeof a === 'number' ? 0 : typeof a === 'boolean' ? false : '';
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

/**
 * Excel-style equality for the `=` / `<>` operators. Unlike compareValues it is
 * type-aware: text never equals a number ("1" != 1, "" != 0), text compares
 * case-insensitively, and an empty cell (null) coerces to the other operand's
 * zero-value (blank = 0, blank = "", blank = FALSE).
 */
export function excelEqual(a: unknown, b: unknown): boolean {
  if (a == null) a = typeof b === 'number' ? 0 : typeof b === 'boolean' ? false : '';
  if (b == null) b = typeof a === 'number' ? 0 : typeof a === 'boolean' ? false : '';
  if (typeof a === 'number' && typeof b === 'number') return a === b;
  if (typeof a === 'boolean' && typeof b === 'boolean') return a === b;
  if (typeof a === 'string' && typeof b === 'string') return a.toLowerCase() === b.toLowerCase();
  return false;
}
