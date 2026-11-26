// MAXIFS, MINIFS and AGGREGATE.

import type { EagerFn } from './helpers';
import {
  checkArity,
  flatten,
  isErrorValue,
  matchesCriterion,
  strictNumbers,
  toMatrix,
  toNum,
} from './helpers';
import { mean, median, modeOf, percentile, variance } from './stats';
import { percentileExc, quartileExc, quartileInc } from './stats-ext';

/**
 * Max or min of `valueRange` over the rows where every (criteria_range, criterion)
 * pair matches. 0 when nothing matches; an error value in a matching row propagates.
 */
function extremeIfs(args: unknown[], pick: (a: number, b: number) => number): number {
  if (args.length < 3 || args.length % 2 === 0) throw new Error('#VALUE!');
  const shape = (v: unknown): string => {
    const m = toMatrix(v);
    return `${m.length}x${m.reduce((w, row) => Math.max(w, row.length), 0)}`;
  };
  const values = flatten([args[0]]);
  const ranges: unknown[][] = [];
  const criteria: unknown[] = [];
  for (let i = 1; i < args.length; i += 2) {
    const range = flatten([args[i]]);
    if (shape(args[i]) !== shape(args[0])) throw new Error('#VALUE!');
    ranges.push(range);
    criteria.push(args[i + 1]);
  }
  let result: number | null = null;
  for (let i = 0; i < values.length; i++) {
    if (!ranges.every((r, k) => matchesCriterion(r[i], criteria[k]))) continue;
    const v = values[i];
    if (typeof v === 'string' && isErrorValue(v)) throw new Error(v);
    if (typeof v === 'number' && !isNaN(v)) result = result === null ? v : pick(result, v);
  }
  return result ?? 0;
}

const need = (nums: number[]): number[] => {
  if (nums.length === 0) throw new Error('#DIV/0!');
  return nums;
};

/** AGGREGATE function numbers 1-13 (reference form): the values are the pooled arguments. */
const REFERENCE_FORMS: Array<(nums: number[], all: unknown[]) => number> = [
  (n) => mean(need(n)), // 1 AVERAGE
  (n) => n.length, // 2 COUNT
  (_n, all) => all.filter((v) => v != null && v !== '').length, // 3 COUNTA
  (n) => n.reduce((a, b) => (b > a ? b : a), n.length ? n[0] : 0), // 4 MAX
  (n) => n.reduce((a, b) => (b < a ? b : a), n.length ? n[0] : 0), // 5 MIN
  (n) => (n.length ? n.reduce((a, b) => a * b, 1) : 0), // 6 PRODUCT
  (n) => Math.sqrt(variance(n, true)), // 7 STDEV.S
  (n) => Math.sqrt(variance(n, false)), // 8 STDEV.P
  (n) => n.reduce((a, b) => a + b, 0), // 9 SUM
  (n) => variance(n, true), // 10 VAR.S
  (n) => variance(n, false), // 11 VAR.P
  median, // 12
  modeOf, // 13 MODE.SNGL
];

/** AGGREGATE function numbers 14-19 (array form): `k` is the fourth argument. */
const ARRAY_FORMS: Array<(sorted: number[], k: number) => number> = [
  (s, k) => {
    // 14 LARGE
    const i = Math.trunc(k);
    if (i < 1 || i > s.length) throw new Error('#NUM!');
    return s[s.length - i];
  },
  (s, k) => {
    // 15 SMALL
    const i = Math.trunc(k);
    if (i < 1 || i > s.length) throw new Error('#NUM!');
    return s[i - 1];
  },
  (s, k) => {
    // 16 PERCENTILE.INC
    if (s.length === 0 || k < 0 || k > 1) throw new Error('#NUM!');
    return percentile(s, k);
  },
  quartileInc, // 17
  percentileExc, // 18
  quartileExc, // 19
];

export const aggregateFunctions: Record<string, EagerFn> = {
  MAXIFS: (args) => extremeIfs(args, Math.max),
  MINIFS: (args) => extremeIfs(args, Math.min),

  // The options that ignore hidden rows or nested SUBTOTAL/AGGREGATE results (1, 3, 5, 7) are
  // accepted but act like 0, 2, 4, 6: the evaluation context exposes neither row visibility nor
  // which cells hold subtotals. Options 2, 3, 6 and 7 skip error values; the others propagate them.
  AGGREGATE: (args) => {
    const fn = Math.trunc(toNum(args[0]));
    const option = Math.trunc(toNum(args[1]));
    if (fn < 1 || fn > 19 || option < 0 || option > 7) throw new Error('#VALUE!');
    const ignoreErrors = option === 2 || option === 3 || option === 6 || option === 7;
    const clean = (values: unknown[]): unknown[] => {
      const out: unknown[] = [];
      for (const v of values) {
        if (typeof v === 'string' && isErrorValue(v)) {
          if (!ignoreErrors) throw new Error(v);
        } else out.push(v);
      }
      return out;
    };
    if (fn <= 13) {
      checkArity(args, 3, Infinity);
      const pooled = flatten(args.slice(2));
      // COUNT and COUNTA never propagate an error cell: COUNT skips it, COUNTA counts it.
      const counting = fn === 2 || fn === 3;
      const all = counting && !ignoreErrors ? pooled : clean(pooled);
      const numeric = counting ? all.filter((v) => !isErrorValue(v)) : all;
      return REFERENCE_FORMS[fn - 1](strictNumbers(numeric), all);
    }
    checkArity(args, 4);
    const sorted = strictNumbers(clean(flatten([args[2]]))).sort((a, b) => a - b);
    return ARRAY_FORMS[fn - 14](sorted, toNum(args[3]));
  },
};
