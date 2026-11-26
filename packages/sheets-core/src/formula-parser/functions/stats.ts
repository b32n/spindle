// Statistical functions.

import type { EagerFn } from './helpers';
import { flatten, isErrorValue, strictNumbers, toNum } from './helpers';

export function mean(nums: number[]): number {
  return nums.reduce((a, b) => a + b, 0) / nums.length;
}

/** Population (sample=false) or sample (sample=true) variance. */
export function variance(nums: number[], sample: boolean): number {
  const n = nums.length;
  if (sample ? n < 2 : n < 1) throw new Error('#DIV/0!');
  const m = mean(nums);
  const ss = nums.reduce((acc, x) => acc + (x - m) * (x - m), 0);
  return ss / (sample ? n - 1 : n);
}

/** The middle value (mean of the two middle values for an even count); `#NUM!` when empty. */
export function median(nums: number[]): number {
  if (nums.length === 0) throw new Error('#NUM!');
  const s = [...nums].sort((a, b) => a - b);
  const mid = Math.floor(s.length / 2);
  return s.length % 2 ? s[mid] : (s[mid - 1] + s[mid]) / 2;
}

/** The value occurring most often; among ties the one that appears first (as Excel does). `#N/A` when all are distinct. */
export function modeOf(nums: number[]): number {
  const counts = new Map<number, number>();
  for (const n of nums) counts.set(n, (counts.get(n) ?? 0) + 1);
  let best: number | null = null;
  let bestCount = 1;
  for (const n of nums) {
    const c = counts.get(n) as number;
    if (c > bestCount) {
      bestCount = c;
      best = n;
    }
  }
  if (best === null) throw new Error('#N/A');
  return best;
}

/** Linear-interpolation percentile (the inclusive variant Excel uses). */
export function percentile(sorted: number[], k: number): number {
  const idx = k * (sorted.length - 1);
  const lo = Math.floor(idx);
  const hi = Math.ceil(idx);
  if (lo === hi) return sorted[lo];
  return sorted[lo] + (idx - lo) * (sorted[hi] - sorted[lo]);
}

/** Coerce a single value to a number, or null when it is not numeric. */
function numericOf(v: unknown): number | null {
  if (typeof v === 'number') return isNaN(v) ? null : v;
  if (typeof v === 'string' && v.trim() !== '' && !isNaN(Number(v))) return Number(v);
  return null;
}

/** Pair two ranges index-wise, dropping pairs where either side is non-numeric (an error propagates). */
export function numberPairs(a: unknown, b: unknown): [number[], number[]] {
  const fa = flatten([a]);
  const fb = flatten([b]);
  if (fa.length !== fb.length) throw new Error('#N/A');
  const xs: number[] = [];
  const ys: number[] = [];
  for (let i = 0; i < fa.length; i++) {
    // An error value in either series propagates, like every other statistic.
    for (const v of [fa[i], fb[i]]) if (typeof v === 'string' && isErrorValue(v)) throw new Error(v);
    const x = numericOf(fa[i]);
    const y = numericOf(fb[i]);
    if (x !== null && y !== null) {
      xs.push(x);
      ys.push(y);
    }
  }
  return [xs, ys];
}

export const statsFunctions: Record<string, EagerFn> = {
  MEDIAN: (args) => median(strictNumbers(args)),

  MODE: (args) => modeOf(strictNumbers(args)),

  VAR: (args) => variance(strictNumbers(args), true),

  VARP: (args) => variance(strictNumbers(args), false),

  STDEV: (args) => Math.sqrt(variance(strictNumbers(args), true)),

  STDEVP: (args) => Math.sqrt(variance(strictNumbers(args), false)),

  PERCENTILE: (args) => {
    const nums = strictNumbers([args[0]]).sort((a, b) => a - b);
    const k = toNum(args[1]);
    if (nums.length === 0 || k < 0 || k > 1) throw new Error('#NUM!');
    return percentile(nums, k);
  },

  QUARTILE: (args) => {
    const nums = strictNumbers([args[0]]).sort((a, b) => a - b);
    const quart = Math.trunc(toNum(args[1]));
    if (nums.length === 0 || quart < 0 || quart > 4) throw new Error('#NUM!');
    return percentile(nums, quart / 4);
  },

  RANK: (args) => {
    const number = toNum(args[0]);
    const ref = strictNumbers([args[1]]);
    const ascending = args[2] !== undefined && toNum(args[2]) !== 0;
    const sorted = [...ref].sort((a, b) => (ascending ? a - b : b - a));
    const idx = sorted.indexOf(number);
    if (idx < 0) throw new Error('#N/A');
    return idx + 1;
  },

  LARGE: (args) => {
    const nums = strictNumbers([args[0]]).sort((a, b) => b - a);
    const k = Math.trunc(toNum(args[1]));
    if (k < 1 || k > nums.length) throw new Error('#NUM!');
    return nums[k - 1];
  },

  SMALL: (args) => {
    const nums = strictNumbers([args[0]]).sort((a, b) => a - b);
    const k = Math.trunc(toNum(args[1]));
    if (k < 1 || k > nums.length) throw new Error('#NUM!');
    return nums[k - 1];
  },

  CORREL: (args) => {
    const [x, y] = numberPairs(args[0], args[1]);
    if (x.length === 0) throw new Error('#DIV/0!');
    const mx = mean(x);
    const my = mean(y);
    let sxy = 0;
    let sxx = 0;
    let syy = 0;
    for (let i = 0; i < x.length; i++) {
      const dx = x[i] - mx;
      const dy = y[i] - my;
      sxy += dx * dy;
      sxx += dx * dx;
      syy += dy * dy;
    }
    if (sxx === 0 || syy === 0) throw new Error('#DIV/0!');
    return sxy / Math.sqrt(sxx * syy);
  },

  COVAR: (args) => {
    const [x, y] = numberPairs(args[0], args[1]);
    if (x.length === 0) throw new Error('#DIV/0!');
    const mx = mean(x);
    const my = mean(y);
    let s = 0;
    for (let i = 0; i < x.length; i++) {
      s += (x[i] - mx) * (y[i] - my);
    }
    return s / x.length;
  },
};
