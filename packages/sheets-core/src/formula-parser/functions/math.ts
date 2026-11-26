// Math and aggregation functions.

import type { EagerFn } from './helpers';
import { flatten, toNum, strictNumbers, matchesCriterion, toSumNum, isErrorValue } from './helpers';

/** Number of a summed cell, propagating an error value but skipping non-numeric text. */
function summedNumberOrThrow(v: unknown): number | null {
  if (typeof v === 'string' && isErrorValue(v)) throw new Error(v);
  const n = Number(v);
  return isNaN(n) ? null : n;
}

function sum(nums: number[]): number {
  return nums.reduce((a, b) => a + b, 0);
}

function gcd2(a: number, b: number): number {
  a = Math.abs(Math.trunc(a));
  b = Math.abs(Math.trunc(b));
  while (b) {
    [a, b] = [b, a % b];
  }
  return a;
}

/** Population (sample=false) or sample (sample=true) variance. */
function variance(nums: number[], sample: boolean): number {
  if (sample ? nums.length < 2 : nums.length < 1) {
    throw new Error('#DIV/0!');
  }
  const mean = sum(nums) / nums.length;
  const ss = nums.reduce((acc, n) => acc + (n - mean) * (n - mean), 0);
  return ss / (sample ? nums.length - 1 : nums.length);
}

/** Round away from zero at the given number of decimal places. */
function roundWith(n: number, digits: number, mode: 'half' | 'up' | 'down'): number {
  const factor = Math.pow(10, digits);
  const scaled = Math.abs(n) * factor;
  const sign = n < 0 ? -1 : 1;
  const op = mode === 'half' ? Math.round : mode === 'up' ? Math.ceil : Math.floor;
  return (sign * op(scaled)) / factor;
}

/**
 * Apply a SUMIF/COUNTIF/AVERAGEIF-style set of (range, criterion) pairs and
 * collect the indices of rows where every criterion matches.
 */
function matchingIndices(pairs: unknown[]): number[] {
  const ranges: unknown[][] = [];
  const criteria: unknown[] = [];
  for (let i = 0; i + 1 < pairs.length; i += 2) {
    ranges.push(flatten([pairs[i]]));
    criteria.push(pairs[i + 1]);
  }
  if (ranges.length === 0) return [];
  const length = ranges[0].length;
  const result: number[] = [];
  for (let i = 0; i < length; i++) {
    if (ranges.every((r, k) => matchesCriterion(r[i], criteria[k]))) {
      result.push(i);
    }
  }
  return result;
}

export const mathFunctions: Record<string, EagerFn> = {
  SUM: (args) => sum(strictNumbers(args)),

  AVERAGE: (args) => {
    const nums = strictNumbers(args);
    if (nums.length === 0) throw new Error('#DIV/0!');
    return sum(nums) / nums.length;
  },

  COUNT: (args) => flatten(args).filter((v) => typeof v === 'number' && !isNaN(v)).length,

  COUNTA: (args) => flatten(args).filter((v) => v != null && v !== '').length,

  COUNTBLANK: (args) => flatten(args).filter((v) => v == null || v === '').length,

  MAX: (args) => {
    const nums = strictNumbers(args);
    return nums.length > 0 ? nums.reduce((a, b) => (b > a ? b : a)) : 0;
  },

  MIN: (args) => {
    const nums = strictNumbers(args);
    return nums.length > 0 ? nums.reduce((a, b) => (b < a ? b : a)) : 0;
  },

  PRODUCT: (args) => {
    const nums = strictNumbers(args);
    return nums.length > 0 ? nums.reduce((a, b) => a * b, 1) : 0;
  },

  SUMPRODUCT: (args) => {
    const lists = args.map((a) => flatten([a]));
    if (lists.length === 0) return 0;
    const length = lists[0].length;
    if (lists.some((l) => l.length !== length)) throw new Error('#VALUE!');
    let total = 0;
    for (let i = 0; i < length; i++) {
      let product = 1;
      for (const list of lists) {
        product *= toSumNum(list[i]); // error propagates; text/blank → 0
      }
      total += product;
    }
    return total;
  },

  SUMIF: (args) => {
    const range = flatten([args[0]]);
    const criterion = args[1];
    const sumRange = args[2] !== undefined ? flatten([args[2]]) : range;
    let total = 0;
    for (let i = 0; i < range.length; i++) {
      if (matchesCriterion(range[i], criterion)) {
        const n = summedNumberOrThrow(sumRange[i]);
        if (n !== null) total += n;
      }
    }
    return total;
  },

  SUMIFS: (args) => {
    const sumRange = flatten([args[0]]);
    const indices = matchingIndices(args.slice(1));
    let total = 0;
    for (const i of indices) {
      const n = summedNumberOrThrow(sumRange[i]);
      if (n !== null) total += n;
    }
    return total;
  },

  COUNTIF: (args) => {
    const range = flatten([args[0]]);
    const criterion = args[1];
    return range.filter((v) => matchesCriterion(v, criterion)).length;
  },

  COUNTIFS: (args) => matchingIndices(args).length,

  AVERAGEIF: (args) => {
    const range = flatten([args[0]]);
    const criterion = args[1];
    const avgRange = args[2] !== undefined ? flatten([args[2]]) : range;
    const matched: number[] = [];
    for (let i = 0; i < range.length; i++) {
      if (matchesCriterion(range[i], criterion)) {
        const n = summedNumberOrThrow(avgRange[i]);
        if (n !== null) matched.push(n);
      }
    }
    if (matched.length === 0) throw new Error('#DIV/0!');
    return sum(matched) / matched.length;
  },

  AVERAGEIFS: (args) => {
    const avgRange = flatten([args[0]]);
    const indices = matchingIndices(args.slice(1));
    const matched: number[] = [];
    for (const i of indices) {
      const n = summedNumberOrThrow(avgRange[i]);
      if (n !== null) matched.push(n);
    }
    if (matched.length === 0) throw new Error('#DIV/0!');
    return sum(matched) / matched.length;
  },

  SUBTOTAL: (args) => {
    const raw = toNum(args[0]);
    const code = raw > 100 ? raw - 100 : raw;
    const rest = args.slice(1);
    const nums = strictNumbers(rest);
    switch (code) {
      case 1:
        if (nums.length === 0) throw new Error('#DIV/0!');
        return sum(nums) / nums.length;
      case 2:
        return flatten(rest).filter((v) => typeof v === 'number' && !isNaN(v)).length;
      case 3:
        return flatten(rest).filter((v) => v != null && v !== '').length;
      case 4:
        return nums.length > 0 ? nums.reduce((a, b) => (b > a ? b : a)) : 0;
      case 5:
        return nums.length > 0 ? nums.reduce((a, b) => (b < a ? b : a)) : 0;
      case 6:
        return nums.length > 0 ? nums.reduce((a, b) => a * b, 1) : 0;
      case 7:
        return Math.sqrt(variance(nums, true));
      case 8:
        return Math.sqrt(variance(nums, false));
      case 9:
        return sum(nums);
      case 10:
        return variance(nums, true);
      case 11:
        return variance(nums, false);
      default:
        throw new Error('#VALUE!');
    }
  },

  ABS: (args) => Math.abs(toNum(args[0])),

  SIGN: (args) => {
    const n = toNum(args[0]);
    return n > 0 ? 1 : n < 0 ? -1 : 0;
  },

  INT: (args) => Math.floor(toNum(args[0])),

  TRUNC: (args) => {
    const n = toNum(args[0]);
    const digits = args[1] !== undefined ? toNum(args[1]) : 0;
    const factor = Math.pow(10, digits);
    return Math.trunc(n * factor) / factor;
  },

  ROUND: (args) => roundWith(toNum(args[0]), args[1] !== undefined ? toNum(args[1]) : 0, 'half'),

  ROUNDUP: (args) => roundWith(toNum(args[0]), args[1] !== undefined ? toNum(args[1]) : 0, 'up'),

  ROUNDDOWN: (args) => roundWith(toNum(args[0]), args[1] !== undefined ? toNum(args[1]) : 0, 'down'),

  MROUND: (args) => {
    const n = toNum(args[0]);
    const multiple = toNum(args[1]);
    if (multiple === 0) return 0;
    if (Math.sign(n) !== Math.sign(multiple) && n !== 0) throw new Error('#NUM!');
    return Math.round(n / multiple) * multiple;
  },

  CEILING: (args) => {
    const n = toNum(args[0]);
    const significance = args[1] !== undefined ? toNum(args[1]) : 1;
    if (significance === 0) return 0;
    // Excel: a positive number with negative significance is #NUM! (signs must
    // agree unless the number is 0).
    if (n > 0 && significance < 0) throw new Error('#NUM!');
    return Math.ceil(n / significance) * significance;
  },

  FLOOR: (args) => {
    const n = toNum(args[0]);
    const significance = args[1] !== undefined ? toNum(args[1]) : 1;
    if (significance === 0) return 0;
    if (n > 0 && significance < 0) throw new Error('#NUM!');
    return Math.floor(n / significance) * significance;
  },

  MOD: (args) => {
    const n = toNum(args[0]);
    const divisor = toNum(args[1]);
    if (divisor === 0) throw new Error('#DIV/0!');
    return n - divisor * Math.floor(n / divisor);
  },

  POWER: (args) => {
    const result = Math.pow(toNum(args[0]), toNum(args[1]));
    if (!isFinite(result)) throw new Error('#NUM!');
    return result;
  },

  SQRT: (args) => {
    const n = toNum(args[0]);
    if (n < 0) throw new Error('#NUM!');
    return Math.sqrt(n);
  },

  EXP: (args) => Math.exp(toNum(args[0])),

  LN: (args) => {
    const n = toNum(args[0]);
    if (n <= 0) throw new Error('#NUM!');
    return Math.log(n);
  },

  LOG: (args) => {
    const n = toNum(args[0]);
    const base = args[1] !== undefined ? toNum(args[1]) : 10;
    if (n <= 0 || base <= 0 || base === 1) throw new Error('#NUM!');
    return Math.log(n) / Math.log(base);
  },

  LOG10: (args) => {
    const n = toNum(args[0]);
    if (n <= 0) throw new Error('#NUM!');
    return Math.log10(n);
  },

  GCD: (args) => {
    const nums = flatten(args).map(toNum);
    if (nums.some((n) => n < 0)) throw new Error('#NUM!');
    return nums.reduce((a, b) => gcd2(a, b), 0);
  },

  LCM: (args) => {
    const nums = flatten(args).map(toNum);
    if (nums.some((n) => n < 0)) throw new Error('#NUM!');
    return nums.reduce((a, b) => {
      const x = Math.abs(Math.trunc(a));
      const y = Math.abs(Math.trunc(b));
      if (x === 0 || y === 0) return 0;
      return (x / gcd2(x, y)) * y;
    }, 1);
  },

  RAND: () => Math.random(),

  RANDBETWEEN: (args) => {
    const low = Math.ceil(toNum(args[0]));
    const high = Math.floor(toNum(args[1]));
    if (low > high) throw new Error('#NUM!');
    return Math.floor(Math.random() * (high - low + 1)) + low;
  },
};
