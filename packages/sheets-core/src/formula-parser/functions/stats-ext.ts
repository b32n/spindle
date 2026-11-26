// Statistical functions, batch A: the dotted (Excel 2010) names, exclusive
// percentiles, GEOMEAN, linear regression and FREQUENCY.

import type { EagerFn } from './helpers';
import { checkArity, strictNumbers, toNum, toStrictNum } from './helpers';
import { mean, numberPairs, percentile, statsFunctions } from './stats';

/** PERCENTILE.EXC: interpolation on rank k·(n+1); k must lie strictly inside (1/(n+1), n/(n+1)). */
export function percentileExc(sorted: number[], k: number): number {
  const n = sorted.length;
  const rank = k * (n + 1);
  if (n === 0 || rank < 1 || rank > n) throw new Error('#NUM!');
  const lo = Math.floor(rank);
  if (lo === rank) return sorted[lo - 1];
  return sorted[lo - 1] + (rank - lo) * (sorted[lo] - sorted[lo - 1]);
}

/** QUARTILE.INC: the quartile `quart` (0-4, truncated) of ascending data. */
export function quartileInc(sorted: number[], quart: number): number {
  const q = Math.trunc(quart);
  if (sorted.length === 0 || q < 0 || q > 4) throw new Error('#NUM!');
  return percentile(sorted, q / 4);
}

/** QUARTILE.EXC: the quartile `quart` (1-3, truncated) of ascending data. */
export function quartileExc(sorted: number[], quart: number): number {
  const q = Math.trunc(quart);
  if (q < 1 || q > 3) throw new Error('#NUM!');
  return percentileExc(sorted, q / 4);
}

/** Least-squares line through (xs, ys): slope, intercept and the sums needed for RSQ. */
function regression(ys: unknown, xs: unknown): { slope: number; intercept: number; r2: number } {
  const [y, x] = numberPairs(ys, xs);
  if (x.length < 2) throw new Error('#DIV/0!');
  const mx = mean(x);
  const my = mean(y);
  let sxy = 0;
  let sxx = 0;
  let syy = 0;
  for (let i = 0; i < x.length; i++) {
    sxy += (x[i] - mx) * (y[i] - my);
    sxx += (x[i] - mx) * (x[i] - mx);
    syy += (y[i] - my) * (y[i] - my);
  }
  if (sxx === 0) throw new Error('#DIV/0!');
  const slope = sxy / sxx;
  return { slope, intercept: my - slope * mx, r2: syy === 0 ? NaN : (sxy * sxy) / (sxx * syy) };
}

const forecast: EagerFn = (args) => {
  checkArity(args, 3);
  const x = toStrictNum(args[0]);
  const { slope, intercept } = regression(args[1], args[2]);
  return intercept + slope * x;
};

/** An existing function under another name, with Excel's argument count enforced. */
function alias(fn: EagerFn, min: number, max: number = min): EagerFn {
  return (args, ctx) => {
    checkArity(args, min, max);
    return fn(args, ctx);
  };
}

const sorted = (arg: unknown): number[] => strictNumbers([arg]).sort((a, b) => a - b);

export const statsExtFunctions: Record<string, EagerFn> = {
  'STDEV.S': alias(statsFunctions.STDEV, 1, Infinity),
  'STDEV.P': alias(statsFunctions.STDEVP, 1, Infinity),
  'VAR.S': alias(statsFunctions.VAR, 1, Infinity),
  'VAR.P': alias(statsFunctions.VARP, 1, Infinity),
  'PERCENTILE.INC': alias(statsFunctions.PERCENTILE, 2),
  'QUARTILE.INC': alias(statsFunctions.QUARTILE, 2),
  'RANK.EQ': alias(statsFunctions.RANK, 2, 3),
  'MODE.SNGL': alias(statsFunctions.MODE, 1, Infinity),

  'PERCENTILE.EXC': (args) => {
    checkArity(args, 2);
    return percentileExc(sorted(args[0]), toNum(args[1]));
  },

  'QUARTILE.EXC': (args) => {
    checkArity(args, 2);
    return quartileExc(sorted(args[0]), toNum(args[1]));
  },

  GEOMEAN: (args) => {
    const nums = strictNumbers(args);
    if (nums.length === 0 || nums.some((n) => n <= 0)) throw new Error('#NUM!');
    return Math.exp(mean(nums.map(Math.log)));
  },

  SLOPE: (args) => {
    checkArity(args, 2);
    return regression(args[0], args[1]).slope;
  },

  INTERCEPT: (args) => {
    checkArity(args, 2);
    return regression(args[0], args[1]).intercept;
  },

  RSQ: (args) => {
    checkArity(args, 2);
    const { r2 } = regression(args[0], args[1]);
    if (Number.isNaN(r2)) throw new Error('#DIV/0!');
    return r2;
  },

  FORECAST: forecast,
  'FORECAST.LINEAR': forecast,

  // Counts per bin as a column of bins+1 rows: bin i holds the values in (the
  // next smaller bin, bin i], the last row everything above the largest bin.
  FREQUENCY: (args) => {
    checkArity(args, 2);
    const data = strictNumbers([args[0]]);
    const bins = strictNumbers([args[1]]);
    const counts = new Array<number>(bins.length + 1).fill(0);
    // Bins by ascending value, the first of equal bins taking the count.
    const order = bins.map((_, i) => i).sort((a, b) => bins[a] - bins[b] || a - b);
    for (const d of data) {
      let lo = 0;
      let hi = order.length;
      while (lo < hi) {
        const mid = (lo + hi) >> 1;
        if (bins[order[mid]] >= d) hi = mid;
        else lo = mid + 1;
      }
      counts[lo === order.length ? bins.length : order[lo]]++;
    }
    return counts.map((c) => [c]);
  },
};
