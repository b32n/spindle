// Math functions, batch A: trigonometry, factorials and combinatorics, parity,
// SUMSQ and the CEILING.MATH / FLOOR.MATH rounding pair; batch B adds QUOTIENT, SQRTPI and the
// SUMX2MY2 / SUMX2PY2 / SUMXMY2 sums over pairs of arrays.
//
// Scalar functions here are lifted over arrays (`=SIN(A1:A3)` spills one result
// per cell, a failing element carries its own error value).

import type { EagerFn } from './helpers';
import { checkArity, flatten, isErrorValue, liftScalar, strictNumbers, toMatrix, toStrictNum } from './helpers';

/** The result of a numeric function, or `#NUM!` when it overflowed to Infinity / NaN (and never -0). */
function finite(n: number): number {
  if (!Number.isFinite(n)) throw new Error('#NUM!');
  return n === 0 ? 0 : n; // no negative zero: Excel shows 0
}

/**
 * A quotient snapped to the nearest integer when it is within floating-point
 * noise of one, so `0.9 / 0.3` (3.0000000000000004) does not ceil up to 4.
 */
function snap(q: number): number {
  const r = Math.round(q);
  return Math.abs(q - r) <= 4 * Number.EPSILON * Math.max(1, Math.abs(q)) ? r : q;
}

/** A one-number function lifted over arrays. */
function unary(fn: (n: number) => number): EagerFn {
  return (args) => {
    checkArity(args, 1);
    return liftScalar(args, ([x]) => finite(fn(toStrictNum(x))));
  };
}

/** An n-ary number function lifted over arrays (optional trailing arguments arrive as `undefined`). */
function scalarFn(min: number, max: number, fn: (nums: number[]) => unknown): EagerFn {
  return (args) => {
    checkArity(args, min, max);
    return liftScalar(args, (xs) =>
      fn(Array.from({ length: max }, (_, i) => (xs[i] === undefined ? NaN : toStrictNum(xs[i]))))
    );
  };
}

/** `n!` for a non-negative integer, or `#NUM!` beyond the double range (n > 170). */
function factorial(n: number): number {
  let r = 1;
  for (let i = 2; i <= n; i++) r = finite(r * i);
  return r;
}

/** Parity check shared by ISEVEN / ISODD: booleans and text are `#VALUE!`, blank is 0. */
function integerPart(v: unknown): number {
  if (typeof v === 'boolean') throw new Error('#VALUE!');
  return Math.trunc(toStrictNum(v));
}

/** CEILING.MATH / FLOOR.MATH: `up` selects the direction on the positive side. */
function roundMath(n: number, significance: number, mode: number, up: boolean): number {
  if (significance === 0 || n === 0) return 0;
  const q = snap(n / Math.abs(significance));
  // Negative numbers go the opposite way unless a non-zero mode asks for "away from zero".
  const ceil = n > 0 ? up : mode === 0 ? up : !up;
  return finite((ceil ? Math.ceil(q) : Math.floor(q)) * Math.abs(significance));
}

/**
 * A sum over the pairs of two same-shaped arrays (SUMX2MY2 and its relatives). A pair counts when both
 * elements are numbers; text, logicals and blanks are skipped, an error value is the result, and arrays
 * of different shapes are `#N/A`.
 */
function pairSum(term: (x: number, y: number) => number): EagerFn {
  return (args) => {
    checkArity(args, 2);
    const [xm, ym] = [toMatrix(args[0]), toMatrix(args[1])];
    const width = (m: unknown[][]): number => m.reduce((w, row) => Math.max(w, row.length), 0);
    if (xm.length !== ym.length || width(xm) !== width(ym)) throw new Error('#N/A');
    const xs = flatten([args[0]]);
    const ys = flatten([args[1]]);
    let total = 0;
    for (let i = 0; i < xs.length; i++) {
      for (const v of [xs[i], ys[i]]) if (typeof v === 'string' && isErrorValue(v)) throw new Error(v);
      if (typeof xs[i] === 'number' && typeof ys[i] === 'number') total += term(xs[i] as number, ys[i] as number);
    }
    return finite(total);
  };
}

export const mathExtFunctions: Record<string, EagerFn> = {
  SIN: unary(Math.sin),
  COS: unary(Math.cos),
  TAN: unary(Math.tan),

  ASIN: unary((n) => {
    if (n < -1 || n > 1) throw new Error('#NUM!');
    return Math.asin(n);
  }),

  ACOS: unary((n) => {
    if (n < -1 || n > 1) throw new Error('#NUM!');
    return Math.acos(n);
  }),

  ATAN: unary(Math.atan),

  // Excel's argument order is (x_num, y_num); the angle is atan2(y, x).
  ATAN2: scalarFn(2, 2, ([x, y]) => {
    if (x === 0 && y === 0) throw new Error('#DIV/0!');
    return Math.atan2(y, x);
  }),

  SINH: unary(Math.sinh),
  COSH: unary(Math.cosh),
  TANH: unary(Math.tanh),
  ASINH: unary(Math.asinh),

  ACOSH: unary((n) => {
    if (n < 1) throw new Error('#NUM!');
    return Math.acosh(n);
  }),

  ATANH: unary((n) => {
    if (n <= -1 || n >= 1) throw new Error('#NUM!');
    return Math.atanh(n);
  }),

  PI: (args) => {
    checkArity(args, 0);
    return Math.PI;
  },

  RADIANS: unary((n) => (n * Math.PI) / 180),
  DEGREES: unary((n) => (n * 180) / Math.PI),

  FACT: unary((n) => {
    if (n < 0) throw new Error('#NUM!');
    return factorial(Math.trunc(n));
  }),

  FACTDOUBLE: unary((n) => {
    // -1!! is 1 (Excel); anything below is #NUM!.
    if (n < -1) throw new Error('#NUM!');
    let r = 1;
    for (let i = Math.trunc(n); i > 1; i -= 2) r = finite(r * i);
    return r;
  }),

  COMBIN: scalarFn(2, 2, ([a, b]) => {
    const n = Math.trunc(a);
    const k = Math.trunc(b);
    if (n < 0 || k < 0 || k > n) throw new Error('#NUM!');
    const small = Math.min(k, n - k);
    let r = 1;
    // Multiplicative form: every partial product is itself a binomial coefficient.
    for (let i = 1; i <= small; i++) r = finite((r * (n - small + i)) / i);
    return Math.round(r);
  }),

  PERMUT: scalarFn(2, 2, ([a, b]) => {
    const n = Math.trunc(a);
    const k = Math.trunc(b);
    if (n < 0 || k < 0 || k > n) throw new Error('#NUM!');
    let r = 1;
    for (let i = 0; i < k; i++) r = finite(r * (n - i));
    return r;
  }),

  SUMSQ: (args) => strictNumbers(args).reduce((acc, n) => acc + n * n, 0),

  EVEN: unary((n) => Math.sign(n) * Math.ceil(Math.abs(n) / 2) * 2),

  ODD: unary((n) => {
    const up = Math.ceil(Math.abs(n));
    return (n < 0 ? -1 : 1) * (up % 2 === 0 ? up + 1 : up);
  }),

  ISEVEN: (args) => {
    checkArity(args, 1);
    return liftScalar(args, ([v]) => integerPart(v) % 2 === 0);
  },

  ISODD: (args) => {
    checkArity(args, 1);
    return liftScalar(args, ([v]) => integerPart(v) % 2 !== 0);
  },

  'CEILING.MATH': scalarFn(1, 3, ([n, significance, mode]) =>
    roundMath(n, Number.isNaN(significance) ? 1 : significance, Number.isNaN(mode) ? 0 : mode, true)
  ),

  'FLOOR.MATH': scalarFn(1, 3, ([n, significance, mode]) =>
    roundMath(
      n,
      Number.isNaN(significance) ? 1 : significance,
      Number.isNaN(mode) ? 0 : mode,
      false
    )
  ),

  // The integer part of a quotient, snapped first so 0.6 / 0.2 is 3 and not 2.9999999999999996.
  QUOTIENT: scalarFn(2, 2, ([n, d]) => {
    if (d === 0) throw new Error('#DIV/0!');
    return Math.trunc(snap(n / d)) + 0; // + 0: never -0
  }),

  SQRTPI: unary((n) => {
    if (n < 0) throw new Error('#NUM!');
    return Math.sqrt(n * Math.PI);
  }),

  SUMX2MY2: pairSum((x, y) => x * x - y * y),
  SUMX2PY2: pairSum((x, y) => x * x + y * y),
  SUMXMY2: pairSum((x, y) => (x - y) * (x - y)),
};
