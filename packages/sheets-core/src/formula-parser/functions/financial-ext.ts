// Financial functions, batch B: XNPV, XIRR, MIRR, SLN, SYD, DDB, EFFECT, NOMINAL, RRI, PDURATION, ISPMT and FVSCHEDULE.
//
// The scalar functions lift over arrays like the rest of the engine. Argument errors follow Excel: `#NUM!` for a
// value out of range (LibreOffice reports `#VALUE!` for most of them, so the conformance rows that hit one are
// Excel-documented and tagged unverified).

import type { EagerFn } from './helpers';
import { checkArity, flatten, isErrorValue, liftScalar, strictNumbers, toNum } from './helpers';
import { toSerial } from './date';

/** A finite result, or `#NUM!`. */
function finite(n: number): number {
  if (!Number.isFinite(n)) throw new Error('#NUM!');
  return n === 0 ? 0 : n;
}

/** The numbers of a cash-flow argument in order (an empty cell is 0; text and logicals are `#VALUE!`). */
function cashFlows(arg: unknown): number[] {
  return flatten([arg]).map((v) => {
    if (typeof v === 'number') return v;
    if (v === null || v === undefined) return 0;
    if (typeof v === 'string' && isErrorValue(v)) throw new Error(v);
    throw new Error('#VALUE!');
  });
}

/** The whole-day serials of a dates argument in order (text dates are read, other text is `#VALUE!`). */
function cashDates(arg: unknown): number[] {
  return flatten([arg]).map((v) => {
    if (typeof v === 'boolean') throw new Error('#VALUE!');
    const day = Math.floor(toSerial(v));
    if (day < 0) throw new Error('#NUM!'); // a date before 1900-01-00
    return day;
  });
}

/** Years from the first date, on a 365-day year, for each cash flow of an XNPV / XIRR. */
function yearsFromStart(values: number[], dates: number[]): number[] {
  if (values.length !== dates.length || values.length === 0) throw new Error('#NUM!');
  const start = dates[0];
  return dates.map((d) => {
    if (d < start) throw new Error('#NUM!');
    return (d - start) / 365;
  });
}

/** Net present value of dated cash flows at `rate` (> -1). */
function xnpv(rate: number, values: number[], years: number[]): number {
  let total = 0;
  for (let i = 0; i < values.length; i++) total += values[i] * Math.pow(1 + rate, -years[i]);
  return total;
}

/** Slope of `xnpv` in the rate. */
function xnpvSlope(rate: number, values: number[], years: number[]): number {
  let total = 0;
  for (let i = 0; i < values.length; i++) total -= (years[i] * values[i]) * Math.pow(1 + rate, -years[i] - 1);
  return total;
}

/**
 * A root of `f` on r > -1 found by scanning a grid of rates r = e^s - 1 (from just above -100% to 10^10) for a sign
 * change and bisecting the bracket nearest `guess`; undefined when `f` never changes sign there.
 */
export function bracketedRoot(f: (r: number) => number, guess: number): number | undefined {
  const steps = 600;
  const lo = Math.log(1e-9);
  const hi = Math.log(1e10);
  const rateAt = (s: number): number => Math.exp(s) - 1;
  let best: [number, number] | undefined;
  let previous = { s: lo, v: f(rateAt(lo)) };
  for (let k = 1; k <= steps; k++) {
    const s = lo + ((hi - lo) * k) / steps;
    const v = f(rateAt(s));
    // a sign change (compared as signs: a product of two tiny values would underflow to 0); an exact zero is its own root
    if (Number.isFinite(v) && v === 0) return rateAt(s);
    if (Number.isFinite(v) && Number.isFinite(previous.v) && previous.v !== 0 && (previous.v < 0) !== (v < 0)) {
      const distance = (pair: [number, number]): number => Math.abs(rateAt((pair[0] + pair[1]) / 2) - guess);
      if (best === undefined || distance([previous.s, s]) < distance(best)) best = [previous.s, s];
    }
    previous = { s, v };
  }
  if (best === undefined) return undefined;
  let [a, b] = best;
  const fa = f(rateAt(a));
  for (let i = 0; i < 200; i++) {
    const mid = (a + b) / 2;
    if ((f(rateAt(mid)) < 0) === (fa < 0)) a = mid;
    else b = mid;
  }
  return rateAt((a + b) / 2);
}

/** The rate at which the dated cash flows net to zero: Newton from the guess, else a bracketed search. */
function solveXirr(values: number[], years: number[], guess: number): number {
  const f = (r: number): number => xnpv(r, values, years);
  const scale = values.reduce((sum, v) => sum + Math.abs(v), 0);
  let rate = guess;
  for (let i = 0; i < 100; i++) {
    const value = f(rate);
    if (value === 0) return rate; // an exact root
    const slope = xnpvSlope(rate, values, years);
    if (!Number.isFinite(value) || !Number.isFinite(slope) || slope === 0) break;
    let next = rate - value / slope;
    if (next <= -1) next = (rate - 1) / 2; // stay inside the domain: halve the way to -1
    if (Math.abs(next - rate) <= 1e-14 * Math.max(1, Math.abs(rate))) {
      // a vanishing step is a root only if the cash flows really net to nothing there (not a flat tail next to -100%)
      if (Math.abs(f(next)) <= 1e-9 * scale) return next;
      break;
    }
    rate = next;
  }
  const root = bracketedRoot(f, guess);
  if (root === undefined) throw new Error('#NUM!');
  return root;
}

export const financialExtFunctions: Record<string, EagerFn> = {
  XNPV: (args) => {
    checkArity(args, 3);
    const rate = toNum(args[0]);
    const values = cashFlows(args[1]);
    const years = yearsFromStart(values, cashDates(args[2]));
    if (rate <= -1) throw new Error('#NUM!');
    return finite(xnpv(rate, values, years));
  },

  XIRR: (args) => {
    checkArity(args, 2, 3);
    const values = cashFlows(args[0]);
    const years = yearsFromStart(values, cashDates(args[1]));
    const guess = args[2] === undefined ? 0.1 : toNum(args[2]);
    if (guess <= -1 || !values.some((v) => v > 0) || !values.some((v) => v < 0)) throw new Error('#NUM!');
    return finite(solveXirr(values, years, guess));
  },

  MIRR: (args) => {
    checkArity(args, 3);
    const flows = strictNumbers([args[0]]);
    const financeRate = toNum(args[1]);
    const reinvestRate = toNum(args[2]);
    const n = flows.length;
    if (n < 2 || !flows.some((v) => v > 0) || !flows.some((v) => v < 0)) throw new Error('#DIV/0!');
    let negativesNow = 0; // value now of the negative flows, at the finance rate
    let positivesLater = 0; // value at the end of the positive flows, at the reinvestment rate
    for (let i = 0; i < n; i++) {
      if (flows[i] < 0) negativesNow += flows[i] / Math.pow(1 + financeRate, i);
      else positivesLater += flows[i] * Math.pow(1 + reinvestRate, n - 1 - i);
    }
    if (!Number.isFinite(negativesNow) || negativesNow === 0) throw new Error('#DIV/0!');
    const ratio = positivesLater / -negativesNow;
    if (!(ratio > 0)) throw new Error('#NUM!'); // a reinvestment rate of -100% leaves nothing to compound
    return finite(Math.pow(ratio, 1 / (n - 1)) - 1);
  },

  SLN: (args) => {
    checkArity(args, 3);
    return liftScalar(args, ([cost, salvage, life]) => {
      if (toNum(life) === 0) throw new Error('#DIV/0!');
      return finite((toNum(cost) - toNum(salvage)) / toNum(life));
    });
  },

  SYD: (args) => {
    checkArity(args, 4);
    return liftScalar(args, ([costArg, salvageArg, lifeArg, perArg]) => {
      const [cost, salvage, life, per] = [toNum(costArg), toNum(salvageArg), toNum(lifeArg), toNum(perArg)];
      if (life <= 0 || per <= 0 || per > life) throw new Error('#NUM!');
      return finite(((cost - salvage) * (life - per + 1) * 2) / (life * (life + 1)));
    });
  },

  DDB: (args) => {
    checkArity(args, 4, 5);
    return liftScalar(args, ([costArg, salvageArg, lifeArg, periodArg, factorArg]) => {
      const [cost, salvage, life, period] = [toNum(costArg), toNum(salvageArg), toNum(lifeArg), toNum(periodArg)];
      const factor = factorArg === undefined ? 2 : toNum(factorArg);
      if (cost < 0 || salvage < 0 || salvage > cost || life <= 0 || period <= 0 || period > life || factor <= 0) {
        throw new Error('#NUM!');
      }
      // The book value falls by `factor / life` of itself each period, but never below the salvage value: the value
      // after k periods is max(cost * (1 - rate)^k, salvage), and a period depreciates the difference of two of them.
      const rate = Math.min(factor / life, 1);
      const bookValue = (k: number): number => Math.max(cost * Math.pow(1 - rate, k), salvage);
      return finite(bookValue(Math.max(period - 1, 0)) - bookValue(period));
    });
  },

  EFFECT: (args) => {
    checkArity(args, 2);
    return liftScalar(args, ([nominalArg, nperyArg]) => {
      const [nominal, npery] = [toNum(nominalArg), Math.trunc(toNum(nperyArg))];
      if (nominal <= 0 || npery < 1) throw new Error('#NUM!');
      return finite(Math.pow(1 + nominal / npery, npery) - 1);
    });
  },

  NOMINAL: (args) => {
    checkArity(args, 2);
    return liftScalar(args, ([effectArg, nperyArg]) => {
      const [effect, npery] = [toNum(effectArg), Math.trunc(toNum(nperyArg))];
      if (effect <= 0 || npery < 1) throw new Error('#NUM!');
      return finite(npery * (Math.pow(1 + effect, 1 / npery) - 1));
    });
  },

  RRI: (args) => {
    checkArity(args, 3);
    return liftScalar(args, ([nperArg, pvArg, fvArg]) => {
      const [nper, pv, fv] = [toNum(nperArg), toNum(pvArg), toNum(fvArg)];
      if (nper <= 0) throw new Error('#NUM!');
      if (pv === 0) throw new Error('#DIV/0!');
      return finite(Math.pow(fv / pv, 1 / nper) - 1);
    });
  },

  PDURATION: (args) => {
    checkArity(args, 3);
    return liftScalar(args, ([rateArg, pvArg, fvArg]) => {
      const [rate, pv, fv] = [toNum(rateArg), toNum(pvArg), toNum(fvArg)];
      if (rate <= 0 || pv <= 0 || fv <= 0) throw new Error('#NUM!');
      return finite((Math.log(fv) - Math.log(pv)) / Math.log(1 + rate));
    });
  },

  ISPMT: (args) => {
    checkArity(args, 4);
    return liftScalar(args, ([rateArg, perArg, nperArg, pvArg]) => {
      const [rate, per, nper, pv] = [toNum(rateArg), toNum(perArg), toNum(nperArg), toNum(pvArg)];
      if (nper === 0) throw new Error('#DIV/0!');
      return finite(pv * rate * (per / nper - 1));
    });
  },

  FVSCHEDULE: (args) => {
    checkArity(args, 2);
    const principal = toNum(args[0]);
    let value = principal;
    for (const rate of cashFlows(args[1])) value *= 1 + rate;
    return finite(value);
  },
};
