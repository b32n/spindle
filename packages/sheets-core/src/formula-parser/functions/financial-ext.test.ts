import type { EvaluationContext } from '../types';
import { bracketedRoot, financialExtFunctions } from './financial-ext';

const ctx: EvaluationContext = { getCellValue: () => null, getRangeValues: () => [] };

function call(name: string, ...args: unknown[]): unknown {
  return financialExtFunctions[name](args, ctx);
}

function err(name: string, ...args: unknown[]): string {
  try {
    call(name, ...args);
  } catch (e) {
    return (e as Error).message;
  }
  return 'no error';
}

const flows = [[-1000], [300], [400], [500]];
const dates = [[43831], [43922], [44013], [44197]];

describe('XNPV', () => {
  it('discounts dated cash flows on a 365-day year from the first date', () => {
    expect(call('XNPV', 0.1, flows, dates)).toBeCloseTo(128.8169613944, 8);
    expect(call('XNPV', 0, flows, dates)).toBe(200);
    expect(call('XNPV', -0.5, [-1000, 300], [43831, 43922])).toBeCloseTo(-643.4072011, 6);
    expect(call('XNPV', 0.1, [[-1000, 300]], [[43831, 43922.9]])).toBe(call('XNPV', 0.1, [-1000, 300], [43831, 43922]));
    expect(call('XNPV', 0.1, [-1000], [43831])).toBe(-1000);
    expect(call('XNPV', 0.1, [[-1000], [null]], [[43831], [43832]])).toBe(-1000);
    expect(call('XNPV', 0.1, [-1000, 1100], ['2020-01-01', '2021-01-01'])).toBeCloseTo(-1000 + 1100 / Math.pow(1.1, 366 / 365), 9);
  });

  it('rejects mismatched counts, early dates, a rate of -1 or below, text and wrong arity', () => {
    expect(err('XNPV', 0.1, [-1000, 300, 400], [43831, 43922])).toBe('#NUM!');
    expect(err('XNPV', 0.1, [], [])).toBe('#NUM!');
    expect(err('XNPV', 0.1, [-1000, 300], [43831, 43822])).toBe('#NUM!');
    expect(err('XNPV', 0.1, [-1000, 1100], [-5, 10])).toBe('#NUM!'); // a negative date
    expect(err('XIRR', [-1000, 1100], [-5, 10])).toBe('#NUM!');
    expect(err('XNPV', -1, [-1000, 300], [43831, 43922])).toBe('#NUM!');
    expect(err('XNPV', -2, [-1000, 300], [43831, 43922])).toBe('#NUM!');
    expect(err('XNPV', 0.1, [-1000, 300], [43831, 'x'])).toBe('#VALUE!');
    expect(err('XNPV', 0.1, [-1000, 300], [43831, true])).toBe('#VALUE!');
    expect(err('XNPV', 0.1, [-1000, 'x'], [43831, 43922])).toBe('#VALUE!');
    expect(err('XNPV', 0.1, [-1000, '#N/A'], [43831, 43922])).toBe('#N/A');
    expect(err('XNPV', 'x', [-1000], [43831])).toBe('#VALUE!');
    expect(err('XNPV', 0.1, [-1000, 300])).toBe('#VALUE!');
    expect(err('XNPV', -0.5, [1e308, 1e308], [1, 40000])).toBe('#NUM!'); // overflow
  });
});

describe('XIRR', () => {
  const npvAt = (r: number, v: number[], d: number[]): number => v.reduce((s, x, i) => s + x / Math.pow(1 + r, (d[i] - d[0]) / 365), 0);

  it('finds the rate that nets the cash flows to zero', () => {
    const r = call('XIRR', flows, dates) as number;
    expect(r).toBeCloseTo(0.334386096, 8);
    expect(npvAt(r, [-1000, 300, 400, 500], [43831, 43922, 44013, 44197])).toBeCloseTo(0, 8);
    expect(call('XIRR', flows, dates, 0.5)).toBeCloseTo(0.334386096, 8);
    expect(call('XIRR', flows, dates, undefined)).toBeCloseTo(0.334386096, 8);
    expect(call('XIRR', [-1000, 1100], [43831, 44196])).toBeCloseTo(0.1, 6);
    expect(call('XIRR', [0, -1000, 1500], [43831, 43922, 44196])).toBeCloseTo(0.7162245988, 8);
    expect(call('XIRR', [-1000, 1500], [43831, 43922], -0.5)).toBeCloseTo(4.0851071088, 7);
    expect(call('XIRR', [-1000, 1000], [43831, 43922])).toBeCloseTo(0, 9);
    expect(call('XIRR', [-1000, 1500], ['2020-01-01', '2021-01-01'])).toBeCloseTo(Math.pow(1.5, 365 / 366) - 1, 9);
  });

  it('falls back to a bracketed search when Newton does not converge', () => {
    // two roots, 10% and 20%, a guess on neither side still finds one of them
    const two = [-100, 230, -132];
    const years = [43831, 44196, 44561];
    const r = call('XIRR', two, years, 1e300) as number;
    expect(npvAt(r, two, years)).toBeCloseTo(0, 8);
    const low = call('XIRR', two, years, -0.9) as number;
    expect(npvAt(low, two, years)).toBeCloseTo(0, 8);
    expect(err('XIRR', [-1000, 1], [43831, 43922])).toBe('#NUM!'); // the root is a hair above -100%: no convergence
    const loss = call('XIRR', [-1000, 100], [43831, 44196]) as number; // a loss of 90% in a year
    expect(loss).toBeCloseTo(-0.9, 9);
  });

  it('is #NUM! without both signs, with a bad guess, early dates or mismatched counts', () => {
    expect(err('XIRR', [1000, 300], [43831, 43922])).toBe('#NUM!');
    expect(err('XIRR', [-1000, -300], [43831, 43922])).toBe('#NUM!');
    expect(err('XIRR', [0, 0], [43831, 43922])).toBe('#NUM!');
    expect(err('XIRR', [-1000, 1500], [43831, 43922], -1)).toBe('#NUM!');
    expect(err('XIRR', [-1000, 300], [43831, 43822])).toBe('#NUM!');
    expect(err('XIRR', [-1000, 300, 400], [43831, 43922])).toBe('#NUM!');
    expect(err('XIRR', [-1000, 1], [5, 5])).toBe('#NUM!'); // all on one day: no rate changes anything
    expect(call('XIRR', [-1, 1], [5, 5])).toBe(0.1); // already a root: every rate nets to zero
    expect(err('XIRR', [-1000, 'x'], [43831, 43922])).toBe('#VALUE!');
    expect(err('XIRR', [-1000, 1500], [43831, 'x'])).toBe('#VALUE!');
    expect(err('XIRR', [-1000])).toBe('#VALUE!');
    expect(err('XIRR', [-1000, 1500], [43831, 43922], 0.1, 1)).toBe('#VALUE!');
  });

  it('recovers the rate of any simple loan, whatever the guess', () => {
    for (const rate of [-0.9, -0.5, -0.1, 0.01, 0.05, 0.3, 1, 5]) {
      const d = [40000, 40400, 41000, 42000];
      const v = [-1000, 100, 100, 0];
      v[3] = -v.slice(0, 3).reduce((s, x, i) => s + x * Math.pow(1 + rate, (d[3] - d[i]) / 365), 0);
      for (const guess of [undefined, -0.5, 0.5, 3]) {
        const r = call('XIRR', v, d, guess) as number;
        expect(Math.abs(npvAt(r, v, d))).toBeLessThan(1e-6 * Math.max(1, Math.abs(v[3])));
      }
    }
  });
});

describe('bracketedRoot', () => {
  const cubic = (r: number): number => (r - 0.1) * (r - 0.2) * (r - 3);

  it('finds the root nearest the guess', () => {
    expect(bracketedRoot(cubic, 2)).toBeCloseTo(3, 9);
    expect(bracketedRoot(cubic, 0.19)).toBeCloseTo(0.2, 9);
    expect(bracketedRoot(cubic, -5)).toBeCloseTo(0.1, 9);
    expect(bracketedRoot((r) => r - 1e6, 0)).toBeCloseTo(1e6, 3);
    expect(bracketedRoot((r) => r + 0.999, 0)).toBeCloseTo(-0.999, 9);
  });

  it('compares signs, so tiny values do not hide a sign change, and takes an exact zero on the grid', () => {
    expect(bracketedRoot((r) => (r < 1 ? -1e-200 : 1e-200), 0)).toBeCloseTo(1, 6); // the product 1e-400 would underflow
    const lo = Math.log(1e-9);
    const hi = Math.log(1e10);
    const zero = Math.exp(lo + ((hi - lo) * 100) / 600) - 1;
    expect(bracketedRoot((r) => r - zero, 0)).toBe(zero);
  });

  it('gives up when the function does not change sign, or is not finite', () => {
    expect(bracketedRoot(() => 1, 0)).toBeUndefined();
    expect(bracketedRoot(() => NaN, 0)).toBeUndefined();
    expect(bracketedRoot((r) => (r > 0 ? 1 : Infinity), 0)).toBeUndefined();
  });
});

describe('MIRR', () => {
  it('computes the modified internal rate of return', () => {
    expect(call('MIRR', [[-100], [50], [60], [70]], 0.1, 0.12)).toBeCloseTo(0.259753038, 8);
    expect(call('MIRR', [-120000, 39000, 30000, 21000, 37000, 46000], 0.1, 0.12)).toBeCloseTo(0.126094, 6);
    expect(call('MIRR', [-100, 'x', 60, null, 70], 0.1, 0.12)).toBe(call('MIRR', [-100, 60, 70], 0.1, 0.12));
    expect(call('MIRR', [-100, -50, 200, -30, 90], 0.08, 0.15)).toBeGreaterThan(0);
  });

  it('is #DIV/0! without both signs or enough values, and #NUM! when the ratio vanishes', () => {
    expect(err('MIRR', [100, 50], 0.1, 0.12)).toBe('#DIV/0!');
    expect(err('MIRR', [-100, -50], 0.1, 0.12)).toBe('#DIV/0!');
    expect(err('MIRR', [-100], 0.1, 0.12)).toBe('#DIV/0!');
    expect(err('MIRR', [100, -100], -1, 0.12)).toBe('#DIV/0!'); // dividing by (1 + rate)^1 = 0
    expect(err('MIRR', [100, 50, -50], 1e300, 0.1)).toBe('#DIV/0!'); // the discounted outflow underflows to 0
    expect(err('MIRR', [-100, 50, 0], 0.1, -1)).toBe('#NUM!');
    expect(err('MIRR', [-100, 50], 0.1)).toBe('#VALUE!');
    expect(err('MIRR', [-100, '#N/A'], 0.1, 0.1)).toBe('#N/A');
  });
});

describe('depreciation', () => {
  it('SLN divides the depreciable cost by the life', () => {
    expect(call('SLN', 30000, 7500, 10)).toBe(2250);
    expect(call('SLN', 30000, 0, 5)).toBe(6000);
    expect(call('SLN', 1000, 2000, 4)).toBe(-250);
    expect(err('SLN', 30000, 7500, 0)).toBe('#DIV/0!');
    expect(err('SLN', 'x', 1, 2)).toBe('#VALUE!');
    expect(err('SLN', 1, 2)).toBe('#VALUE!');
    expect(call('SLN', [[1000], [600]], 100, 5)).toEqual([[180], [100]]);
  });

  it('SYD weights each period by the years left', () => {
    expect(call('SYD', 30000, 7500, 10, 1)).toBeCloseTo(4090.909090909, 8);
    expect(call('SYD', 30000, 7500, 10, 10)).toBeCloseTo(409.0909090909, 8);
    expect(call('SYD', 30000, 7500, 10, 5)).toBeCloseTo(2454.5454545, 6);
    let total = 0;
    for (let p = 1; p <= 10; p++) total += call('SYD', 30000, 7500, 10, p) as number;
    expect(total).toBeCloseTo(22500, 8);
    expect(err('SYD', 30000, 7500, 10, 11)).toBe('#NUM!');
    expect(err('SYD', 30000, 7500, 10, 0)).toBe('#NUM!');
    expect(err('SYD', 30000, 7500, 0, 1)).toBe('#NUM!');
    expect(err('SYD', 1, 2, 3)).toBe('#VALUE!');
  });

  it('DDB depreciates a fraction of the book value, never below the salvage value', () => {
    expect(call('DDB', 2400, 300, 10, 1)).toBeCloseTo(480, 9);
    expect(call('DDB', 2400, 300, 10, 2, 1.5)).toBeCloseTo(306, 9);
    expect(call('DDB', 2400, 300, 10, 10)).toBeCloseTo(22.1225472, 7);
    expect(call('DDB', 100, 10, 2, 1, 5)).toBe(90);
    expect(call('DDB', 100, 10, 2, 2, 5)).toBe(0);
    expect(call('DDB', 2400, 0, 10, 7)).toBeCloseTo(2400 * 0.8 ** 6 * 0.2, 9);
    expect(call('DDB', 2400, 300, 10, 1.5)).toBeCloseTo(2400 * (0.8 ** 0.5 - 0.8 ** 1.5), 9);
    expect(call('DDB', 2400, 300, 10, 0.5)).toBeCloseTo(2400 - 2400 * 0.8 ** 0.5, 9);
  });

  it('DDB over the whole life adds up to the depreciable cost', () => {
    for (const [cost, salvage, life, factor] of [[2400, 300, 10, 2], [1000, 0, 7, 1.5], [5000, 5000, 4, 2], [900, 100, 3, 9]]) {
      let total = 0;
      for (let p = 1; p <= life; p++) total += call('DDB', cost, salvage, life, p, factor) as number;
      expect(total).toBeLessThanOrEqual(cost - salvage + 1e-9);
      if (factor === 9) expect(total).toBeCloseTo(cost - salvage, 9);
    }
  });

  it('DDB is #NUM! for arguments out of range', () => {
    expect(err('DDB', 2400, 300, 10, 11)).toBe('#NUM!');
    expect(err('DDB', 2400, 300, 10, 0)).toBe('#NUM!');
    expect(err('DDB', -1, 300, 10, 1)).toBe('#NUM!');
    expect(err('DDB', 2400, -1, 10, 1)).toBe('#NUM!');
    expect(err('DDB', 2400, 3000, 10, 1)).toBe('#NUM!');
    expect(err('DDB', 2400, 300, 0, 1)).toBe('#NUM!');
    expect(err('DDB', 2400, 300, 10, 1, 0)).toBe('#NUM!');
    expect(err('DDB', 2400, 300, 10)).toBe('#VALUE!');
    expect(call('DDB', [[2400], [1200]], 300, 10, 1)).toEqual([[480], [240]]);
  });
});

describe('rates', () => {
  it('EFFECT and NOMINAL convert between nominal and effective rates', () => {
    expect(call('EFFECT', 0.1, 4)).toBeCloseTo(0.103812890625, 12);
    expect(call('EFFECT', 0.0525, 4)).toBeCloseTo(0.053542667, 8);
    expect(call('EFFECT', 0.1, 4.9)).toBeCloseTo(0.103812890625, 12);
    expect(call('NOMINAL', 0.103812890625, 4)).toBeCloseTo(0.1, 12);
    expect(call('NOMINAL', 0.053543, 4)).toBeCloseTo(0.0525, 5);
    expect(call('NOMINAL', call('EFFECT', 0.07, 12), 12)).toBeCloseTo(0.07, 12);
    expect(err('EFFECT', 0.1, 0)).toBe('#NUM!');
    expect(err('EFFECT', 0, 4)).toBe('#NUM!');
    expect(err('EFFECT', -0.1, 4)).toBe('#NUM!');
    expect(err('NOMINAL', 0.1, 0.5)).toBe('#NUM!');
    expect(err('NOMINAL', 0, 4)).toBe('#NUM!');
    expect(err('EFFECT', 0.1)).toBe('#VALUE!');
    const lifted = call('EFFECT', [[0.1], [0.2]], 2) as number[][];
    expect(lifted[0][0]).toBeCloseTo(0.1025, 12);
    expect(lifted[1][0]).toBeCloseTo(0.21, 12);
  });

  it('RRI and PDURATION give the rate and the time of a growth', () => {
    expect(call('RRI', 10, 100, 200)).toBeCloseTo(0.0717734625, 9);
    expect(call('RRI', 5, 200, 100)).toBeCloseTo(Math.pow(0.5, 0.2) - 1, 12);
    expect(call('PDURATION', 0.05, 100, 200)).toBeCloseTo(14.2066990829, 8);
    expect(call('PDURATION', 0.05, 200, 100)).toBeCloseTo(-14.2066990829, 8);
    expect(err('RRI', 0, 100, 200)).toBe('#NUM!');
    expect(err('RRI', 10, -100, 200)).toBe('#NUM!');
    expect(err('RRI', 10, 0, 200)).toBe('#DIV/0!');
    expect(err('PDURATION', 0, 100, 200)).toBe('#NUM!');
    expect(err('PDURATION', 0.05, -100, 200)).toBe('#NUM!');
    expect(err('PDURATION', 0.05, 100, 0)).toBe('#NUM!');
    expect(err('RRI', 1, 2)).toBe('#VALUE!');
  });

  it('ISPMT is the interest of a period of a loan with even principal payments', () => {
    expect(call('ISPMT', 0.1, 1, 4, 1000)).toBeCloseTo(-75, 12);
    expect(call('ISPMT', 0.1, 0, 4, 1000)).toBeCloseTo(-100, 12);
    expect(call('ISPMT', 0.1, 4, 4, 1000)).toBe(0);
    expect(err('ISPMT', 0.1, 2, 0, 1000)).toBe('#DIV/0!');
    expect(err('ISPMT', 0.1, 1, 3, 'x')).toBe('#VALUE!');
    expect(err('ISPMT', 0.1, 1, 3)).toBe('#VALUE!');
  });

  it('FVSCHEDULE compounds a principal through a list of rates', () => {
    expect(call('FVSCHEDULE', 1000, [0.09, 0.11, 0.1])).toBeCloseTo(1330.89, 9);
    expect(call('FVSCHEDULE', 1, [[0.09], [0.11], [0.1]])).toBeCloseTo(1.33089, 12);
    expect(call('FVSCHEDULE', 1000, [[0]])).toBe(1000);
    expect(call('FVSCHEDULE', 1000, [[0.5, null]])).toBe(1500);
    expect(err('FVSCHEDULE', 500, [0.5, '#N/A'])).toBe('#N/A');
    expect(err('FVSCHEDULE', 500, [0.5, 'x'])).toBe('#VALUE!');
    expect(err('FVSCHEDULE', 500, [0.5, true])).toBe('#VALUE!');
    expect(err('FVSCHEDULE', 1e300, [1e10, 1e10])).toBe('#NUM!');
    expect(err('FVSCHEDULE', 500)).toBe('#VALUE!');
  });
});
