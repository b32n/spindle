import type { EvaluationContext } from '../types';
import { eagerFunctions } from './index';

const ctx: EvaluationContext = { getCellValue: () => null, getRangeValues: () => [] };

function call(name: string, ...args: unknown[]): unknown {
  return eagerFunctions[name](args, ctx);
}

function err(name: string, ...args: unknown[]): string {
  try {
    call(name, ...args);
  } catch (e) {
    return (e as Error).message;
  }
  return 'no error';
}

/** A one-column range. */
const col = (...v: unknown[]): unknown[][] => v.map((x) => [x]);

const DATA = col(1, 2, 2, 3, 10);

describe('dotted names', () => {
  it('are the same implementations as the classic names', () => {
    expect(call('STDEV.S', DATA)).toBe(call('STDEV', DATA));
    expect(call('STDEV.P', DATA)).toBe(call('STDEVP', DATA));
    expect(call('VAR.S', DATA)).toBe(call('VAR', DATA));
    expect(call('VAR.P', DATA)).toBe(call('VARP', DATA));
    expect(call('PERCENTILE.INC', DATA, 0.9)).toBe(call('PERCENTILE', DATA, 0.9));
    expect(call('QUARTILE.INC', DATA, 1)).toBe(call('QUARTILE', DATA, 1));
    expect(call('RANK.EQ', 2, DATA)).toBe(call('RANK', 2, DATA));
    expect(call('MODE.SNGL', DATA)).toBe(call('MODE', DATA));
  });

  it('compute the expected statistics', () => {
    expect(call('STDEV.S', DATA)).toBeCloseTo(3.646916505762094, 12);
    expect(call('STDEV.P', DATA)).toBeCloseTo(3.261901286060018, 12);
    expect(call('VAR.S', DATA)).toBeCloseTo(13.3, 12);
    expect(call('VAR.P', DATA)).toBeCloseTo(10.64, 12);
    expect(call('RANK.EQ', 2, DATA)).toBe(3);
    expect(call('RANK.EQ', 2, DATA, 1)).toBe(2);
    expect(call('MODE.SNGL', DATA)).toBe(2);
  });

  it('report #DIV/0!, #NUM! and #N/A like Excel', () => {
    expect(err('STDEV.S', col(5))).toBe('#DIV/0!');
    expect(call('STDEV.P', col(5))).toBe(0);
    expect(err('VAR.P', col('x'))).toBe('#DIV/0!');
    expect(err('PERCENTILE.INC', DATA, 1.1)).toBe('#NUM!');
    expect(err('QUARTILE.INC', DATA, 5)).toBe('#NUM!');
    expect(err('RANK.EQ', 99, DATA)).toBe('#N/A');
    expect(err('MODE.SNGL', col(1, 2, 3))).toBe('#N/A');
    expect(err('STDEV.S', col(1, '#DIV/0!', 3))).toBe('#DIV/0!');
  });
});

describe('MODE tie-breaking', () => {
  it('returns the first value that appears among the most frequent', () => {
    expect(call('MODE.SNGL', col(1, 2, 2, 1, 3, 3))).toBe(1);
    expect(call('MODE', col(3, 3, 1, 1))).toBe(3);
    expect(call('MODE', col(5, 1, 1, 5, 1))).toBe(1);
  });
});

describe('PERCENTILE.EXC and QUARTILE.EXC', () => {
  it('interpolate on rank k*(n+1)', () => {
    expect(call('PERCENTILE.EXC', DATA, 0.5)).toBe(2);
    expect(call('PERCENTILE.EXC', DATA, 0.25)).toBe(1.5);
    expect(call('PERCENTILE.EXC', DATA, 5 / 6)).toBe(10);
    expect(call('QUARTILE.EXC', DATA, 1)).toBe(1.5);
    expect(call('QUARTILE.EXC', DATA, 3)).toBe(6.5);
    expect(call('QUARTILE.EXC', DATA, 2.9)).toBe(2);
  });

  it('reject k outside (1/(n+1), n/(n+1)), quartiles outside 1-3 and empty data', () => {
    expect(err('PERCENTILE.EXC', DATA, 0.1)).toBe('#NUM!');
    expect(err('PERCENTILE.EXC', DATA, 0.9)).toBe('#NUM!');
    expect(err('PERCENTILE.EXC', col('x'), 0.5)).toBe('#NUM!');
    expect(err('QUARTILE.EXC', DATA, 0)).toBe('#NUM!');
    expect(err('QUARTILE.EXC', DATA, 4)).toBe('#NUM!');
    expect(err('PERCENTILE.EXC', DATA)).toBe('#VALUE!');
    expect(err('QUARTILE.EXC', DATA, 1, 2)).toBe('#VALUE!');
  });
});

describe('GEOMEAN', () => {
  it('is the n-th root of the product', () => {
    expect(call('GEOMEAN', col(2, 8, 4))).toBeCloseTo(4, 12);
    expect(call('GEOMEAN', 1, 10, 100)).toBeCloseTo(10, 12);
    expect(call('GEOMEAN', col(2, 'x', null, 8))).toBeCloseTo(4, 12);
  });

  it('needs positive numbers', () => {
    expect(err('GEOMEAN', col(2, 0, 8))).toBe('#NUM!');
    expect(err('GEOMEAN', 4, -1)).toBe('#NUM!');
    expect(err('GEOMEAN', col('x', null))).toBe('#NUM!');
    expect(err('GEOMEAN', col(2, '#N/A'))).toBe('#N/A');
  });
});

describe('regression', () => {
  const ys = col(2, 4, 5, 4, 5);
  const xs = col(1, 2, 3, 4, 5);

  it('SLOPE, INTERCEPT, RSQ and FORECAST fit the least-squares line', () => {
    expect(call('SLOPE', ys, xs)).toBeCloseTo(0.6, 12);
    expect(call('INTERCEPT', ys, xs)).toBeCloseTo(2.2, 12);
    expect(call('RSQ', ys, xs)).toBeCloseTo(0.6, 12);
    expect(call('FORECAST', 6, ys, xs)).toBeCloseTo(5.8, 12);
    expect(call('FORECAST.LINEAR', 0, ys, xs)).toBeCloseTo(2.2, 12);
    expect(call('FORECAST', '6', ys, xs)).toBeCloseTo(5.8, 12);
  });

  it('drops pairs with a non-numeric side', () => {
    expect(call('SLOPE', col(2, 9, 6), col(1, 'x', 3))).toBeCloseTo(2, 12);
    expect(call('INTERCEPT', col(2, 9, 6), col(1, null, 3))).toBeCloseTo(0, 12);
  });

  it('series of different length are #N/A', () => {
    expect(err('SLOPE', col(1, 2), col(1, 2, 3))).toBe('#N/A');
    expect(err('FORECAST', 1, col(1, 2), col(1, 2, 3))).toBe('#N/A');
  });

  it('too few points, constant x or constant y are #DIV/0!', () => {
    expect(err('SLOPE', col(1), col(1))).toBe('#DIV/0!');
    expect(err('SLOPE', col(1, 2, 3), col(2, 2, 2))).toBe('#DIV/0!');
    expect(err('INTERCEPT', col(1, 2, 3), col(2, 2, 2))).toBe('#DIV/0!');
    expect(err('RSQ', col(5, 5, 5), col(1, 2, 3))).toBe('#DIV/0!');
    expect(call('SLOPE', col(5, 5, 5), col(1, 2, 3))).toBe(0);
    expect(err('FORECAST', 1, col('x'), col('y'))).toBe('#DIV/0!');
  });

  it('errors propagate from either series and from x', () => {
    expect(err('SLOPE', col(1, '#N/A', 3), col(1, 2, 3))).toBe('#N/A');
    expect(err('RSQ', col(1, 2, 3), col(1, '#DIV/0!', 3))).toBe('#DIV/0!');
    expect(err('FORECAST', '#NUM!', ys, xs)).toBe('#NUM!');
    expect(err('FORECAST', 'x', ys, xs)).toBe('#VALUE!');
    expect(err('FORECAST', '', ys, xs)).toBe('#VALUE!');
    expect(err('CORREL', col(1, '#N/A'), col(1, 2))).toBe('#N/A');
  });

  it('check their argument count', () => {
    expect(err('SLOPE', ys)).toBe('#VALUE!');
    expect(err('INTERCEPT', ys, xs, xs)).toBe('#VALUE!');
    expect(err('RSQ', ys)).toBe('#VALUE!');
    expect(err('FORECAST', 1, ys)).toBe('#VALUE!');
  });
});

describe('FREQUENCY', () => {
  const scores = col(79, 85, 78, 85, 50, 81, 95, 88, 97);

  it('counts per bin and above the last bin', () => {
    expect(call('FREQUENCY', scores, col(70, 79, 89))).toEqual([[1], [2], [4], [2]]);
  });

  it('values equal to a bin belong to that bin', () => {
    expect(call('FREQUENCY', col(1, 2, 3), col(2))).toEqual([[2], [1]]);
  });

  it('keeps the order of unsorted bins', () => {
    // bins 8, 2, 5: 6 falls in the first, 1 in the second, 5 in the third, 10 above them all
    expect(call('FREQUENCY', col(1, 10, 5, 6), col(8, 2, 5))).toEqual([[1], [1], [1], [1]]);
  });

  it('ignores text and blanks in data and bins; no bins gives the count', () => {
    expect(call('FREQUENCY', col('x', null, 5, 6), col(5, 7, 100))).toEqual([[1], [1], [0], [0]]);
    expect(call('FREQUENCY', col(1, 2, 3), col('x', null))).toEqual([[3]]);
    expect(call('FREQUENCY', col('x'), col(1, 2))).toEqual([[0], [0], [0]]);
  });

  it('equal bins give the count to the first one', () => {
    expect(call('FREQUENCY', col(5), col(5, 5))).toEqual([[1], [0], [0]]);
  });

  it('propagates errors and checks the argument count', () => {
    expect(err('FREQUENCY', col(1, '#N/A'), col(1))).toBe('#N/A');
    expect(err('FREQUENCY', col(1))).toBe('#VALUE!');
  });
});
