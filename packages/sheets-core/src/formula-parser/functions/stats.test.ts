import type { EvaluationContext } from '../types';
import { statsFunctions } from './stats';

const ctx: EvaluationContext = {
  getCellValue: () => null,
  getRangeValues: () => [],
};

function call(name: string, args: unknown[]): unknown {
  return statsFunctions[name](args, ctx);
}

describe('central tendency', () => {
  it('MEDIAN handles odd and even counts', () => {
    expect(call('MEDIAN', [[[1, 2, 3]]])).toBe(2);
    expect(call('MEDIAN', [[[1, 2, 3, 4]]])).toBe(2.5);
    expect(() => call('MEDIAN', [[]])).toThrow('#NUM!');
  });

  it('MODE returns the most frequent value, #N/A when all are unique', () => {
    expect(call('MODE', [[[1, 2, 2, 3, 3, 3]]])).toBe(3);
    expect(() => call('MODE', [[[1, 2, 3]]])).toThrow('#N/A');
  });
});

describe('spread', () => {
  // Classic dataset: mean 5, population variance 4.
  const data = [[2, 4, 4, 4, 5, 5, 7, 9]];

  it('VARP / STDEVP use the population formula', () => {
    expect(call('VARP', [data])).toBe(4);
    expect(call('STDEVP', [data])).toBe(2);
  });

  it('VAR / STDEV use the sample formula', () => {
    expect(call('VAR', [data])).toBeCloseTo(32 / 7);
    expect(call('STDEV', [data])).toBeCloseTo(Math.sqrt(32 / 7));
  });

  it('VAR needs at least two values', () => {
    expect(() => call('VAR', [[[5]]])).toThrow('#DIV/0!');
  });
});

describe('order statistics', () => {
  it('PERCENTILE interpolates between points', () => {
    expect(call('PERCENTILE', [[[1, 2, 3, 4]], 0.5])).toBe(2.5);
    expect(call('PERCENTILE', [[[1, 2, 3, 4]], 0])).toBe(1);
    expect(call('PERCENTILE', [[[1, 2, 3, 4]], 0.25])).toBe(1.75);
    expect(() => call('PERCENTILE', [[[1, 2]], 2])).toThrow('#NUM!');
  });

  it('QUARTILE maps quart 0-4 onto percentiles', () => {
    expect(call('QUARTILE', [[[1, 2, 3, 4]], 2])).toBe(2.5);
    expect(call('QUARTILE', [[[1, 2, 3, 4]], 0])).toBe(1);
    expect(call('QUARTILE', [[[1, 2, 3, 4]], 4])).toBe(4);
  });

  it('LARGE / SMALL pick the k-th ranked value', () => {
    expect(call('LARGE', [[[3, 1, 4, 1, 5]], 2])).toBe(4);
    expect(call('SMALL', [[[3, 1, 4, 1, 5]], 2])).toBe(1);
    expect(() => call('LARGE', [[[1, 2]], 9])).toThrow('#NUM!');
    expect(() => call('SMALL', [[[1, 2]], 9])).toThrow('#NUM!');
  });

  it('RANK ranks descending by default, ascending on request', () => {
    expect(call('RANK', [3, [[1, 2, 3, 4]]])).toBe(2);
    expect(call('RANK', [3, [[1, 2, 3, 4]], 1])).toBe(3);
  });
});

describe('correlation', () => {
  it('CORREL returns +1 / -1 for perfectly correlated series', () => {
    expect(call('CORREL', [[[1, 2, 3]], [[1, 2, 3]]])).toBeCloseTo(1);
    expect(call('CORREL', [[[1, 2, 3]], [[3, 2, 1]]])).toBeCloseTo(-1);
  });

  it('COVAR computes the population covariance', () => {
    expect(call('COVAR', [[[1, 2, 3]], [[1, 2, 3]]])).toBeCloseTo(2 / 3);
  });

  it('CORREL errors when the series differ in length', () => {
    expect(() => call('CORREL', [[[1, 2, 3]], [[1, 2]]])).toThrow('#N/A');
  });

  it('CORREL and COVAR: numeric text counts, text pairs drop, degenerate series are #DIV/0!', () => {
    expect(call('CORREL', [[['1', 'x', 3, 2]], [[1, 9, 3, 2]]])).toBeCloseTo(1);
    expect(call('CORREL', [[[NaN, 1, 2]], [[5, 1, 2]]])).toBeCloseTo(1);
    expect(() => call('CORREL', [[['x']], [['y']]])).toThrow('#DIV/0!');
    expect(() => call('CORREL', [[[2, 2, 2]], [[1, 2, 3]]])).toThrow('#DIV/0!');
    expect(() => call('CORREL', [[[1, 2, 3]], [[2, 2, 2]]])).toThrow('#DIV/0!');
    expect(() => call('COVAR', [[['x']], [['y']]])).toThrow('#DIV/0!');
    expect(() => call('CORREL', [[[1, '#N/A']], [[1, 2]]])).toThrow('#N/A');
  });

  it('MODE returns the first of the most frequent values', () => {
    expect(call('MODE', [[[1, 2, 2, 1]]])).toBe(1);
    expect(() => call('MODE', [[[1, 2, 3]]])).toThrow('#N/A');
  });
});
