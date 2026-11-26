import type { EvaluationContext } from '../types';
import { mathFunctions } from './math';

const ctx: EvaluationContext = {
  getCellValue: () => null,
  getRangeValues: () => [],
};

/** Invoke a function from the math registry by name. */
function call(name: string, args: unknown[]): unknown {
  return mathFunctions[name](args, ctx);
}

describe('aggregation', () => {
  it('SUM adds numbers and ignores text/blanks', () => {
    expect(call('SUM', [1, 2, 3])).toBe(6);
    expect(call('SUM', [[[1, 'x', 2, null, 3]]])).toBe(6);
    expect(call('SUM', [])).toBe(0);
  });

  it('AVERAGE ignores blanks and errors on an empty set', () => {
    expect(call('AVERAGE', [[[2, 4, 6]]])).toBe(4);
    expect(call('AVERAGE', [[[2, '', 4]]])).toBe(3);
    expect(() => call('AVERAGE', [[]])).toThrow('#DIV/0!');
  });

  it('COUNT counts numbers only; COUNTA / COUNTBLANK count cells', () => {
    expect(call('COUNT', [[[1, 'a', 2, null, 3]]])).toBe(3);
    expect(call('COUNTA', [[[1, 'a', 2, null, 3]]])).toBe(4);
    expect(call('COUNTBLANK', [[[1, 'a', null, '', 3]]])).toBe(2);
  });

  it('MAX / MIN / PRODUCT', () => {
    expect(call('MAX', [[[3, 9, 1]]])).toBe(9);
    expect(call('MIN', [[[3, 9, 1]]])).toBe(1);
    expect(call('MAX', [[]])).toBe(0);
    expect(call('PRODUCT', [[[2, 3, 4]]])).toBe(24);
  });

  it('SUMPRODUCT multiplies aligned arrays', () => {
    expect(call('SUMPRODUCT', [[[1, 2, 3]], [[4, 5, 6]]])).toBe(32);
    expect(() => call('SUMPRODUCT', [[[1, 2]], [[1, 2, 3]]])).toThrow('#VALUE!');
  });
});

describe('conditional aggregation', () => {
  it('SUMIF with and without a separate sum range', () => {
    expect(call('SUMIF', [[[1], [2], [3], [4]], '>2'])).toBe(7);
    expect(call('SUMIF', [[[1], [2], [3]], '>1', [[10], [20], [30]]])).toBe(50);
  });

  it('COUNTIF supports operators and wildcards', () => {
    expect(call('COUNTIF', [[['a'], ['b'], ['a']], 'a'])).toBe(2);
    expect(call('COUNTIF', [[[1], [5], [9]], '>=5'])).toBe(2);
    expect(call('COUNTIF', [[['apple'], ['apricot'], ['berry']], 'ap*'])).toBe(2);
  });

  it('AVERAGEIF averages the matching rows', () => {
    expect(call('AVERAGEIF', [[[1], [2], [3], [4]], '>2'])).toBe(3.5);
  });

  it('SUMIFS / COUNTIFS / AVERAGEIFS apply every criterion', () => {
    const sumRange = [[10], [20], [30]];
    const a = [[1], [2], [3]];
    const b = [['x'], ['y'], ['x']];
    expect(call('SUMIFS', [sumRange, a, '>1', b, 'x'])).toBe(30);
    expect(call('COUNTIFS', [a, '>1', b, 'x'])).toBe(1);
    expect(call('AVERAGEIFS', [sumRange, a, '>=1'])).toBe(20);
  });

  it('SUBTOTAL dispatches by function code', () => {
    expect(call('SUBTOTAL', [9, [[1], [2], [3]]])).toBe(6);
    expect(call('SUBTOTAL', [1, [[2], [4]]])).toBe(3);
    expect(call('SUBTOTAL', [109, [[5], [5]]])).toBe(10);
  });
});

describe('rounding', () => {
  it('ROUND rounds half away from zero', () => {
    expect(call('ROUND', [2.5, 0])).toBe(3);
    expect(call('ROUND', [-2.5, 0])).toBe(-3);
    expect(call('ROUND', [3.14159, 2])).toBe(3.14);
  });

  it('ROUNDUP / ROUNDDOWN / MROUND', () => {
    expect(call('ROUNDUP', [2.1, 0])).toBe(3);
    expect(call('ROUNDDOWN', [2.9, 0])).toBe(2);
    expect(call('MROUND', [10, 3])).toBe(9);
    expect(() => call('MROUND', [-10, 3])).toThrow('#NUM!');
  });

  it('INT / TRUNC / CEILING / FLOOR', () => {
    expect(call('INT', [-2.5])).toBe(-3);
    expect(call('TRUNC', [-2.5])).toBe(-2);
    expect(call('TRUNC', [3.14159, 2])).toBe(3.14);
    expect(call('CEILING', [2.5, 1])).toBe(3);
    expect(call('FLOOR', [2.9, 1])).toBe(2);
  });
});

describe('arithmetic', () => {
  it('ABS / SIGN', () => {
    expect(call('ABS', [-5])).toBe(5);
    expect(call('SIGN', [-3])).toBe(-1);
    expect(call('SIGN', [0])).toBe(0);
  });

  it('MOD takes the sign of the divisor and rejects a zero divisor', () => {
    expect(call('MOD', [10, 3])).toBe(1);
    expect(call('MOD', [-10, 3])).toBe(2);
    expect(() => call('MOD', [10, 0])).toThrow('#DIV/0!');
  });

  it('POWER / SQRT and their domain errors', () => {
    expect(call('POWER', [2, 10])).toBe(1024);
    expect(call('SQRT', [16])).toBe(4);
    expect(() => call('SQRT', [-1])).toThrow('#NUM!');
    expect(() => call('POWER', [0, -1])).toThrow('#NUM!');
  });

  it('EXP / LN / LOG / LOG10', () => {
    expect(call('LN', [Math.E])).toBeCloseTo(1);
    expect(call('LOG', [8, 2])).toBeCloseTo(3);
    expect(call('LOG', [100])).toBeCloseTo(2);
    expect(call('LOG10', [1000])).toBeCloseTo(3);
    expect(() => call('LN', [0])).toThrow('#NUM!');
  });

  it('GCD / LCM', () => {
    expect(call('GCD', [12, 18])).toBe(6);
    expect(call('LCM', [4, 6])).toBe(12);
  });

  it('RAND / RANDBETWEEN stay within range', () => {
    const r = call('RAND', []) as number;
    expect(r).toBeGreaterThanOrEqual(0);
    expect(r).toBeLessThan(1);
    expect(call('RANDBETWEEN', [5, 5])).toBe(5);
    const n = call('RANDBETWEEN', [1, 10]) as number;
    expect(n).toBeGreaterThanOrEqual(1);
    expect(n).toBeLessThanOrEqual(10);
    expect(Number.isInteger(n)).toBe(true);
  });
});

describe('aggregates over very large ranges (a whole column)', () => {
  // 300,000 values: spreading that many arguments into Math.max / Array.push overflows the stack.
  const column: number[][] = Array.from({ length: 300_000 }, (_, i) => [i % 1000]);

  it('MAX, MIN and SUBTOTAL(4/5) do not overflow the call stack', () => {
    expect(call('MAX', [column])).toBe(999);
    expect(call('MIN', [column])).toBe(0);
    expect(call('SUBTOTAL', [4, column])).toBe(999);
    expect(call('SUBTOTAL', [5, column])).toBe(0);
    expect(call('SUM', [column])).toBe(300 * (999 * 1000) / 2);
  });

  it('MIN / MAX / SUBTOTAL return 0 for no numbers', () => {
    expect(call('MIN', [[['a']]])).toBe(0);
    expect(call('MAX', [[['a']]])).toBe(0);
    expect(call('SUBTOTAL', [4, [['a']]])).toBe(0);
    expect(call('SUBTOTAL', [5, [['a']]])).toBe(0);
  });

  it('MIN and MAX pick the extreme of mixed values', () => {
    expect(call('MIN', [3, [[1, 2]], -4])).toBe(-4);
    expect(call('MAX', [3, [[1, 2]], -4])).toBe(3);
    expect(call('SUBTOTAL', [5, [[3], [1], [2]]])).toBe(1);
    expect(call('SUBTOTAL', [4, [[3], [1], [2]]])).toBe(3);
  });
});
