import type { EvaluationContext } from '../types';
import { mathExtFunctions } from './math-ext';

const ctx: EvaluationContext = { getCellValue: () => null, getRangeValues: () => [] };

function call(name: string, ...args: unknown[]): unknown {
  return mathExtFunctions[name](args, ctx);
}

function err(name: string, ...args: unknown[]): string {
  try {
    call(name, ...args);
  } catch (e) {
    return (e as Error).message;
  }
  return 'no error';
}

describe('trigonometry', () => {
  it('computes the direct functions', () => {
    expect(call('SIN', 0)).toBe(0);
    expect(call('SIN', Math.PI / 2)).toBe(1);
    expect(call('COS', 0)).toBe(1);
    expect(call('COS', Math.PI)).toBe(-1);
    expect(call('TAN', 0)).toBe(0);
    expect(call('TAN', Math.PI / 4)).toBeCloseTo(1, 12);
    expect(call('ASIN', 1)).toBeCloseTo(Math.PI / 2, 12);
    expect(call('ACOS', -1)).toBeCloseTo(Math.PI, 12);
    expect(call('ATAN', 1)).toBeCloseTo(Math.PI / 4, 12);
  });

  it('ATAN2 takes (x, y) and rejects (0, 0)', () => {
    expect(call('ATAN2', 1, 1)).toBeCloseTo(Math.PI / 4, 12);
    expect(call('ATAN2', -1, 0)).toBeCloseTo(Math.PI, 12);
    expect(call('ATAN2', 0, -1)).toBeCloseTo(-Math.PI / 2, 12);
    expect(err('ATAN2', 0, 0)).toBe('#DIV/0!');
    expect(err('ATAN2', 1)).toBe('#VALUE!');
  });

  it('ASIN / ACOS reject arguments outside [-1, 1]', () => {
    for (const name of ['ASIN', 'ACOS']) {
      expect(err(name, 1.0001)).toBe('#NUM!');
      expect(err(name, -1.0001)).toBe('#NUM!');
    }
    expect(call('ASIN', -1)).toBeCloseTo(-Math.PI / 2, 12);
    expect(call('ACOS', 1)).toBe(0);
  });

  it('computes the hyperbolic functions and their domains', () => {
    expect(call('SINH', 1)).toBeCloseTo(1.1752011936438014, 12);
    expect(call('COSH', 0)).toBe(1);
    expect(call('TANH', 100)).toBe(1);
    expect(call('ASINH', -3)).toBeCloseTo(-1.8184464592320668, 12);
    expect(call('ACOSH', 1)).toBe(0);
    expect(err('ACOSH', 0.99)).toBe('#NUM!');
    expect(call('ATANH', 0.5)).toBeCloseTo(0.5493061443340548, 12);
    expect(err('ATANH', 1)).toBe('#NUM!');
    expect(err('ATANH', -1)).toBe('#NUM!');
  });

  it('overflow is #NUM!, and zero results are never negative zero', () => {
    expect(err('COSH', 1000)).toBe('#NUM!');
    expect(err('SINH', -1000)).toBe('#NUM!');
    expect(Object.is(call('SIN', -0), 0)).toBe(true);
  });

  it('PI takes no arguments; RADIANS / DEGREES convert', () => {
    expect(call('PI')).toBe(Math.PI);
    expect(err('PI', 1)).toBe('#VALUE!');
    expect(call('RADIANS', 180)).toBe(Math.PI);
    expect(call('RADIANS', -90)).toBe(-Math.PI / 2);
    expect(call('DEGREES', Math.PI)).toBe(180);
    expect(call('DEGREES', 1)).toBeCloseTo(57.29577951308232, 12);
  });
});

describe('argument coercion', () => {
  it('accepts numeric text, booleans and blanks', () => {
    expect(call('COS', '0')).toBe(1);
    expect(call('COS', ' 0 ')).toBe(1);
    expect(call('COS', true)).toBeCloseTo(Math.cos(1), 12);
    expect(call('COS', null)).toBe(1);
  });

  it('text, empty text and errors fail with their own codes', () => {
    expect(err('SIN', 'abc')).toBe('#VALUE!');
    expect(err('SIN', '')).toBe('#VALUE!');
    expect(err('SIN', '  ')).toBe('#VALUE!');
    expect(err('SIN', '#N/A')).toBe('#N/A');
    expect(err('SIN', '#DIV/0!')).toBe('#DIV/0!');
  });

  it('wrong argument counts are #VALUE!', () => {
    expect(err('SIN')).toBe('#VALUE!');
    expect(err('SIN', 1, 2)).toBe('#VALUE!');
    expect(err('COMBIN', 1)).toBe('#VALUE!');
    expect(err('CEILING.MATH')).toBe('#VALUE!');
    expect(err('CEILING.MATH', 1, 2, 3, 4)).toBe('#VALUE!');
    expect(err('ISEVEN')).toBe('#VALUE!');
    expect(err('ISODD', 1, 2)).toBe('#VALUE!');
  });
});

describe('array lifting', () => {
  it('maps a unary function over a range, element errors stay local', () => {
    expect(call('COS', [[0], ['x'], [Math.PI]])).toEqual([[1], ['#VALUE!'], [-1]]);
    expect(call('ASIN', [[0, 2]])).toEqual([[0, '#NUM!']]);
  });

  it('broadcasts a scalar and a single row or column', () => {
    expect(call('ATAN2', [[1], [-1]], 1)).toEqual([[Math.PI / 4], [(3 * Math.PI) / 4]]);
    expect(call('COMBIN', [[4], [5]], [[1, 2]])).toEqual([
      [4, 6],
      [5, 10],
    ]);
  });

  it('positions without a counterpart are #N/A', () => {
    expect(call('COMBIN', [[3], [4], [5]], [[1], [2]])).toEqual([[3], [6], ['#N/A']]);
    expect(call('COMBIN', [[3, 4, 5]], [[1, 2]])).toEqual([[3, 6, '#N/A']]);
  });
});

describe('FACT and FACTDOUBLE', () => {
  it('FACT truncates and rejects negatives and overflow', () => {
    expect(call('FACT', 0)).toBe(1);
    expect(call('FACT', 5)).toBe(120);
    expect(call('FACT', 3.9)).toBe(6);
    expect(call('FACT', 170)).toBeCloseTo(7.257415615307994e306, -293);
    expect(err('FACT', 171)).toBe('#NUM!');
    expect(err('FACT', -1)).toBe('#NUM!');
    expect(err('FACT', -0.5)).toBe('#NUM!');
  });

  it('FACTDOUBLE steps by two', () => {
    expect(call('FACTDOUBLE', 7)).toBe(105);
    expect(call('FACTDOUBLE', 8)).toBe(384);
    expect(call('FACTDOUBLE', 0)).toBe(1);
    expect(call('FACTDOUBLE', 1)).toBe(1);
    expect(err('FACTDOUBLE', -1)).toBe('#NUM!');
    expect(err('FACTDOUBLE', 400)).toBe('#NUM!');
  });
});

describe('COMBIN and PERMUT', () => {
  it('COMBIN counts combinations', () => {
    expect(call('COMBIN', 5, 2)).toBe(10);
    expect(call('COMBIN', 10, 0)).toBe(1);
    expect(call('COMBIN', 10, 10)).toBe(1);
    expect(call('COMBIN', 52, 5)).toBe(2598960);
    expect(call('COMBIN', 5.9, 2.9)).toBe(10);
    expect(call('COMBIN', 60, 30)).toBe(118264581564861420);
  });

  it('COMBIN rejects negatives, k > n and overflow without looping over huge counts', () => {
    expect(err('COMBIN', 3, 5)).toBe('#NUM!');
    expect(err('COMBIN', -1, 0)).toBe('#NUM!');
    expect(err('COMBIN', 3, -1)).toBe('#NUM!');
    expect(err('COMBIN', 2000, 1000)).toBe('#NUM!');
    expect(err('COMBIN', 1e15, 5e14)).toBe('#NUM!');
  });

  it('PERMUT counts arrangements', () => {
    expect(call('PERMUT', 5, 2)).toBe(20);
    expect(call('PERMUT', 5, 0)).toBe(1);
    expect(call('PERMUT', 10, 10)).toBe(3628800);
    expect(err('PERMUT', 3, 5)).toBe('#NUM!');
    expect(err('PERMUT', -1, 0)).toBe('#NUM!');
    expect(err('PERMUT', 3, -1)).toBe('#NUM!');
    expect(err('PERMUT', 1000, 500)).toBe('#NUM!');
  });
});

describe('SUMSQ', () => {
  it('sums squares of numbers, ignoring text and blanks in ranges', () => {
    expect(call('SUMSQ', 3, 4)).toBe(25);
    expect(call('SUMSQ', [[3], ['x'], [null], [-4]])).toBe(25);
    expect(call('SUMSQ')).toBe(0);
    expect(err('SUMSQ', [[1], ['#DIV/0!']])).toBe('#DIV/0!');
  });
});

describe('EVEN, ODD, ISEVEN, ISODD', () => {
  it('EVEN rounds away from zero to an even integer', () => {
    expect([3, 1.5, -1, 0, 2, -2.1, 2.0001].map((n) => call('EVEN', n))).toEqual([
      4, 2, -2, 0, 2, -4, 4,
    ]);
  });

  it('ODD rounds away from zero to an odd integer', () => {
    expect([2, 1.5, -2, 0, 3, -1.5, 2.5].map((n) => call('ODD', n))).toEqual([
      3, 3, -3, 1, 3, -3, 3,
    ]);
  });

  it('ISEVEN / ISODD truncate, treat blank as 0 and reject booleans and text', () => {
    expect(call('ISEVEN', 4)).toBe(true);
    expect(call('ISEVEN', 3)).toBe(false);
    expect(call('ISEVEN', -2.5)).toBe(true);
    expect(call('ISEVEN', null)).toBe(true);
    expect(call('ISEVEN', '6')).toBe(true);
    expect(call('ISODD', 3)).toBe(true);
    expect(call('ISODD', -3.9)).toBe(true);
    expect(call('ISODD', 4)).toBe(false);
    expect(err('ISEVEN', true)).toBe('#VALUE!');
    expect(err('ISODD', 'x')).toBe('#VALUE!');
    expect(err('ISODD', '#N/A')).toBe('#N/A');
    expect(call('ISODD', [[1], [2]])).toEqual([[true], [false]]);
  });
});

describe('CEILING.MATH and FLOOR.MATH', () => {
  it('default significance is 1', () => {
    expect(call('CEILING.MATH', 6.7)).toBe(7);
    expect(call('FLOOR.MATH', 6.7)).toBe(6);
    expect(call('CEILING.MATH', -6.7)).toBe(-6);
    expect(call('FLOOR.MATH', -6.7)).toBe(-7);
  });

  it('rounds to a multiple of significance (its sign is ignored)', () => {
    expect(call('CEILING.MATH', 24.3, 5)).toBe(25);
    expect(call('FLOOR.MATH', 24.3, 5)).toBe(20);
    expect(call('CEILING.MATH', -5.5, 2)).toBe(-4);
    expect(call('CEILING.MATH', -5.5, -2)).toBe(-4);
    expect(call('FLOOR.MATH', -5.5, 2)).toBe(-6);
    expect(call('FLOOR.MATH', 5.5, -2)).toBe(4);
  });

  it('a non-zero mode rounds negative numbers away from zero for CEILING and toward zero for FLOOR', () => {
    expect(call('CEILING.MATH', -5.5, 2, -1)).toBe(-6);
    expect(call('FLOOR.MATH', -5.5, 2, -1)).toBe(-4);
    expect(call('CEILING.MATH', 5.5, 2, -1)).toBe(6);
    expect(call('FLOOR.MATH', 5.5, 2, 1)).toBe(4);
  });

  it('zero significance or zero number gives 0, never -0', () => {
    expect(call('CEILING.MATH', 5, 0)).toBe(0);
    expect(call('FLOOR.MATH', 5, 0)).toBe(0);
    expect(call('CEILING.MATH', 0, 3)).toBe(0);
    expect(Object.is(call('CEILING.MATH', -0.4), 0)).toBe(true);
  });

  it('absorbs floating-point noise in the quotient', () => {
    expect(call('CEILING.MATH', 0.9, 0.3)).toBeCloseTo(0.9, 12);
    expect(call('FLOOR.MATH', 0.7, 0.1)).toBeCloseTo(0.7, 12);
  });

  it('lifts over arrays and propagates errors', () => {
    expect(call('FLOOR.MATH', [[7], [-7]], 2)).toEqual([[6], [-8]]);
    expect(err('CEILING.MATH', 'x')).toBe('#VALUE!');
    expect(err('FLOOR.MATH', 1, '#NUM!')).toBe('#NUM!');
  });
});
