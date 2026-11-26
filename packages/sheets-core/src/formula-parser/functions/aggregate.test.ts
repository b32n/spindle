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

const col = (...v: unknown[]): unknown[][] => v.map((x) => [x]);

describe('MAXIFS / MINIFS', () => {
  const values = col(10, 25, 7, 40, 15);
  const tags = col('x', 'y', 'x', 'y', 'x');
  const nums = col(1, 2, 3, 4, 5);

  it('pick the extreme of the matching rows', () => {
    expect(call('MAXIFS', values, tags, 'x')).toBe(15);
    expect(call('MINIFS', values, tags, 'x')).toBe(7);
    expect(call('MAXIFS', values, tags, 'x', nums, '>3')).toBe(15);
    expect(call('MINIFS', values, tags, 'y', nums, '<>2')).toBe(40);
    expect(call('MAXIFS', values, nums, '>=2', nums, '<=4')).toBe(40);
    expect(call('MAXIFS', values, tags, '?')).toBe(40);
  });

  it('give 0 when nothing matches', () => {
    expect(call('MAXIFS', values, tags, 'z')).toBe(0);
    expect(call('MINIFS', values, tags, 'z')).toBe(0);
  });

  it('only numbers count: text, booleans and blanks are skipped, negatives work', () => {
    const odd = col('t', -3, true, null, -8);
    const all = col('x', 'x', 'x', 'x', 'x');
    expect(call('MAXIFS', odd, all, 'x')).toBe(-3);
    expect(call('MINIFS', odd, all, 'x')).toBe(-8);
    expect(call('MAXIFS', col('t', null), col('x', 'x'), 'x')).toBe(0);
  });

  it('a matching error value propagates; a non-matching one does not', () => {
    const withErr = col(5, '#N/A', 9);
    expect(err('MAXIFS', withErr, col('x', 'x', 'x'), 'x')).toBe('#N/A');
    expect(call('MAXIFS', withErr, col('x', 'y', 'x'), 'x')).toBe(9);
  });

  it('need ranges of the same shape, not just the same cell count', () => {
    const grid = [
      ['x', 'x'],
      ['x', 'x'],
    ];
    expect(err('MAXIFS', col(1, 2, 3, 4), grid, 'x')).toBe('#VALUE!');
    expect(
      call(
        'MAXIFS',
        grid.map((r) => r.map(() => 5)),
        grid,
        'x'
      )
    ).toBe(5);
    expect(err('MAXIFS', [[1, 2, 3, 4]], col('x', 'x', 'x', 'x'), 'x')).toBe('#VALUE!');
  });

  it('need matching sizes and complete (range, criterion) pairs', () => {
    expect(err('MAXIFS', values, col('x', 'y'), 'x')).toBe('#VALUE!');
    expect(err('MAXIFS', values, tags)).toBe('#VALUE!');
    expect(err('MAXIFS', values)).toBe('#VALUE!');
    expect(err('MINIFS', values, tags, 'x', nums)).toBe('#VALUE!');
  });
});

describe('large inputs', () => {
  it('AGGREGATE MAX / MIN do not overflow the stack on a long column', () => {
    const big = Array.from({ length: 300000 }, (_, i) => [i % 1000]);
    expect(call('AGGREGATE', 4, 6, big)).toBe(999);
    expect(call('AGGREGATE', 5, 6, big)).toBe(0);
    expect(call('AGGREGATE', 5, 6, col(9, 2, 5))).toBe(2);
  });
});

describe('AGGREGATE', () => {
  const data = col(3, 7, '#DIV/0!', 10, 4, 'x', null, 7);

  it('reference form, functions 1-13 with errors ignored (option 6)', () => {
    const expected: Record<number, number> = {
      1: 31 / 5,
      2: 5,
      3: 6,
      4: 10,
      5: 3,
      6: 5880,
      9: 31,
      12: 7,
      13: 7,
    };
    for (const [fn, value] of Object.entries(expected)) {
      expect(call('AGGREGATE', Number(fn), 6, data)).toBeCloseTo(value, 12);
    }
    expect(call('AGGREGATE', 7, 6, data)).toBeCloseTo(
      call('STDEV.S', col(3, 7, 10, 4, 7)) as number,
      12
    );
    expect(call('AGGREGATE', 8, 6, data)).toBeCloseTo(
      call('STDEV.P', col(3, 7, 10, 4, 7)) as number,
      12
    );
    expect(call('AGGREGATE', 10, 6, data)).toBeCloseTo(
      call('VAR.S', col(3, 7, 10, 4, 7)) as number,
      12
    );
    expect(call('AGGREGATE', 11, 6, data)).toBeCloseTo(
      call('VAR.P', col(3, 7, 10, 4, 7)) as number,
      12
    );
  });

  it('options 2, 3, 6 and 7 skip errors; 0, 1, 4 and 5 propagate them', () => {
    for (const option of [2, 3, 6, 7]) expect(call('AGGREGATE', 9, option, data)).toBe(31);
    for (const option of [0, 1, 4, 5]) expect(err('AGGREGATE', 9, option, data)).toBe('#DIV/0!');
    // COUNT skips an error cell and COUNTA counts it, whatever the option.
    for (const option of [0, 1, 4, 5]) {
      expect(call('AGGREGATE', 2, option, data)).toBe(5);
      expect(call('AGGREGATE', 3, option, data)).toBe(7);
    }
    expect(call('AGGREGATE', 9, 4, col(1, 2))).toBe(3);
  });

  it('pools several references', () => {
    expect(call('AGGREGATE', 9, 6, col(3, 7), col(10, 4))).toBe(24);
    expect(call('AGGREGATE', 4, 6, col(3, 7), col(10, 4), 20)).toBe(20);
  });

  it('empty input: AVERAGE #DIV/0!, MAX / MIN / PRODUCT 0, MEDIAN #NUM!, MODE #N/A', () => {
    const none = col('x', null, true);
    expect(err('AGGREGATE', 1, 4, none)).toBe('#DIV/0!');
    expect(call('AGGREGATE', 4, 4, none)).toBe(0);
    expect(call('AGGREGATE', 5, 4, none)).toBe(0);
    expect(call('AGGREGATE', 6, 4, none)).toBe(0);
    expect(call('AGGREGATE', 9, 4, none)).toBe(0);
    expect(call('AGGREGATE', 2, 4, none)).toBe(0);
    expect(call('AGGREGATE', 3, 4, none)).toBe(2);
    expect(err('AGGREGATE', 12, 4, none)).toBe('#NUM!');
    expect(err('AGGREGATE', 13, 4, none)).toBe('#N/A');
    expect(err('AGGREGATE', 7, 4, col(5))).toBe('#DIV/0!');
  });

  it('array form, functions 14-19', () => {
    expect(call('AGGREGATE', 14, 6, data, 2)).toBe(7);
    expect(call('AGGREGATE', 14, 6, data, 1)).toBe(10);
    expect(call('AGGREGATE', 15, 6, data, 1)).toBe(3);
    expect(call('AGGREGATE', 16, 6, data, 0.5)).toBe(7);
    expect(call('AGGREGATE', 17, 6, data, 3)).toBe(7);
    expect(call('AGGREGATE', 18, 6, data, 0.5)).toBe(7);
    expect(call('AGGREGATE', 19, 6, data, 1)).toBe(3.5);
  });

  it('array form over a computed array with error elements', () => {
    const computed = [[1 / 1], ['#DIV/0!'], [5]];
    expect(call('AGGREGATE', 14, 6, computed, 1)).toBe(5);
    expect(err('AGGREGATE', 14, 0, computed, 1)).toBe('#DIV/0!');
  });

  it('array form argument errors are #NUM!', () => {
    expect(err('AGGREGATE', 14, 6, data, 0)).toBe('#NUM!');
    expect(err('AGGREGATE', 14, 6, data, 9)).toBe('#NUM!');
    expect(err('AGGREGATE', 15, 6, data, 0)).toBe('#NUM!');
    expect(err('AGGREGATE', 16, 6, data, 1.5)).toBe('#NUM!');
    expect(err('AGGREGATE', 16, 6, col('x'), 0.5)).toBe('#NUM!');
    expect(err('AGGREGATE', 17, 6, data, 5)).toBe('#NUM!');
    expect(err('AGGREGATE', 18, 6, data, 0.01)).toBe('#NUM!');
    expect(err('AGGREGATE', 19, 6, data, 4)).toBe('#NUM!');
  });

  it('validates the function number, the option and the argument count', () => {
    expect(err('AGGREGATE', 0, 6, data)).toBe('#VALUE!');
    expect(err('AGGREGATE', 20, 6, data)).toBe('#VALUE!');
    expect(err('AGGREGATE', 9, 8, data)).toBe('#VALUE!');
    expect(err('AGGREGATE', 9, -1, data)).toBe('#VALUE!');
    expect(err('AGGREGATE', 9, 6)).toBe('#VALUE!');
    expect(err('AGGREGATE', 14, 6, data)).toBe('#VALUE!');
    expect(err('AGGREGATE', 14, 6, data, 1, 2)).toBe('#VALUE!');
    expect(err('AGGREGATE')).toBe('#VALUE!');
    expect(err('AGGREGATE', 'x', 6, data)).toBe('#VALUE!');
    expect(err('AGGREGATE', '#N/A', 6, data)).toBe('#N/A');
    expect(call('AGGREGATE', 9.9, 6.9, data)).toBe(31);
  });
});
