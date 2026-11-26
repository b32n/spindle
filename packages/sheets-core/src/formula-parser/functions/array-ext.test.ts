import type { EvaluationContext } from '../types';
import { arrayExtFunctions } from './array-ext';

const ctx: EvaluationContext = { getCellValue: () => null, getRangeValues: () => [] };

function call(name: string, ...args: unknown[]): unknown {
  return arrayExtFunctions[name](args, ctx);
}

function err(name: string, ...args: unknown[]): string {
  try {
    call(name, ...args);
  } catch (e) {
    return (e as Error).message;
  }
  return 'no error';
}

describe('TRANSPOSE', () => {
  it('swaps rows and columns, for tables, vectors and scalars', () => {
    expect(call('TRANSPOSE', [[1, 2, 3], [4, 5, 6]])).toEqual([
      [1, 4],
      [2, 5],
      [3, 6],
    ]);
    expect(call('TRANSPOSE', [[1], [2], [3]])).toEqual([[1, 2, 3]]);
    expect(call('TRANSPOSE', [1, 2, 3])).toEqual([[1], [2], [3]]);
    expect(call('TRANSPOSE', 5)).toEqual([[5]]);
    expect(call('TRANSPOSE', [])).toEqual([]);
  });

  it('turns an empty cell into 0 and keeps errors and logicals', () => {
    expect(call('TRANSPOSE', [[1], [null], [undefined]])).toEqual([[1, 0, 0]]);
    expect(call('TRANSPOSE', [['x', true], ['#N/A', '']])).toEqual([
      ['x', '#N/A'],
      [true, ''],
    ]);
  });

  it('does not change its argument and checks the argument count', () => {
    const m = [[1, 2]];
    call('TRANSPOSE', m);
    expect(m).toEqual([[1, 2]]);
    expect(err('TRANSPOSE')).toBe('#VALUE!');
    expect(err('TRANSPOSE', 1, 2)).toBe('#VALUE!');
  });
});

describe('large arrays and properties', () => {
  const grid = (rows: number, cols: number): number[][] =>
    Array.from({ length: rows }, (_, r) => Array.from({ length: cols }, (_, c) => r * cols + c));

  it('stays linear and never overflows the stack on a million cells', () => {
    const started = Date.now();
    expect((call('TRANSPOSE', grid(1000, 1000)) as unknown[][]).length).toBe(1000);
    expect(Date.now() - started).toBeLessThan(5000);
  });

  it('TRANSPOSE twice is the identity', () => {
    for (const [rows, cols] of [[1, 1], [1, 7], [7, 1], [4, 6], [13, 3]]) {
      const m = grid(rows, cols);
      expect(call('TRANSPOSE', call('TRANSPOSE', m))).toEqual(m);
    }
  });
});
