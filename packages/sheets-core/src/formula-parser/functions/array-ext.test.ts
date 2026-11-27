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

const G = [
  [1, 2, 3],
  [4, 5, 6],
  [7, 8, 9],
];

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

describe('CHOOSECOLS and CHOOSEROWS', () => {
  it('pick columns or rows in the order given, negative from the end', () => {
    expect(call('CHOOSECOLS', G, 1, 3)).toEqual([[1, 3], [4, 6], [7, 9]]);
    expect(call('CHOOSECOLS', G, -1)).toEqual([[3], [6], [9]]);
    expect(call('CHOOSECOLS', G, [[2, 2, 1]])).toEqual([[2, 2, 1], [5, 5, 4], [8, 8, 7]]);
    expect(call('CHOOSECOLS', G, -3, -1)).toEqual([[1, 3], [4, 6], [7, 9]]);
    expect(call('CHOOSECOLS', G, 2.9)).toEqual([[2], [5], [8]]);
    expect(call('CHOOSEROWS', G, 1, 3)).toEqual([[1, 2, 3], [7, 8, 9]]);
    expect(call('CHOOSEROWS', G, -1)).toEqual([[7, 8, 9]]);
    expect(call('CHOOSEROWS', G, [[3], [1]])).toEqual([[7, 8, 9], [1, 2, 3]]);
    expect(call('CHOOSECOLS', [[1, null]], 2)).toEqual([[0]]);
    const twice = call('CHOOSEROWS', G, 1, 1) as unknown[][];
    expect(twice[0]).not.toBe(twice[1]); // two copies, not one row array twice
  });

  it('reject 0 and positions past the array, and check arguments', () => {
    expect(err('CHOOSECOLS', G, 0)).toBe('#VALUE!');
    expect(err('CHOOSECOLS', G, 4)).toBe('#VALUE!');
    expect(err('CHOOSECOLS', G, -4)).toBe('#VALUE!');
    expect(err('CHOOSEROWS', G, 0)).toBe('#VALUE!');
    expect(err('CHOOSEROWS', G, -4)).toBe('#VALUE!');
    expect(err('CHOOSEROWS', G, 'x')).toBe('#VALUE!');
    expect(err('CHOOSEROWS', G)).toBe('#VALUE!');
    expect(err('CHOOSECOLS', G)).toBe('#VALUE!');
  });
});

describe('TAKE and DROP', () => {
  it('TAKE keeps rows and columns from the start, or the end when negative', () => {
    expect(call('TAKE', G, 2)).toEqual([[1, 2, 3], [4, 5, 6]]);
    expect(call('TAKE', G, -1)).toEqual([[7, 8, 9]]);
    expect(call('TAKE', G, undefined, 2)).toEqual([[1, 2], [4, 5], [7, 8]]);
    expect(call('TAKE', G, 2, -2)).toEqual([[2, 3], [5, 6]]);
    expect(call('TAKE', G, 10, 10)).toEqual(G);
    expect(call('TAKE', G, -10)).toEqual(G);
    expect(call('TAKE', G, 1.9)).toEqual([[1, 2, 3]]);
    expect(call('TAKE', [1, 2, 3], 1, 2)).toEqual([[1, 2]]);
    expect(call('TAKE', 5, 1)).toEqual([[5]]);
  });

  it('TAKE reports #CALC! for an empty result and #VALUE! with nothing to take', () => {
    expect(err('TAKE', G, 0)).toBe('#CALC!');
    expect(err('TAKE', G, 1, 0)).toBe('#CALC!');
    expect(err('TAKE', G, undefined, undefined)).toBe('#VALUE!');
    expect(err('TAKE', G)).toBe('#VALUE!');
    expect(err('TAKE', G, 1, 1, 1)).toBe('#VALUE!');
    expect(err('TAKE', G, 'x')).toBe('#VALUE!');
  });

  it('DROP removes rows and columns from the start, or the end when negative', () => {
    expect(call('DROP', G, 1)).toEqual([[4, 5, 6], [7, 8, 9]]);
    expect(call('DROP', G, -1)).toEqual([[1, 2, 3], [4, 5, 6]]);
    expect(call('DROP', G, undefined, 1)).toEqual([[2, 3], [5, 6], [8, 9]]);
    expect(call('DROP', G, 1, -1)).toEqual([[4, 5], [7, 8]]);
    expect(call('DROP', G, 0)).toEqual(G);
    expect(call('DROP', G, 0, 0)).toEqual(G);
  });

  it('DROP reports #CALC! when nothing is left', () => {
    expect(err('DROP', G, 3)).toBe('#CALC!');
    expect(err('DROP', G, -5)).toBe('#CALC!');
    expect(err('DROP', G, 0, 3)).toBe('#CALC!');
    expect(err('DROP', G)).toBe('#VALUE!');
  });
});

describe('HSTACK and VSTACK', () => {
  it('join arrays and pad the shorter ones with #N/A', () => {
    expect(call('HSTACK', [[1], [4]], [[2], [5]])).toEqual([[1, 2], [4, 5]]);
    expect(call('HSTACK', [[1], [4]], [[2], [5], [8]])).toEqual([[1, 2], [4, 5], ['#N/A', 8]]);
    expect(call('HSTACK', [[1, 2]], 7, 3)).toEqual([[1, 2, 7, 3]]);
    expect(call('HSTACK', [[1], [2]], G.slice(0, 2).map((r) => r.slice(0, 2)))).toEqual([[1, 1, 2], [2, 4, 5]]);
    expect(call('HSTACK', [[1], [4]])).toEqual([[1], [4]]);
    expect(call('VSTACK', [[1, 2]], [[4, 5]])).toEqual([[1, 2], [4, 5]]);
    expect(call('VSTACK', [[1, 2, 3]], [[4, 5]])).toEqual([[1, 2, 3], [4, 5, '#N/A']]);
    expect(call('VSTACK', 1, 2, 3)).toEqual([[1], [2], [3]]);
    expect(call('VSTACK', [[1, 2]], 'x')).toEqual([[1, 2], ['x', '#N/A']]);
    expect(call('VSTACK', [[1], [null]], [[2]])).toEqual([[1], [0], [2]]);
    expect(call('VSTACK', [1, 2, 3], [[4]])).toEqual([[1, 2, 3], [4, '#N/A', '#N/A']]);
  });

  it('check argument counts and refuse a result of more than a million cells', () => {
    expect(err('HSTACK')).toBe('#VALUE!');
    expect(err('VSTACK')).toBe('#VALUE!');
    const big = Array.from({ length: 1001 }, () => Array<number>(1000).fill(1));
    expect(err('VSTACK', big, [[1]])).toBe('#NUM!');
    expect(err('HSTACK', big, [[1]])).toBe('#NUM!');
  });
});

describe('EXPAND', () => {
  const small = [
    [1, 2],
    [4, 5],
  ];

  it('grows an array, padding with #N/A or pad_with', () => {
    expect(call('EXPAND', small, 3, 3)).toEqual([[1, 2, '#N/A'], [4, 5, '#N/A'], ['#N/A', '#N/A', '#N/A']]);
    expect(call('EXPAND', small, 3, undefined, 0)).toEqual([[1, 2], [4, 5], [0, 0]]);
    expect(call('EXPAND', small, undefined, 3, 'x')).toEqual([[1, 2, 'x'], [4, 5, 'x']]);
    expect(call('EXPAND', small, 2, 2)).toEqual(small);
    expect(call('EXPAND', 5, 1, 2)).toEqual([[5, '#N/A']]);
    expect(call('EXPAND', [[null]], 1, 1)).toEqual([[0]]);
  });

  it('rejects shrinking, a result over a million cells and a wrong argument count', () => {
    expect(err('EXPAND', small, 1)).toBe('#VALUE!');
    expect(err('EXPAND', small, 2, 1)).toBe('#VALUE!');
    expect(err('EXPAND', small, 2000, 2000)).toBe('#NUM!');
    expect(err('EXPAND', small)).toBe('#VALUE!');
    expect(err('EXPAND', small, 2, 2, 0, 0)).toBe('#VALUE!');
  });
});

describe('TOCOL and TOROW', () => {
  const mixed = [
    [1, 2],
    [null, '#N/A'],
  ];

  it('list the elements by row, or by column on request', () => {
    expect(call('TOCOL', G)).toEqual([[1], [2], [3], [4], [5], [6], [7], [8], [9]]);
    expect(call('TOCOL', [[1, 2], [4, 5]], undefined, true)).toEqual([[1], [4], [2], [5]]);
    expect(call('TOROW', G)).toEqual([[1, 2, 3, 4, 5, 6, 7, 8, 9]]);
    expect(call('TOROW', [[1, 2], [4, 5]], 0, true)).toEqual([[1, 4, 2, 5]]);
    expect(call('TOROW', [[1], [4], [7]])).toEqual([[1, 4, 7]]);
    expect(call('TOCOL', 5)).toEqual([[5]]);
  });

  it('skip blanks, errors or both, and turn the blanks they keep into 0', () => {
    expect(call('TOCOL', mixed)).toEqual([[1], [2], [0], ['#N/A']]);
    expect(call('TOCOL', mixed, 1)).toEqual([[1], [2], ['#N/A']]);
    expect(call('TOCOL', mixed, 2)).toEqual([[1], [2], [0]]);
    expect(call('TOCOL', mixed, 3)).toEqual([[1], [2]]);
    expect(call('TOROW', mixed, 3)).toEqual([[1, 2]]);
    expect(call('TOROW', [['', 'x', 'x']], 1)).toEqual([['', 'x', 'x']]); // an empty text is not a blank cell
  });

  it('report #CALC! when nothing is left and #VALUE! for a bad ignore mode', () => {
    expect(err('TOCOL', [[null, '#N/A']], 3)).toBe('#CALC!');
    expect(err('TOCOL', [])).toBe('#CALC!');
    expect(err('TOCOL', G, 4)).toBe('#VALUE!');
    expect(err('TOROW', G, -1)).toBe('#VALUE!');
    expect(err('TOROW')).toBe('#VALUE!');
    expect(err('TOROW', G, 0, true, 1)).toBe('#VALUE!');
  });
});

describe('WRAPROWS and WRAPCOLS', () => {
  const five = [[1], [2], [3], [4], [5]];

  it('wrap a vector into lines of a given length, padding the last', () => {
    expect(call('WRAPROWS', five, 2)).toEqual([[1, 2], [3, 4], [5, '#N/A']]);
    expect(call('WRAPROWS', five, 2, 0)).toEqual([[1, 2], [3, 4], [5, 0]]);
    expect(call('WRAPROWS', five, 5)).toEqual([[1, 2, 3, 4, 5]]);
    expect(call('WRAPROWS', [[1, 2, 3]], 2, 'x')).toEqual([[1, 2], [3, 'x']]);
    expect(call('WRAPROWS', five, 2.9)).toEqual([[1, 2], [3, 4], [5, '#N/A']]);
    expect(call('WRAPROWS', 7, 2)).toEqual([[7, '#N/A']]);
    expect(call('WRAPROWS', [[1], [null]], 2)).toEqual([[1, 0]]);
    expect(call('WRAPCOLS', five, 2)).toEqual([[1, 3, 5], [2, 4, '#N/A']]);
    expect(call('WRAPCOLS', [[1, 2, 3, 4, 5]], 2, '-')).toEqual([[1, 3, 5], [2, 4, '-']]);
    expect(call('WRAPCOLS', five, 5)).toEqual([[1], [2], [3], [4], [5]]);
    expect(call('WRAPCOLS', five, 8)).toEqual([[1], [2], [3], [4], [5], ['#N/A'], ['#N/A'], ['#N/A']]);
  });

  it('require a vector and a wrap count of at least 1', () => {
    expect(err('WRAPROWS', G, 2)).toBe('#VALUE!');
    expect(err('WRAPCOLS', G, 2)).toBe('#VALUE!');
    expect(err('WRAPROWS', five, 0)).toBe('#NUM!');
    expect(err('WRAPCOLS', five, -1)).toBe('#NUM!');
    expect(err('WRAPROWS', five, 0.5)).toBe('#NUM!');
    expect(err('WRAPROWS', five, 'x')).toBe('#VALUE!');
    expect(err('WRAPROWS', five)).toBe('#VALUE!');
    expect(err('WRAPCOLS', five, 2, 0, 0)).toBe('#VALUE!');
  });

  it('refuse a result of more than a million cells', () => {
    const many = Array.from({ length: 10 }, () => [1]);
    expect(err('WRAPROWS', many, 2_000_000)).toBe('#NUM!');
    expect(err('WRAPCOLS', many, 2_000_000)).toBe('#NUM!');
  });
});

describe('large arrays and properties', () => {
  const grid = (rows: number, cols: number): number[][] =>
    Array.from({ length: rows }, (_, r) => Array.from({ length: cols }, (_, c) => r * cols + c));

  it('stay linear and never overflow the stack on a million cells', () => {
    const big = grid(1000, 1000);
    const started = Date.now();
    expect((call('TRANSPOSE', big) as unknown[][]).length).toBe(1000);
    expect((call('TOCOL', big) as unknown[][]).length).toBe(1_000_000);
    expect((call('TOROW', big, 0, true) as unknown[][])[0].length).toBe(1_000_000);
    expect((call('TAKE', big, -10, -10) as unknown[][]).length).toBe(10);
    expect((call('DROP', big, 990) as unknown[][]).length).toBe(10);
    expect((call('CHOOSEROWS', big, -1, 1) as unknown[][]).length).toBe(2);
    expect((call('WRAPROWS', [big[0].concat(big[1])], 7) as unknown[][]).length).toBe(Math.ceil(2000 / 7));
    expect(Date.now() - started).toBeLessThan(5000);
  });

  it('TRANSPOSE twice is the identity, and TOCOL / TOROW hold the same elements', () => {
    for (const [rows, cols] of [[1, 1], [1, 7], [7, 1], [4, 6], [13, 3]]) {
      const m = grid(rows, cols);
      expect(call('TRANSPOSE', call('TRANSPOSE', m))).toEqual(m);
      expect((call('TOCOL', m) as unknown[][]).map((r) => r[0])).toEqual((call('TOROW', m) as unknown[][])[0]);
      expect((call('TOCOL', m, 0, true) as unknown[][]).map((r) => r[0])).toEqual((call('TOROW', call('TRANSPOSE', m)) as unknown[][])[0]);
    }
  });

  it('TAKE and DROP split an array: take n then drop n leaves what was left', () => {
    const m = grid(6, 5);
    for (let n = 1; n < 6; n++) {
      expect(call('VSTACK', call('TAKE', m, n), call('DROP', m, n))).toEqual(m);
      expect(call('VSTACK', call('DROP', m, -n), call('TAKE', m, -n))).toEqual(m);
    }
    for (let n = 1; n < 5; n++) expect(call('HSTACK', call('TAKE', m, undefined, n), call('DROP', m, undefined, n))).toEqual(m);
  });
});
