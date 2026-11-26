import type { EvaluationContext } from '../types';
import { lookupFunctions } from './lookup';

const ctx: EvaluationContext = {
  getCellValue: () => null,
  getRangeValues: () => [],
};

function call(name: string, args: unknown[]): unknown {
  return lookupFunctions[name](args, ctx);
}

describe('VLOOKUP / HLOOKUP', () => {
  const table = [
    [1, 'a'],
    [2, 'b'],
    [3, 'c'],
  ];

  it('VLOOKUP exact match', () => {
    expect(call('VLOOKUP', [2, table, 2, false])).toBe('b');
    expect(() => call('VLOOKUP', [9, table, 2, false])).toThrow('#N/A');
  });

  it('VLOOKUP approximate match finds the largest key not exceeding the lookup', () => {
    const sorted = [
      [1, 'a'],
      [3, 'b'],
      [5, 'c'],
    ];
    expect(call('VLOOKUP', [4, sorted, 2, true])).toBe('b');
    expect(call('VLOOKUP', [4, sorted, 2])).toBe('b');
  });

  it('VLOOKUP rejects an out-of-range column index', () => {
    expect(() => call('VLOOKUP', [1, table, 5, false])).toThrow('#REF!');
  });

  it('HLOOKUP searches the header row', () => {
    const grid = [
      ['x', 'y', 'z'],
      [1, 2, 3],
    ];
    expect(call('HLOOKUP', ['y', grid, 2, false])).toBe(2);
  });
});

describe('XLOOKUP / XMATCH', () => {
  const keys = [[1], [2], [3]];
  const values = [['a'], ['b'], ['c']];

  it('XLOOKUP returns the aligned value or the not-found fallback', () => {
    expect(call('XLOOKUP', [2, keys, values])).toBe('b');
    expect(call('XLOOKUP', [9, keys, values, 'missing'])).toBe('missing');
    expect(() => call('XLOOKUP', [9, keys, values])).toThrow('#N/A');
  });

  it('XLOOKUP match mode 1 picks the next larger value', () => {
    expect(call('XLOOKUP', [2.5, keys, values, undefined, 1])).toBe('c');
  });

  it('XMATCH returns a 1-based position', () => {
    expect(call('XMATCH', [20, [[10], [20], [30]]])).toBe(2);
    expect(call('XMATCH', [25, [[10], [20], [30]], -1])).toBe(2);
  });
});

describe('MATCH / INDEX / LOOKUP', () => {
  it('MATCH supports exact, wildcard and approximate modes', () => {
    expect(call('MATCH', [20, [[10], [20], [30]], 0])).toBe(2);
    expect(call('MATCH', [25, [[10], [20], [30]], 1])).toBe(2);
    expect(call('MATCH', ['ban*', [['apple'], ['banana']], 0])).toBe(2);
    expect(() => call('MATCH', [5, [[10], [20]], 1])).toThrow('#N/A');
  });

  it('INDEX reads a cell from a matrix or vector', () => {
    const matrix = [
      [1, 2, 3],
      [4, 5, 6],
    ];
    expect(call('INDEX', [matrix, 2, 3])).toBe(6);
    expect(call('INDEX', [[[1, 2, 3]], 2])).toBe(2);
    expect(call('INDEX', [[[1], [2], [3]], 3])).toBe(3);
    expect(() => call('INDEX', [matrix, 9, 1])).toThrow('#REF!');
  });

  it('INDEX with a row or column of 0, or one left blank, returns the whole column or row', () => {
    const m = [
      [1, 2, 3],
      [4, 5, 6],
      [7, 8, 9],
    ];
    expect(call('INDEX', [m, 0, 2])).toEqual([[2], [5], [8]]);
    expect(call('INDEX', [m, undefined, 3])).toEqual([[3], [6], [9]]);
    expect(call('INDEX', [m, 2, 0])).toEqual([[4, 5, 6]]);
    expect(call('INDEX', [m, 3, undefined])).toEqual([[7, 8, 9]]);
    expect(call('INDEX', [m, 0, 0])).toEqual(m);
    expect(call('INDEX', [m, undefined, undefined])).toEqual(m);
    expect(call('INDEX', [m, 2])).toEqual([[4, 5, 6]]); // one index into a table is a whole row
    expect(call('INDEX', [m, undefined])).toEqual(m);
    expect(call('INDEX', [m, 2.9, 1.9])).toBe(4);
    expect(call('INDEX', [m, 2, 2, 1])).toBe(5);
    expect(call('INDEX', [m, 2, 2, undefined])).toBe(5); // an omitted area is the first
    expect(call('INDEX', [[[1, 2, 3]], 0])).toEqual([[1, 2, 3]]);
    expect(call('INDEX', [[[1], [2]], 0])).toEqual([[1], [2]]);
    expect(call('INDEX', [[[1, 2, 3]], 0, 2])).toBe(2);
    expect(call('INDEX', [[[1, 2, 3]]])).toBe(1);
    expect(call('INDEX', [[[1], [2]]])).toBe(1);
    expect(call('INDEX', [7, 1, 1])).toBe(7);
    expect(call('INDEX', [5, 1])).toBe(5);
  });

  it('INDEX rejects positions outside the array, negative ones and a missing index', () => {
    const m = [
      [1, 2],
      [3, 4],
    ];
    expect(() => call('INDEX', [m, 3, 0])).toThrow('#REF!');
    expect(() => call('INDEX', [m, 0, 3])).toThrow('#REF!');
    expect(() => call('INDEX', [m, 3])).toThrow('#REF!');
    expect(() => call('INDEX', [m, -1, 1])).toThrow('#VALUE!');
    expect(() => call('INDEX', [m, 1, -1])).toThrow('#VALUE!');
    expect(() => call('INDEX', [m])).toThrow('#REF!');
    expect(() => call('INDEX', [m, 1, 1, 2])).toThrow('#REF!'); // there is only one area
    expect(() => call('INDEX', [])).toThrow('#VALUE!');
    expect(() => call('INDEX', [m, 1, 1, 1, 1])).toThrow('#VALUE!');
    expect(() => call('INDEX', [m, 'x', 1])).toThrow('#VALUE!');
  });

  it('LOOKUP finds the value in the result vector', () => {
    expect(call('LOOKUP', [4, [[1], [3], [5]], [['a'], ['b'], ['c']]])).toBe('b');
  });
});

describe('reference helpers', () => {
  it('CHOOSE returns the indexed argument', () => {
    expect(call('CHOOSE', [2, 'a', 'b', 'c'])).toBe('b');
    expect(() => call('CHOOSE', [0, 'a', 'b'])).toThrow('#VALUE!');
  });

  it('ROWS / COLUMNS count range dimensions', () => {
    expect(call('ROWS', [[[1], [2], [3]]])).toBe(3);
    expect(call('COLUMNS', [[[1, 2, 3]]])).toBe(3);
    expect(call('ROWS', [42])).toBe(1);
  });

  it('ADDRESS builds an A1 reference', () => {
    expect(call('ADDRESS', [1, 1])).toBe('$A$1');
    expect(call('ADDRESS', [2, 3, 4])).toBe('C2');
    expect(call('ADDRESS', [5, 27, 1])).toBe('$AA$5');
    expect(call('ADDRESS', [1, 1, 1, true, 'Sheet1'])).toBe('Sheet1!$A$1');
  });

  it('HYPERLINK returns the friendly name when given one', () => {
    expect(call('HYPERLINK', ['https://example.com', 'click'])).toBe('click');
    expect(call('HYPERLINK', ['https://example.com'])).toBe('https://example.com');
  });
});
