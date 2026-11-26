import { checkArity, liftScalar, toStrictNum } from './helpers';

describe('checkArity', () => {
  it('accepts counts inside the range and throws #VALUE! outside', () => {
    expect(() => checkArity([1], 1)).not.toThrow();
    expect(() => checkArity([1, 2], 1, 3)).not.toThrow();
    expect(() => checkArity([], 1)).toThrow('#VALUE!');
    expect(() => checkArity([1, 2, 3, 4], 1, 3)).toThrow('#VALUE!');
    expect(() => checkArity([1, 2, 3, 4], 1, Infinity)).not.toThrow();
  });
});

describe('toStrictNum', () => {
  it('is toNum except that empty text is #VALUE!', () => {
    expect(toStrictNum(' 4 ')).toBe(4);
    expect(toStrictNum(null)).toBe(0);
    expect(toStrictNum(true)).toBe(1);
    expect(() => toStrictNum('')).toThrow('#VALUE!');
    expect(() => toStrictNum('   ')).toThrow('#VALUE!');
    expect(() => toStrictNum('#N/A')).toThrow('#N/A');
    expect(() => toStrictNum('abc')).toThrow('#VALUE!');
  });
});

describe('liftScalar', () => {
  const add = (xs: unknown[]): unknown => (xs[0] as number) + ((xs[1] as number) ?? 0);

  it('calls the function directly when no argument is an array', () => {
    expect(liftScalar([1, 2], add)).toBe(3);
    expect(() =>
      liftScalar([1], () => {
        throw new Error('#NUM!');
      })
    ).toThrow('#NUM!');
  });

  it('maps over arrays and broadcasts scalars, rows and columns', () => {
    expect(liftScalar([[[1, 2]], 10], add)).toEqual([[11, 12]]);
    expect(liftScalar([[[1], [2]], [[10, 20]]], add)).toEqual([
      [11, 21],
      [12, 22],
    ]);
  });

  it('a bare array row (not nested) is read as one row', () => {
    expect(liftScalar([[1, 2], 1], add)).toEqual([[2, 3]]);
  });

  it('turns a failing element into its error code, and unknown failures into #VALUE!', () => {
    const fn = (xs: unknown[]): unknown => {
      if (xs[0] === 0) throw new Error('#DIV/0! detail');
      if (xs[0] === 1) throw new Error('boom');
      return xs[0];
    };
    expect(liftScalar([[[0], [1], [2]]], fn)).toEqual([['#DIV/0!'], ['#VALUE!'], [2]]);
  });

  it('marks positions without a counterpart #N/A and ragged rows too', () => {
    expect(
      liftScalar(
        [
          [[1], [2], [3]],
          [[1], [2]],
        ],
        add
      )
    ).toEqual([[2], [4], ['#N/A']]);
    expect(liftScalar([[[1, 2], [3]], [[1, 1]]], add)).toEqual([
      [2, 3],
      [4, '#N/A'],
    ]);
  });
});
