import { boolAttr, intAttr, numAttr, parseDecimal, parseIndex, put } from './values';

describe('parseDecimal', () => {
  it.each([
    ['0', 0], ['12', 12], [' 12 ', 12], ['-1.5', -1.5], ['+2', 2], ['.5', 0.5], ['5.', 5], ['1E-3', 0.001], ['1e+3', 1000], ['-0.0', -0], ['123456789012345678901234567890', 1.2345678901234568e29], ['1.7976931348623157e308', 1.7976931348623157e308],
  ])('reads %j as %p', (text, n) => {
    expect(parseDecimal(text)).toBe(n);
  });

  it.each(['', '  ', 'Infinity', '-Infinity', '+Infinity', 'NaN', '0x10', '0b1', '0o7', '1_000', '1,5', '+', '-', '.', 'e5', '1e', '1e+', '1 2', '--1', '1e999', '-1e999', '１２', '12abc'])('does not read %j', (text) => {
    expect(parseDecimal(text)).toBeUndefined();
  });
});

describe('parseIndex', () => {
  it.each([['0', 0], ['7', 7], [' 7 ', 7], ['007', 7], ['999999999', 999_999_999]])('reads %j as %p', (text, n) => {
    expect(parseIndex(text)).toBe(n);
  });

  it.each(['', ' ', '-1', '+1', '1.5', '1e0', '0x1', 'x', '1000000000', '__proto__', '１'])('does not read %j', (text) => {
    expect(parseIndex(text)).toBeUndefined();
  });
});

describe('attribute helpers', () => {
  it('numAttr reads plain decimals only', () => {
    expect([numAttr(undefined), numAttr(''), numAttr('2.5'), numAttr('0x10'), numAttr('Infinity'), numAttr('1e999')]).toEqual([undefined, undefined, 2.5, undefined, undefined, undefined]);
  });

  it('intAttr reads whole numbers only', () => {
    expect([intAttr('3'), intAttr('3.5'), intAttr('0x3'), intAttr(undefined), intAttr('-4')]).toEqual([3, undefined, undefined, undefined, -4]);
  });

  it('boolAttr keeps "not written"', () => {
    expect([boolAttr(undefined), boolAttr('1'), boolAttr('true'), boolAttr('0')]).toEqual([undefined, true, true, false]);
  });

  it('put skips undefined', () => {
    const o: { a?: number } = {};
    put(o, 'a', undefined);
    expect(o).toEqual({});
    put(o, 'a', 1);
    expect(o).toEqual({ a: 1 });
  });
});
