import type { EvaluationContext } from '../types';
import { informationEagerFunctions, informationLazyFunctions } from './information';

const ctx: EvaluationContext = {
  getCellValue: () => null,
  getRangeValues: () => [],
};

function callEager(name: string, args: unknown[]): unknown {
  return informationEagerFunctions[name](args, ctx);
}

function callLazy(name: string, thunks: Array<() => unknown>): unknown {
  return informationLazyFunctions[name](thunks, ctx);
}

const boom = (msg: string) => () => {
  throw new Error(msg);
};

describe('ISNONTEXT', () => {
  it('is TRUE for everything but text, errors and blanks included', () => {
    expect(callLazy('ISNONTEXT', [() => 'x'])).toBe(false);
    expect(callLazy('ISNONTEXT', [() => ''])).toBe(false);
    expect(callLazy('ISNONTEXT', [() => 1])).toBe(true);
    expect(callLazy('ISNONTEXT', [() => null])).toBe(true);
    expect(callLazy('ISNONTEXT', [() => true])).toBe(true);
    expect(callLazy('ISNONTEXT', [() => '#N/A'])).toBe(true);
    expect(callLazy('ISNONTEXT', [boom('#DIV/0!')])).toBe(true);
  });
});

describe('type predicates', () => {
  it('ISBLANK / ISNUMBER / ISTEXT / ISLOGICAL', () => {
    expect(callLazy('ISBLANK', [() => null])).toBe(true);
    expect(callLazy('ISBLANK', [() => 0])).toBe(false);
    expect(callLazy('ISNUMBER', [() => 42])).toBe(true);
    expect(callLazy('ISNUMBER', [() => 'x'])).toBe(false);
    expect(callLazy('ISTEXT', [() => 'x'])).toBe(true);
    expect(callLazy('ISTEXT', [() => '#REF!'])).toBe(false);
    expect(callLazy('ISLOGICAL', [() => true])).toBe(true);
    expect(callLazy('ISLOGICAL', [() => 1])).toBe(false);
  });

  it('a predicate reports false rather than throwing when the argument errors', () => {
    expect(callLazy('ISNUMBER', [boom('#DIV/0!')])).toBe(false);
    expect(callLazy('ISTEXT', [boom('#DIV/0!')])).toBe(false);
  });

  it('TYPE classifies the value, returning 16 for an error', () => {
    expect(callLazy('TYPE', [() => 5])).toBe(1);
    expect(callLazy('TYPE', [() => 'x'])).toBe(2);
    expect(callLazy('TYPE', [() => true])).toBe(4);
    expect(callLazy('TYPE', [() => '#REF!'])).toBe(16);
    expect(callLazy('TYPE', [boom('#DIV/0!')])).toBe(16);
  });
});

describe('error predicates', () => {
  it('ISERROR catches every error; ISERR excludes #N/A', () => {
    expect(callLazy('ISERROR', [boom('#DIV/0!')])).toBe(true);
    expect(callLazy('ISERROR', [() => '#N/A'])).toBe(true);
    expect(callLazy('ISERROR', [() => 5])).toBe(false);
    expect(callLazy('ISERR', [() => '#REF!'])).toBe(true);
    expect(callLazy('ISERR', [() => '#N/A'])).toBe(false);
  });

  it('ISNA only matches #N/A', () => {
    expect(callLazy('ISNA', [() => '#N/A'])).toBe(true);
    expect(callLazy('ISNA', [boom('#N/A')])).toBe(true);
    expect(callLazy('ISNA', [() => '#REF!'])).toBe(false);
  });

  it('ERROR.TYPE maps an error to its code, #N/A for a non-error', () => {
    expect(callLazy('ERROR.TYPE', [() => '#DIV/0!'])).toBe(2);
    expect(callLazy('ERROR.TYPE', [boom('#REF!')])).toBe(4);
    expect(() => callLazy('ERROR.TYPE', [() => 5])).toThrow('#N/A');
  });
});

describe('NA / N', () => {
  it('NA always raises #N/A', () => {
    expect(() => callEager('NA', [])).toThrow('#N/A');
  });

  it('N coerces to a number', () => {
    expect(callEager('N', [42])).toBe(42);
    expect(callEager('N', [true])).toBe(1);
    expect(callEager('N', ['text'])).toBe(0);
    expect(() => callEager('N', ['#REF!'])).toThrow('#REF!');
  });
});
