import type { EvaluationContext } from '../types';
import { textExtFunctions } from './text-ext';

const ctx: EvaluationContext = { getCellValue: () => null, getRangeValues: () => [] };

function call(name: string, ...args: unknown[]): unknown {
  return textExtFunctions[name](args, ctx);
}

function err(name: string, ...args: unknown[]): string {
  try {
    call(name, ...args);
  } catch (e) {
    return (e as Error).message;
  }
  return 'no error';
}

describe('CLEAN', () => {
  it('removes the characters 0-31 and nothing else', () => {
    expect(call('CLEAN', 'a\u0007b\nc\u001fd')).toBe('abcd');
    expect(call('CLEAN', 'x\u007f\u0080y')).toBe('x\u007f\u0080y');
    expect(call('CLEAN', 123)).toBe('123');
    expect(call('CLEAN', null)).toBe('');
  });

  it('propagates an error value, lifts over arrays and checks its arity', () => {
    expect(err('CLEAN', '#N/A')).toBe('#N/A');
    expect(call('CLEAN', [['a\n'], ['b']])).toEqual([['a'], ['b']]);
    expect(err('CLEAN')).toBe('#VALUE!');
    expect(err('CLEAN', 'a', 'b')).toBe('#VALUE!');
  });
});

describe('NUMBERVALUE', () => {
  it('reads a number with the default and the given separators', () => {
    expect(call('NUMBERVALUE', '1,234.5')).toBe(1234.5);
    expect(call('NUMBERVALUE', '1.234,5', ',', '.')).toBe(1234.5);
    expect(call('NUMBERVALUE', '1,5', ',')).toBe(1.5);
    expect(call('NUMBERVALUE', '1.5', '.')).toBe(1.5);
    expect(call('NUMBERVALUE', '1 234,5', ',', ' ')).toBe(1234.5);
    expect(call('NUMBERVALUE', ' 3 000 ')).toBe(3000);
    expect(call('NUMBERVALUE', '-5')).toBe(-5);
    expect(call('NUMBERVALUE', '+.5')).toBe(0.5);
    expect(call('NUMBERVALUE', '5.')).toBe(5);
    expect(call('NUMBERVALUE', '1e3')).toBe(1000);
    expect(call('NUMBERVALUE', 12.5)).toBe(12.5);
  });

  it('divides by 100 for every trailing percent sign', () => {
    expect(call('NUMBERVALUE', '5%')).toBe(0.05);
    expect(call('NUMBERVALUE', '3%%')).toBeCloseTo(0.0003, 15);
    expect(call('NUMBERVALUE', ' 3 % ')).toBe(0.03);
    expect(err('NUMBERVALUE', '%5')).toBe('#VALUE!');
    expect(err('NUMBERVALUE', '%')).toBe('#VALUE!');
  });

  it('treats an empty text as 0 and uses the first character of a separator', () => {
    expect(call('NUMBERVALUE', '')).toBe(0);
    expect(call('NUMBERVALUE', '  ')).toBe(0);
    expect(call('NUMBERVALUE', null)).toBe(0);
    expect(call('NUMBERVALUE', '1;5', ';;')).toBe(1.5);
    expect(call('NUMBERVALUE', '1.5', '', '')).toBe(1.5); // an empty separator is the default
    expect(err('NUMBERVALUE', '2,5', undefined, '.')).toBe('#VALUE!'); // the default decimal point is the given group separator
  });

  it('rejects malformed text, repeated or misplaced separators and booleans', () => {
    expect(err('NUMBERVALUE', 'abc')).toBe('#VALUE!');
    expect(err('NUMBERVALUE', '1.2.3')).toBe('#VALUE!');
    expect(err('NUMBERVALUE', '1.5,2')).toBe('#VALUE!'); // a group separator after the decimal point
    expect(err('NUMBERVALUE', '1', '.', '.')).toBe('#VALUE!'); // the two separators must differ
    expect(err('NUMBERVALUE', true)).toBe('#VALUE!');
    expect(err('NUMBERVALUE', '#N/A')).toBe('#N/A');
    expect(err('NUMBERVALUE')).toBe('#VALUE!');
    expect(err('NUMBERVALUE', '1', '.', ',', 'x')).toBe('#VALUE!');
  });

  it('moves the default group separator out of the way of a comma decimal', () => {
    expect(call('NUMBERVALUE', '1.234,5', ',')).toBe(1234.5);
  });

  it('lifts over arrays', () => {
    expect(call('NUMBERVALUE', [['1,5'], ['x']], ',')).toEqual([[1.5], ['#VALUE!']]);
  });
});

describe('UNICHAR and UNICODE', () => {
  it('convert between code points and characters', () => {
    expect(call('UNICHAR', 65)).toBe('A');
    expect(call('UNICHAR', 8364)).toBe('€');
    expect(call('UNICHAR', 128512)).toBe('😀');
    expect(call('UNICHAR', 66.9)).toBe('B');
    expect(call('UNICODE', 'A')).toBe(65);
    expect(call('UNICODE', '€uro')).toBe(8364);
    expect(call('UNICODE', '😀x')).toBe(128512);
    expect(call('UNICODE', 12)).toBe(49);
  });

  it('reject code points that are not characters', () => {
    expect(err('UNICHAR', 0)).toBe('#VALUE!');
    expect(err('UNICHAR', -5)).toBe('#VALUE!');
    expect(err('UNICHAR', 1114112)).toBe('#VALUE!');
    expect(err('UNICHAR', 55296)).toBe('#VALUE!');
    expect(err('UNICHAR', 57343)).toBe('#VALUE!');
    expect(call('UNICHAR', 55295)).toBe('퟿');
    expect(call('UNICHAR', 57344)).toBe('');
    expect(err('UNICHAR', 'x')).toBe('#VALUE!');
    expect(err('UNICODE', '')).toBe('#VALUE!');
    expect(err('UNICODE', '#N/A')).toBe('#N/A');
    expect(err('UNICHAR')).toBe('#VALUE!');
  });

  it('lift over arrays', () => {
    expect(call('UNICHAR', [[97], [98]])).toEqual([['a'], ['b']]);
    expect(call('UNICODE', [['a'], ['']])).toEqual([[97], ['#VALUE!']]);
  });
});

describe('T', () => {
  it('keeps text and blanks out everything else', () => {
    expect(call('T', 'abc')).toBe('abc');
    expect(call('T', '')).toBe('');
    expect(call('T', 1)).toBe('');
    expect(call('T', true)).toBe('');
    expect(call('T', null)).toBe('');
    expect(err('T', '#N/A')).toBe('#N/A');
    expect(call('T', [['p'], [7]])).toEqual([['p'], ['']]);
    expect(err('T')).toBe('#VALUE!');
  });
});

describe('FIXED and DOLLAR', () => {
  it('round half away from zero on the decimal value, with thousands commas', () => {
    expect(call('FIXED', 1234.5678, 2)).toBe('1,234.57');
    expect(call('FIXED', 1234.5678)).toBe('1,234.57');
    expect(call('FIXED', 1234.5678, 0)).toBe('1,235');
    expect(call('FIXED', 1234.5678, -2)).toBe('1,200');
    expect(call('FIXED', 1234.5678, 1, true)).toBe('1234.6');
    expect(call('FIXED', -1234.5678, 1)).toBe('-1,234.6');
    expect(call('FIXED', 1.005, 2)).toBe('1.01');
    expect(call('FIXED', 2.675, 2)).toBe('2.68');
    expect(call('FIXED', 0.5, 0)).toBe('1');
    expect(call('FIXED', 0.005, 2)).toBe('0.01');
    expect(call('FIXED', 0.001, 2)).toBe('0.00');
    expect(call('FIXED', 0.0001, 2)).toBe('0.00');
    expect(call('FIXED', 0, 2)).toBe('0.00');
    expect(call('FIXED', 999.995, 2)).toBe('1,000.00');
    expect(call('FIXED', 5, -3)).toBe('0');
    expect(call('FIXED', 500, -3)).toBe('1,000');
    expect(call('FIXED', 1234.5678, undefined, true)).toBe('1234.57');
    expect(call('FIXED', 1e15, 1)).toBe('1,000,000,000,000,000.0');
    expect(call('FIXED', 1.2345678901234568e17, 0)).toBe('123,456,789,012,346,000');
    expect(call('FIXED', 1234, -400)).toBe('0');
    expect(call('FIXED', 1.5, 3)).toBe('1.500');
    expect(call('FIXED', 'x'.length)).toBe('1.00');
  });

  it('never shows a negative zero', () => {
    expect(call('FIXED', -0.4, 0)).toBe('0');
    expect(call('FIXED', -0.001, 2)).toBe('0.00');
    expect(call('FIXED', -0.005, 2)).toBe('-0.01');
  });

  it('rejects bad arguments and lifts over arrays', () => {
    expect(err('FIXED', 1, 128)).toBe('#VALUE!');
    expect(call('FIXED', 1, 127)).toBe('1.' + '0'.repeat(127));
    expect(err('FIXED', 'x')).toBe('#VALUE!');
    expect(err('FIXED', '#N/A')).toBe('#N/A');
    expect(err('FIXED')).toBe('#VALUE!');
    expect(call('FIXED', [[1.25], [2.35]], 1)).toEqual([['1.3'], ['2.4']]);
  });

  it('DOLLAR adds a dollar sign and writes a negative amount in parentheses', () => {
    expect(call('DOLLAR', 1234.567)).toBe('$1,234.57');
    expect(call('DOLLAR', 1234.567, 0)).toBe('$1,235');
    expect(call('DOLLAR', 1234.567, -2)).toBe('$1,200');
    expect(call('DOLLAR', -1234.567)).toBe('($1,234.57)');
    expect(call('DOLLAR', -0.001)).toBe('$0.00');
    expect(err('DOLLAR', 1, 200)).toBe('#VALUE!');
    expect(err('DOLLAR')).toBe('#VALUE!');
  });
});

describe('BASE and DECIMAL', () => {
  it('write numbers in another base and read them back', () => {
    expect(call('BASE', 255, 16)).toBe('FF');
    expect(call('BASE', 255, 2, 10)).toBe('0011111111');
    expect(call('BASE', 0, 2)).toBe('0');
    expect(call('BASE', 35, 36)).toBe('Z');
    expect(call('BASE', 7.9, 2)).toBe('111');
    expect(call('BASE', 2 ** 53 - 1, 16)).toBe('1FFFFFFFFFFFFF');
    expect(call('DECIMAL', 'FF', 16)).toBe(255);
    expect(call('DECIMAL', 'ff', 16)).toBe(255);
    expect(call('DECIMAL', 'zz', 36)).toBe(1295);
    expect(call('DECIMAL', '11', 2)).toBe(3);
    expect(call('DECIMAL', '', 2)).toBe(0);
    expect(call('DECIMAL', 11, 2)).toBe(3);
  });

  it('report #NUM! for arguments out of range', () => {
    expect(err('BASE', -1, 2)).toBe('#NUM!');
    expect(err('BASE', 2 ** 53, 2)).toBe('#NUM!');
    expect(err('BASE', 5, 1)).toBe('#NUM!');
    expect(err('BASE', 5, 37)).toBe('#NUM!');
    expect(err('BASE', 5, 2, -1)).toBe('#NUM!');
    expect(err('BASE', 5, 2, 256)).toBe('#NUM!');
    expect(err('BASE', 'x', 2)).toBe('#VALUE!');
    expect(err('DECIMAL', '12', 2)).toBe('#NUM!');
    expect(err('DECIMAL', '1.1', 2)).toBe('#NUM!');
    expect(err('DECIMAL', ' 1', 2)).toBe('#NUM!');
    expect(err('DECIMAL', 'FF', 37)).toBe('#NUM!');
    expect(err('DECIMAL', 'FF', 1)).toBe('#NUM!');
    expect(err('DECIMAL', 'F'.repeat(256), 16)).toBe('#NUM!');
    expect(err('DECIMAL', 'FF')).toBe('#VALUE!');
    expect(err('BASE', 5)).toBe('#VALUE!');
  });

  it('lift over arrays', () => {
    expect(call('BASE', [[5], [6]], 2)).toEqual([['101'], ['110']]);
    expect(call('DECIMAL', [['11'], ['2']], 2)).toEqual([[3], ['#NUM!']]);
  });
});

describe('ROMAN and ARABIC', () => {
  it('writes the five forms of Excel', () => {
    expect(call('ROMAN', 499)).toBe('CDXCIX');
    expect(call('ROMAN', 499, 1)).toBe('LDVLIV');
    expect(call('ROMAN', 499, 2)).toBe('XDIX');
    expect(call('ROMAN', 499, 3)).toBe('VDIV');
    expect(call('ROMAN', 499, 4)).toBe('ID');
    expect(call('ROMAN', 1994)).toBe('MCMXCIV');
    expect(call('ROMAN', 3999)).toBe('MMMCMXCIX');
    expect(call('ROMAN', 2024, true)).toBe('MMXXIV');
    expect(call('ROMAN', 2024, false)).toBe('MMXXIV');
    expect(call('ROMAN', 1999, false)).toBe('MIM');
    expect(call('ROMAN', 4.9)).toBe('IV');
    expect(call('ROMAN', 0)).toBe('');
  });

  it('rejects numbers and forms out of range', () => {
    expect(err('ROMAN', -1)).toBe('#VALUE!');
    expect(err('ROMAN', 4000)).toBe('#VALUE!');
    expect(err('ROMAN', 10, 5)).toBe('#VALUE!');
    expect(err('ROMAN', 10, -1)).toBe('#VALUE!');
    expect(err('ROMAN', 'x')).toBe('#VALUE!');
    expect(err('ROMAN')).toBe('#VALUE!');
  });

  it('reads Roman numerals, a leading minus and an empty text', () => {
    expect(call('ARABIC', 'MCMXCIX')).toBe(1999);
    expect(call('ARABIC', 'mmxxiv')).toBe(2024);
    expect(call('ARABIC', 'LVII')).toBe(57);
    expect(call('ARABIC', '')).toBe(0);
    expect(call('ARABIC', '-XII')).toBe(-12);
    expect(call('ARABIC', 'IIII')).toBe(4);
    expect(err('ARABIC', 'ABC')).toBe('#VALUE!');
    expect(err('ARABIC', 'X I')).toBe('#VALUE!');
    expect(err('ARABIC', 'I'.repeat(256))).toBe('#VALUE!');
    expect(err('ARABIC', '#N/A')).toBe('#N/A');
    expect(err('ARABIC')).toBe('#VALUE!');
  });

  it('round-trips every number from 1 to 3999 in every form', () => {
    for (let n = 1; n <= 3999; n++) {
      for (let form = 0; form <= 4; form++) expect(call('ARABIC', call('ROMAN', n, form))).toBe(n);
    }
  });

  it('lift over arrays', () => {
    expect(call('ROMAN', [[4], [9]])).toEqual([['IV'], ['IX']]);
    expect(call('ARABIC', [['IV'], ['q']])).toEqual([[4], ['#VALUE!']]);
  });
});
