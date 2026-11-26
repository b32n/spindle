// Hang and size limits: Excel's own limits (255 characters per code, five denominator digits, three fractional-second
// digits) keep every format cheap, and adversarial codes inside them still format within a time budget.

import { compileFormat, formatDisplay, formatValue, isValidFormatCode } from './index';
import { nearest } from './render';

const timed = <T,>(work: () => T): { result: T; ms: number } => {
  const start = performance.now();
  const result = work();
  return { result, ms: performance.now() - start };
};

describe('limits of a format code', () => {
  it('a code is at most 255 characters', () => {
    expect(isValidFormatCode('0'.repeat(255))).toBe(true);
    expect(isValidFormatCode('0'.repeat(256))).toBe(false);
    expect(compileFormat('0'.repeat(256)).error).toMatch(/255/);
    expect(formatValue(5, '0'.repeat(1000))).toEqual({ text: '5', invalid: true });
  });

  it('a denominator has at most five digits, variable or fixed', () => {
    expect(isValidFormatCode('# ?/?????')).toBe(true);
    expect(isValidFormatCode('# ?/??????')).toBe(false);
    expect(isValidFormatCode('# ?/??????????')).toBe(false);
    expect(isValidFormatCode('# ?/99999')).toBe(true);
    expect(isValidFormatCode('# ?/100000')).toBe(false);
    expect(isValidFormatCode(`# ?/${'1'.repeat(200)}`)).toBe(false);
    expect(formatDisplay(Math.PI, '# ?/?????')).toBe('3 14093/99532');
  });

  it('at most three fractional-second digits', () => {
    expect(isValidFormatCode('ss.000')).toBe(true);
    expect(isValidFormatCode('ss.0000')).toBe(false);
    expect(isValidFormatCode(`ss.${'0'.repeat(200)}`)).toBe(false);
  });

  it('a number too large for the field shows ####, never an exponent', () => {
    expect(formatDisplay(1e20, '[h]:mm:ss')).toBe('####');
    expect(formatDisplay(1e20, '[ss]')).toBe('####');
    expect(formatDisplay(1e300, '?/?')).toBe('####');
    expect(formatDisplay(1e15, '# ?/?')).toBe('####');
    expect(formatDisplay(1e300, '# ?/8')).toBe('####');
    for (const code of ['[h]:mm', '[mm]:ss.000', '# ??/??', '?/?', '# ?/16', 'h:mm', 'yyyy-mm-dd']) {
      for (const v of [1e15, 1e21, 1e100, 1e300, Number.MAX_VALUE]) expect(formatDisplay(v, code)).not.toMatch(/e[+-]?\d/i);
    }
  });
});

describe('many percent signs and scaling commas', () => {
  const pct = (code: string, n: number) => code + '%'.repeat(n);
  it('overflow of the scaled value shows ####, never NaN, Infinity or an exponent', () => {
    for (const v of [0, 0.5, 1, 1e300, Number.MAX_VALUE]) {
      for (const code of [pct('0', 200), pct('0.0E+0', 200), pct('# ?/?', 200), pct('#,##0.00', 200), pct('?/8', 200), pct('0.0E+00', 240)]) {
        const text = formatDisplay(v, code);
        expect(text).toBe('####');
      }
    }
    expect(formatDisplay(1e308, '0%%%%')).toBe('####');
    expect(formatDisplay(1e307, '0%')).toBe('####');
    expect(formatDisplay(1e305, '0%')).toMatch(/^10{307}%$/);
    expect(formatDisplay(1e300, '0.0E+00%%')).toBe('1.0E+304%%');
  });

  it('a long run of scaling commas divides to zero without garbage', () => {
    expect(formatDisplay(0.5, '0' + ','.repeat(200))).toBe('0');
    expect(formatDisplay(1e300, '0.0E+0' + ','.repeat(3))).toBe('1.0E+291');
  });

  it('whatever the code, a finite value never prints NaN or Infinity', () => {
    for (const code of [pct('0', 254), pct('0.', 253), pct('0E+0', 250), pct('0/', 250), ','.repeat(254) + '0', '0' + ','.repeat(254)]) {
      for (const v of [0, -0, 1, -1, 0.5, 1e-300, 1e300, 2958465]) expect(formatDisplay(v, code)).not.toMatch(/NaN|Infinity/);
    }
  });
});

describe('nearest fraction', () => {
  // brute force: the closest p/q with q <= max, the smallest q among equally close
  const brute = (x: number, max: number): [number, number] => {
    let best: [number, number] = [0, 1];
    let error = Infinity;
    for (let q = 1; q <= max; q++) {
      const p = Math.round(x * q);
      const e = Math.abs(x - p / q);
      if (e < error - 1e-13) {
        best = [p, q];
        error = e;
      }
    }
    return best;
  };

  it('agrees with a brute-force search for one, two and three digit denominators', () => {
    let seed = 99;
    const rand = () => ((seed = (seed * 1103515245 + 12345) & 0x7fffffff) / 0x7fffffff);
    for (const max of [9, 99, 999]) {
      for (let i = 0; i < 3000; i++) {
        const x = rand() * (i % 5 === 0 ? 20 : 1);
        const [p, q] = nearest(x, max);
        const [bp, bq] = brute(x, max);
        // equal fractions, or equally close (a tie between two different fractions)
        expect(Math.abs(x - p / q)).toBeLessThanOrEqual(Math.abs(x - bp / bq) + 1e-12);
        expect(q).toBeLessThanOrEqual(max);
        if (Math.abs(Math.abs(x - p / q) - Math.abs(x - bp / bq)) < 1e-13) expect(q).toBeLessThanOrEqual(bq);
      }
    }
  });

  it('handles whole numbers, tiny and huge values', () => {
    expect(nearest(0, 9)).toEqual([0, 1]);
    expect(nearest(3, 99)).toEqual([3, 1]);
    expect(nearest(0.5, 9)).toEqual([1, 2]);
    expect(nearest(1e-9, 99999)).toEqual([0, 1]);
    expect(nearest(Math.PI, 99999)[1]).toBeLessThanOrEqual(99999);
    expect(nearest(0.3333333333333, 5)).toEqual([1, 3]);
  });
});

describe('adversarial codes stay fast', () => {
  const unit = (s: string, n: number) => s.repeat(Math.floor(n / s.length)).slice(0, n);
  const codes = [
    ','.repeat(254) + '0', '0' + ','.repeat(254), unit('0,', 255), unit(',0', 255), unit('?/', 255), unit('0/', 255), '#'.repeat(255), '.'.repeat(255),
    '0.' + '0'.repeat(253), '[' + 'h'.repeat(253) + ']', '"'.repeat(255), unit('"a"', 255), unit('_0', 255), unit('*0', 255), unit(';', 255), unit('[Red]', 255),
    'y'.repeat(255), unit('m:', 255), unit('mm:ss.0', 255), unit('E+0', 255), unit('0E+0', 255), unit('0%', 255), unit('# ?/?', 255), unit('@', 255),
    unit('[>1]0;', 255), unit('\\', 254), 'General'.repeat(36), unit('AM/PM', 255), unit('[$-409]', 255), unit('0 ', 255),
  ];
  const values = [0, 1, -1, 0.5, 1234.5678, -0.256, 46024.75, 1e15, 1e-300, 2958465];

  it('compiles and formats 31 long and repetitive codes with ten values each within a second', () => {
    const { ms } = timed(() => {
      for (const code of codes) for (const v of values) expect(typeof formatValue(v, code).text).toBe('string');
    });
    expect(ms).toBeLessThan(1000);
  });

  it('compile time grows linearly with the length of a comma run (two sizes, best of three)', () => {
    compileFormat('x'); // warm up
    const best = (n: number) =>
      Math.min(...[0, 1, 2].map(() => timed(() => { for (let i = 0; i < 1500; i++) compileFormat(`0${','.repeat(n)}${i % 7}0`); }).ms));
    const small = best(60);
    const large = best(250);
    // 4x the commas: linear is about 4x; the old per-comma rescan was 16x and worse
    expect(large).toBeLessThan(Math.max(small, 5) * 10);
    expect(large).toBeLessThan(1000);
  });

  it('a five-digit denominator formats 5,000 values quickly', () => {
    const { ms } = timed(() => {
      for (let i = 1; i <= 5000; i++) formatDisplay(i / 7919 + (i % 9), '# ?????/?????');
    });
    expect(ms).toBeLessThan(1500);
  });

  it('the compiled-format cache holds only short codes, so its memory is bounded', () => {
    for (let i = 0; i < 2000; i++) compileFormat(`0.${'0'.repeat(i % 200)}"${i}"`.slice(0, 255));
    expect(compileFormat('0.00').ok).toBe(true);
  });
});

describe('General without a column', () => {
  it('fullGeneral gives the 15 significant digits', () => {
    const full = (v: number, code = 'General') => formatDisplay(v, code, { fullGeneral: true });
    expect(full(Math.PI)).toBe('3.14159265358979');
    expect(full(1 / 3)).toBe('0.333333333333333');
    expect(full(-1 / 3)).toBe('-0.333333333333333');
    expect(full(123456789012)).toBe('123456789012');
    expect(full(1e21)).toBe('1E+21');
    expect(full(1.5e22)).toBe('1.5E+22');
    expect(full(1e-7)).toBe('1E-07');
    expect(full(0.000001)).toBe('0.000001');
    expect(full(100)).toBe('100');
    expect(full(0)).toBe('0');
    expect(full(Math.PI, '@')).toBe('3.14159265358979');
    expect(full(Math.PI, '"x"General')).toBe('x3.14159265358979');
    expect(full(Math.PI, '[>10]0')).toBe('3.14159265358979');
    expect(formatDisplay(Math.PI, 'General')).toBe('3.141592654');
  });
});
