// Seeded property tests: the engine never throws, General is stable, section selection obeys its rules and the
// calendar round-trips in both date systems.

import { compileFormat, dateToSerial, formatDisplay, formatGeneral, formatValue, serialToDate } from './index';
import { selectSection } from './render';

function rng(seed: number): () => number {
  let a = seed >>> 0;
  return () => {
    a = (a + 0x6d2b79f5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

const pick = <T,>(r: () => number, list: readonly T[]): T => list[Math.floor(r() * list.length)];

const PIECES = ['0', '#', '?', '.', ',', '%', ';', '"', '\\', '[', ']', '_', '*', '@', 'E+', 'e-', 'y', 'm', 'd', 'h', 's', 'mm', 'yyyy', 'dd', 'AM/PM', 'a/p',
  '[h]', '[mm]', '[ss]', '[Red]', '[>5]', '[<=0]', '[$€-407]', '[$-F800]', 'General', '/', ' ', '-', '(', ')', '$', ':', '0.00', '#,##0', '"txt"', '.000', '1', '9', 'x', 'é', '€'];

function randomCode(r: () => number): string {
  const n = Math.floor(r() * 12);
  let code = '';
  for (let i = 0; i < n; i++) code += pick(r, PIECES);
  return code;
}

function randomValue(r: () => number): unknown {
  switch (Math.floor(r() * 12)) {
    case 0: return 0;
    case 1: return -0;
    case 2: return (r() - 0.5) * 10 ** Math.floor(r() * 40 - 20);
    case 3: return Math.floor(r() * 3_000_000);
    case 4: return r() * 100;
    case 5: return pick(r, [Infinity, -Infinity, NaN, Number.MAX_VALUE, Number.MIN_VALUE, 1e-300, 1e300, 2958465, 2958466, 0.5, 60, 61, -1]);
    case 6: return pick(r, ['', 'abc', '#N/A', '#VALUE!', ' ', '0']);
    case 7: return pick(r, [true, false, null, undefined]);
    case 8: return -r() * 1e6;
    case 9: return r() * 1e-5;
    case 10: return r();
    default: return Math.floor(r() * 1000) / 8;
  }
}

describe('formatValue never throws', () => {
  it('for random codes and values (25000 pairs)', () => {
    const r = rng(20260126);
    for (let i = 0; i < 25000; i++) {
      const code = randomCode(r);
      const value = randomValue(r);
      const result = formatValue(value, code, { dateSystem: r() < 0.5 ? 1900 : 1904 });
      expect(typeof result.text).toBe('string');
      if (result.invalid) expect(compileFormat(code).ok).toBe(false);
    }
  });

  it('for random printable strings as codes', () => {
    const r = rng(7);
    for (let i = 0; i < 10000; i++) {
      let code = '';
      for (let k = Math.floor(r() * 24); k > 0; k--) code += String.fromCharCode(32 + Math.floor(r() * 95));
      const compiled = compileFormat(code);
      expect(typeof compiled.ok).toBe('boolean');
      expect(typeof formatValue(randomValue(r), code).text).toBe('string');
    }
  });

  it('for extreme digit counts', () => {
    expect(formatValue(1, `0.${'0'.repeat(300)}`).text.length).toBeGreaterThan(300);
    expect(formatValue(Number.MAX_VALUE, '#,##0').text.length).toBeGreaterThan(300);
    expect(formatValue(Number.MIN_VALUE, '0.00E+00').text).toBe('4.94E-324');
    expect(formatValue(1e-300, '0.0000').text).toBe('0.0000');
  });
});

describe('General', () => {
  it('is at most 11 characters and stable when formatted again', () => {
    const r = rng(11);
    for (let i = 0; i < 20000; i++) {
      const x = (r() < 0.5 ? -1 : 1) * r() * 10 ** Math.floor(r() * 44 - 20);
      const text = formatGeneral(x);
      expect(text.length).toBeLessThanOrEqual(11 + 1); // 3-digit exponents are one character wider than the budget
      const again = formatGeneral(Number(text));
      expect(again).toBe(text);
    }
  });

  it('shows whole numbers up to 11 digits exactly', () => {
    const r = rng(12);
    for (let i = 0; i < 2000; i++) {
      const n = Math.floor(r() * 1e11);
      expect(formatGeneral(n)).toBe(String(n));
      expect(formatGeneral(-n)).toBe(n >= 1e10 ? formatGeneral(-n) : `-${n}`);
    }
  });
});

describe('section selection', () => {
  it('without conditions: one section for all, two split at zero, three add a zero section', () => {
    const r = rng(3);
    for (let i = 0; i < 3000; i++) {
      const value = pick(r, [0, -0, r() * 100, -r() * 100, 1e-9, -1e-9]);
      const sections = (n: number) => compileFormat(Array(n).fill('0').join(';')).sections;
      expect(selectSection(sections(1), value).index).toBe(0);
      expect(selectSection(sections(2), value).index).toBe(value < 0 ? 1 : 0);
      expect(selectSection(sections(3), value).index).toBe(value < 0 ? 1 : value === 0 ? 2 : 0);
      expect(selectSection(sections(4), value).index).toBe(selectSection(sections(3), value).index);
      expect(selectSection(sections(1), value).minus).toBe(value < 0);
      expect(selectSection(sections(2), value).minus).toBe(false);
    }
  });

  it('with conditions the index is always a section (or -1 for a lone unmet condition) and a minus means a negative value', () => {
    const r = rng(4);
    const ops = ['<', '>', '<=', '>=', '=', '<>'];
    for (let i = 0; i < 5000; i++) {
      const n = 1 + Math.floor(r() * 3);
      const parts = Array.from({ length: n }, (_, k) => (k < 2 && r() < 0.8 ? `[${pick(r, ops)}${Math.floor(r() * 20 - 10)}]0` : '0'));
      const code = parts.join(';');
      const value = Math.floor(r() * 40 - 20) / 2;
      const { index, minus } = selectSection(compileFormat(code).sections, value);
      expect(index).toBeGreaterThanOrEqual(n === 1 ? -1 : 0);
      expect(index).toBeLessThan(n);
      if (minus) expect(value).toBeLessThan(0);
      // the formatted text never fails and shows the magnitude when the sign is dropped
      const text = formatDisplay(value, code);
      expect(typeof text).toBe('string');
      if (!minus && value < 0 && index >= 0) expect(text.startsWith('-')).toBe(false);
    }
  });
});

describe('rounding and grouping', () => {
  it('rounds to the requested decimals: within half a unit, and formatting the result again is a no-op', () => {
    const r = rng(5);
    for (let i = 0; i < 20000; i++) {
      const places = Math.floor(r() * 7);
      const code = places ? `0.${'0'.repeat(places)}` : '0';
      const x = (r() - 0.3) * 10 ** Math.floor(r() * 9);
      const text = formatDisplay(x, code);
      const back = Number(text);
      expect(Math.abs(back - x)).toBeLessThanOrEqual(0.5 * 10 ** -places * (1 + 1e-9) + Math.abs(x) * 1e-14);
      expect(formatDisplay(back, code)).toBe(text);
      expect(text.split('.')[1]?.length ?? 0).toBe(places);
    }
  });

  it('groups thousands: removing the commas gives the plain integer, and groups have three digits', () => {
    const r = rng(6);
    for (let i = 0; i < 5000; i++) {
      const n = Math.floor(r() * 1e14);
      const text = formatDisplay(n, '#,##0');
      expect(text.replace(/,/g, '')).toBe(String(n));
      expect(text).toMatch(/^\d{1,3}(,\d{3})*$/);
    }
  });

  it('negating a number negates its text in a single-section number format', () => {
    const r = rng(8);
    for (const code of ['0.00', '#,##0', '0.0%', '0.00E+00', '"$"#,##0.00', '# ?/?']) {
      for (let i = 0; i < 400; i++) {
        const x = r() * 10 ** Math.floor(r() * 8 - 2);
        const pos = formatDisplay(x, code);
        const neg = formatDisplay(-x, code);
        expect(neg).toBe(/[1-9]/.test(pos.replace(/E[+-]\d+/i, '')) ? `-${pos}` : pos);
      }
    }
  });
});

describe('dates and times', () => {
  it('serial <-> calendar date round-trips in both date systems', () => {
    const r = rng(9);
    const serials: number[] = [];
    for (let s = 0; s <= 400; s++) serials.push(s);
    for (let i = 0; i < 20000; i++) serials.push(Math.floor(r() * 2_958_466));
    serials.push(2_958_465, 2_957_003, 1461, 1462, 36526, 36525);
    for (const s of serials) {
      const d1900 = serialToDate(s, 1900)!;
      expect(dateToSerial(d1900.year, d1900.month, d1900.day, 1900)).toBe(s);
      if (s <= 2_957_003) {
        const d1904 = serialToDate(s, 1904)!;
        expect(dateToSerial(d1904.year, d1904.month, d1904.day, 1904)).toBe(s);
      }
    }
  });

  it('the two systems differ by 1462 days from March 1900 on', () => {
    const r = rng(10);
    for (let i = 0; i < 2000; i++) {
      const s = 1462 + Math.floor(r() * 2_957_000);
      const a = serialToDate(s, 1900)!;
      const b = serialToDate(s - 1462, 1904)!;
      expect([a.year, a.month, a.day, a.weekday]).toEqual([b.year, b.month, b.day, b.weekday]);
    }
  });

  it('weekdays advance one per day and months have the right length', () => {
    let prev = serialToDate(1)!;
    for (let s = 2; s < 3000; s++) {
      const d = serialToDate(s)!;
      expect(d.weekday).toBe((prev.weekday + 1) % 7);
      prev = d;
    }
    for (const [y, m, days] of [[2026, 2, 28], [2024, 2, 29], [2000, 2, 29], [2100, 2, 28], [2026, 4, 30], [2026, 12, 31]] as const) {
      const last = dateToSerial(y, m, days)!;
      expect(serialToDate(last)!.day).toBe(days);
      expect(serialToDate(last + 1)!.day).toBe(1);
    }
  });

  it('elapsed hours, minutes and seconds read back as the same time', () => {
    const r = rng(13);
    for (let i = 0; i < 5000; i++) {
      const seconds = Math.floor(r() * 400_000_000);
      const serial = seconds / 86400;
      const [h, m, s] = formatDisplay(serial, '[h]:mm:ss').split(':').map(Number);
      expect(h * 3600 + m * 60 + s).toBe(seconds);
      expect(Number(formatDisplay(serial, '[mm]'))).toBe(Math.floor(seconds / 60));
      expect(Number(formatDisplay(serial, '[ss]'))).toBe(seconds);
    }
  });

  it('a time of day formats the same through 24-hour and AM/PM codes', () => {
    const r = rng(14);
    for (let i = 0; i < 3000; i++) {
      const serial = Math.floor(r() * 86400) / 86400;
      const [h24, m24] = formatDisplay(serial, 'hh:mm').split(':').map(Number);
      const [clock, meridiem] = formatDisplay(serial, 'hh:mm AM/PM').split(' ');
      const [h12, m12] = clock.split(':').map(Number);
      expect(m12).toBe(m24);
      expect(h12 % 12).toBe(h24 % 12);
      expect(meridiem).toBe(h24 < 12 ? 'AM' : 'PM');
    }
  });
});
