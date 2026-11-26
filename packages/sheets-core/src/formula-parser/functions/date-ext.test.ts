import type { EvaluationContext } from '../types';
import { dateExtFunctions } from './date-ext';
import { dateFunctions } from './date';

const ctx: EvaluationContext = { getCellValue: () => null, getRangeValues: () => [] };

function call(name: string, ...args: unknown[]): unknown {
  return dateExtFunctions[name](args, ctx);
}

function err(name: string, ...args: unknown[]): string {
  try {
    call(name, ...args);
  } catch (e) {
    return (e as Error).message;
  }
  return 'no error';
}

/** Serial of a date by Excel's calendar, through the engine's own DATE. */
const d = (y: number, m: number, day: number): number => dateFunctions.DATE([y, m, day], ctx) as number;

describe('the 1900 calendar the date functions share', () => {
  it('has a 29 February 1900 at serial 60', () => {
    expect(d(1900, 1, 1)).toBe(1);
    expect(d(1900, 2, 28)).toBe(59);
    expect(d(1900, 2, 29)).toBe(60);
    expect(d(1900, 3, 1)).toBe(61);
    expect(dateFunctions.DAY([60], ctx)).toBe(29);
    expect(dateFunctions.MONTH([60], ctx)).toBe(2);
    expect(dateFunctions.YEAR([1], ctx)).toBe(1900);
    expect(dateFunctions.DAY([1], ctx)).toBe(1);
    expect(dateFunctions.WEEKDAY([1], ctx)).toBe(1); // Excel says Sunday
    expect(dateFunctions.WEEKDAY([61], ctx)).toBe(5);
  });

  it('rolls day and month overflow over in the 1900 calendar too', () => {
    expect(d(1900, 2, 30)).toBe(61);
    expect(d(1900, 3, 0)).toBe(60);
    expect(d(1900, 1, 60)).toBe(60);
    expect(d(1900, 1, 61)).toBe(61);
    expect(d(1900, 4, 0)).toBe(91);
    expect(d(1900, 14, 1)).toBe(d(1901, 2, 1));
    expect(d(1901, -11, 1)).toBe(d(1900, 1, 1));
    expect(dateFunctions.EOMONTH([1, 1], ctx)).toBe(60);
    expect(dateFunctions.EOMONTH([1, 0], ctx)).toBe(31);
    expect(dateFunctions.EOMONTH([61, 0], ctx)).toBe(91);
    expect(dateFunctions.YEAR([0], ctx)).toBe(1900);
    expect(dateFunctions.MONTH([0], ctx)).toBe(1);
    expect(dateFunctions.DAY([0], ctx)).toBe(0);
  });

  it('falls back to the plain calendar outside Excel dates', () => {
    expect(d(10000, 1, 1)).toBe(2958466);
    expect(dateFunctions.YEAR([-1], ctx)).toBe(1899);
    expect(dateFunctions.YEAR([-5], ctx)).toBe(1899);
  });
});

describe('DAYS', () => {
  it('subtracts the start from the end', () => {
    expect(call('DAYS', d(2021, 3, 1), d(2021, 1, 1))).toBe(59);
    expect(call('DAYS', d(2021, 1, 1), d(2021, 3, 1))).toBe(-59);
    expect(call('DAYS', d(2024, 3, 1), d(2024, 2, 1))).toBe(29);
    expect(call('DAYS', '2021-03-01', '2021-01-01')).toBe(59);
    expect(call('DAYS', '2021-03-01 18:00', '2021-01-01 06:00')).toBe(59);
    expect(call('DAYS', null, 5)).toBe(-5);
  });

  it('keeps the time of day of numbers', () => {
    expect(call('DAYS', 44261.9, 44201.1)).toBeCloseTo(60.8, 9);
  });

  it('counts the non-existent 29 February 1900', () => {
    expect(call('DAYS', 61, 59)).toBe(2);
    expect(call('DAYS', d(1900, 3, 1), d(1900, 2, 28))).toBe(2);
    expect(call('DAYS', d(1900, 3, 1), d(1900, 2, 29))).toBe(1);
  });

  it('rejects a date out of range, text that is no date and a wrong argument count', () => {
    expect(err('DAYS', -1, 5)).toBe('#NUM!');
    expect(err('DAYS', 5, 2958466)).toBe('#NUM!');
    expect(err('DAYS', 5, 'x')).toBe('#VALUE!');
    expect(err('DAYS', '#N/A', 1)).toBe('#N/A');
    expect(err('DAYS', 1)).toBe('#VALUE!');
    expect(err('DAYS', 1, 2, 3)).toBe('#VALUE!');
  });

  it('lifts over arrays', () => {
    expect(call('DAYS', [[44200], [44300]], 44100)).toEqual([[100], [200]]);
    expect(call('DAYS', [[44200], [-1]], 44100)).toEqual([[100], ['#NUM!']]);
  });
});

describe('DAYS360', () => {
  it('counts 30/360 days by the US method', () => {
    expect(call('DAYS360', d(2021, 1, 31), d(2021, 2, 28))).toBe(28);
    expect(call('DAYS360', d(2020, 1, 30), d(2020, 3, 31))).toBe(60);
    expect(call('DAYS360', d(2021, 1, 15), d(2021, 2, 28))).toBe(43);
    expect(call('DAYS360', d(2021, 1, 1), d(2022, 1, 1))).toBe(360);
    expect(call('DAYS360', d(2021, 3, 31), d(2021, 1, 31))).toBe(-60);
    expect(call('DAYS360', d(2021, 1, 31), d(2021, 3, 31))).toBe(60);
    expect(call('DAYS360', d(2021, 1, 31), d(2021, 3, 31), false)).toBe(60);
    expect(call('DAYS360', d(2021, 1, 15), d(2021, 3, 31))).toBe(76);
  });

  it('makes the last day of February the 30th', () => {
    expect(call('DAYS360', d(2021, 2, 28), d(2021, 3, 31))).toBe(30);
    expect(call('DAYS360', d(2020, 2, 29), d(2020, 3, 31))).toBe(30);
    expect(call('DAYS360', d(2021, 2, 28), d(2021, 2, 28))).toBe(0);
    expect(call('DAYS360', d(2020, 2, 29), d(2021, 2, 28))).toBe(360);
    expect(call('DAYS360', d(2020, 2, 28), d(2020, 3, 28))).toBe(30); // 28 February 2020 is not the last day
    expect(call('DAYS360', d(1900, 2, 28), d(1900, 3, 1))).toBe(3); // in Excel's 1900 February has 29 days
    expect(call('DAYS360', d(1900, 2, 29), d(1900, 3, 1))).toBe(1);
  });

  it('counts by the European method with a TRUE method', () => {
    expect(call('DAYS360', d(2021, 1, 31), d(2021, 2, 28), true)).toBe(28);
    expect(call('DAYS360', d(2021, 2, 28), d(2021, 3, 31), true)).toBe(32);
    expect(call('DAYS360', d(2021, 2, 28), d(2021, 2, 28), true)).toBe(0);
    expect(call('DAYS360', d(2021, 1, 1), d(2021, 12, 31), 1)).toBe(359);
    expect(call('DAYS360', d(2021, 1, 1), d(2021, 12, 31), 0)).toBe(360);
  });

  it('checks its arguments and lifts over arrays', () => {
    expect(err('DAYS360', d(2021, 1, 1), 'x')).toBe('#VALUE!');
    expect(err('DAYS360', -1, 5)).toBe('#NUM!');
    expect(err('DAYS360', 1)).toBe('#VALUE!');
    expect(err('DAYS360', 1, 2, true, 4)).toBe('#VALUE!');
    expect(call('DAYS360', [[44200], [44300]], d(2021, 12, 31))).toEqual([[357], [257]]);
  });
});

describe('YEARFRAC', () => {
  it('uses 30/360 (US) by default and for basis 0', () => {
    expect(call('YEARFRAC', d(2021, 1, 1), d(2021, 7, 1))).toBe(0.5);
    expect(call('YEARFRAC', d(2021, 1, 1), d(2021, 7, 1), 0)).toBe(0.5);
    expect(call('YEARFRAC', d(2021, 1, 31), d(2021, 2, 28), 0)).toBeCloseTo(28 / 360, 12);
    expect(call('YEARFRAC', d(2021, 2, 28), d(2021, 3, 31), 0)).toBeCloseTo(31 / 360, 12); // Excel's quirk: not 30
    expect(call('YEARFRAC', d(2021, 1, 31), d(2021, 1, 31), 0)).toBe(0);
    expect(call('YEARFRAC', d(2021, 1, 31), d(2021, 3, 31), 0)).toBeCloseTo(60 / 360, 12);
    expect(call('YEARFRAC', d(2021, 1, 30), d(2021, 3, 31), 0)).toBeCloseTo(60 / 360, 12);
    expect(call('YEARFRAC', d(2020, 2, 29), d(2021, 2, 28), 0)).toBe(1);
    expect(call('YEARFRAC', d(2021, 2, 28), d(2021, 2, 28), 0)).toBe(0);
  });

  it('counts actual/actual for basis 1', () => {
    expect(call('YEARFRAC', d(2020, 3, 1), d(2020, 6, 1), 1)).toBeCloseTo(92 / 366, 12); // a leap year
    expect(call('YEARFRAC', d(2021, 3, 1), d(2021, 6, 1), 1)).toBeCloseTo(92 / 365, 12);
    expect(call('YEARFRAC', d(2019, 3, 1), d(2020, 6, 1), 1)).toBeCloseTo(458 / 365.5, 12); // more than a year: the average of the years it touches
    expect(call('YEARFRAC', d(2020, 3, 1), d(2021, 2, 28), 1)).toBeCloseTo(364 / 365, 12); // within a year, no 29 February
    expect(call('YEARFRAC', d(2019, 2, 1), d(2021, 5, 1), 1)).toBeCloseTo(820 / ((365 + 366 + 365) / 3), 12);
    expect(call('YEARFRAC', d(2020, 1, 15), d(2021, 1, 10), 1)).toBeCloseTo(361 / 366, 12); // 29 Feb 2020 inside
    expect(call('YEARFRAC', d(2019, 12, 1), d(2020, 11, 30), 1)).toBeCloseTo(365 / 366, 12);
    expect(call('YEARFRAC', d(2019, 1, 1), d(2020, 1, 1), 1)).toBeCloseTo(365 / 365, 12); // exactly a year, from a common year
    expect(call('YEARFRAC', d(2019, 4, 1), d(2020, 3, 1), 1)).toBeCloseTo(335 / 366, 12);
    expect(call('YEARFRAC', d(2021, 6, 1), d(2021, 6, 1), 1)).toBe(0);
    expect(call('YEARFRAC', d(1900, 1, 1), d(1900, 12, 31), 1)).toBeCloseTo(365 / 366, 12); // 1900 is a leap year in Excel's calendar (unverified)
    expect(call('YEARFRAC', d(2020, 3, 1), d(2021, 2, 1), 1)).toBeCloseTo(337 / 365, 12);
    expect(call('YEARFRAC', d(2023, 6, 1), d(2024, 7, 1), 1)).toBeCloseTo(396 / ((365 + 366) / 2), 12);
  });

  it('counts actual/360, actual/365 and the European 30/360', () => {
    expect(call('YEARFRAC', d(2021, 1, 1), d(2021, 3, 1), 2)).toBeCloseTo(59 / 360, 12);
    expect(call('YEARFRAC', d(2021, 1, 1), d(2021, 3, 1), 3)).toBeCloseTo(59 / 365, 12);
    expect(call('YEARFRAC', d(2021, 1, 31), d(2021, 3, 31), 4)).toBeCloseTo(60 / 360, 12);
    expect(call('YEARFRAC', d(2021, 3, 1), d(2021, 1, 1), 2)).toBeCloseTo(59 / 360, 12); // the dates swap
    expect(call('YEARFRAC', d(2021, 1, 1), d(2021, 3, 1), 2.9)).toBeCloseTo(59 / 360, 12);
  });

  it('rejects a basis outside 0-4, bad dates and a wrong argument count', () => {
    expect(err('YEARFRAC', 100, 200, 5)).toBe('#NUM!');
    expect(err('YEARFRAC', 100, 200, -1)).toBe('#NUM!');
    expect(err('YEARFRAC', -1, 200)).toBe('#NUM!');
    expect(err('YEARFRAC', 100, 'x')).toBe('#VALUE!');
    expect(err('YEARFRAC', 100)).toBe('#VALUE!');
    expect(err('YEARFRAC', 100, 200, 1, 1)).toBe('#VALUE!');
    expect(call('YEARFRAC', [[44197], [44378]], 44562)).toEqual([[1], [0.5]]);
  });
});

describe('ISOWEEKNUM', () => {
  it('numbers the ISO weeks, which can reach into the neighbouring year', () => {
    expect(call('ISOWEEKNUM', d(2021, 1, 3))).toBe(53);
    expect(call('ISOWEEKNUM', d(2021, 1, 4))).toBe(1);
    expect(call('ISOWEEKNUM', d(2024, 12, 30))).toBe(1);
    expect(call('ISOWEEKNUM', d(2020, 12, 31))).toBe(53);
    expect(call('ISOWEEKNUM', d(2026, 1, 1))).toBe(1);
    expect(call('ISOWEEKNUM', d(2023, 1, 1))).toBe(52);
    expect(call('ISOWEEKNUM', d(2015, 12, 31))).toBe(53);
    expect(call('ISOWEEKNUM', d(2022, 12, 31))).toBe(52);
    expect(call('ISOWEEKNUM', d(2021, 6, 15) + 0.75)).toBe(24);
  });

  it('works at the start of the 1900 calendar', () => {
    expect(call('ISOWEEKNUM', 61)).toBe(9);
    expect(call('ISOWEEKNUM', 1)).toBe(52); // Excel's Sunday 1900-01-01 belongs to the last week of the previous year
  });

  it('checks its argument', () => {
    expect(err('ISOWEEKNUM', -1)).toBe('#NUM!');
    expect(err('ISOWEEKNUM', 'x')).toBe('#VALUE!');
    expect(err('ISOWEEKNUM')).toBe('#VALUE!');
    expect(call('ISOWEEKNUM', [[44200], [44300]])).toEqual([[1], [15]]);
  });
});

/** WORKDAY.INTL the slow way: step one day at a time (`weekend` lists the weekend weekdays, 0 = Sunday). */
function walk(start: number, days: number, weekend: number[], holidays: number[] = []): number {
  let day = start;
  const step = days < 0 ? -1 : 1;
  for (let left = Math.abs(days); left > 0; ) {
    day += step;
    if (!weekend.includes((day + 6) % 7) && !holidays.includes(day)) left--;
  }
  return day;
}

const SAT_SUN = [6, 0];

describe('WORKDAY.INTL', () => {
  it('skips Saturday and Sunday by default', () => {
    expect(call('WORKDAY.INTL', d(2021, 1, 1), 10)).toBe(d(2021, 1, 15));
    expect(call('WORKDAY.INTL', d(2021, 1, 1), 1)).toBe(d(2021, 1, 4));
    expect(call('WORKDAY.INTL', d(2021, 1, 15), -5)).toBe(d(2021, 1, 8));
    expect(call('WORKDAY.INTL', d(2021, 1, 16), 0)).toBe(d(2021, 1, 16)); // a weekend start stays put
    expect(call('WORKDAY.INTL', d(2021, 1, 1), 5.9)).toBe(d(2021, 1, 8)); // days are truncated
    expect(call('WORKDAY.INTL', d(2021, 1, 1), 0, undefined)).toBe(d(2021, 1, 1));
  });

  it('takes a weekend code or a seven-character string', () => {
    const start = d(2021, 1, 1);
    const weekends: Array<[unknown, number[]]> = [
      [1, [6, 0]],
      [2, [0, 1]],
      [3, [1, 2]],
      [4, [2, 3]],
      [5, [3, 4]],
      [6, [4, 5]],
      [7, [5, 6]],
      [11, [0]],
      [12, [1]],
      [13, [2]],
      [14, [3]],
      [15, [4]],
      [16, [5]],
      [17, [6]],
      [1.9, [6, 0]],
      ['0000011', [6, 0]],
      ['1000001', [1, 0]],
      ['0111111', [2, 3, 4, 5, 6, 0]],
    ];
    for (const [code, days] of weekends) {
      for (const n of [1, 3, 10, -4, 25]) expect(call('WORKDAY.INTL', start, n, code)).toBe(walk(start, n, days));
    }
    expect(call('WORKDAY.INTL', d(2021, 1, 1), 10, 11)).toBe(d(2021, 1, 13)); // Sunday only
    expect(call('WORKDAY.INTL', d(2021, 1, 1), 3, '0111111')).toBe(d(2021, 1, 18)); // only Mondays work
  });

  it('moves past holidays, ignoring duplicates, weekend days and blanks', () => {
    const hol = [[44207, 44208, null, '', 44214]]; // 11, 12 and 18 January 2021
    expect(call('WORKDAY.INTL', d(2021, 1, 1), 10, 1, hol)).toBe(walk(d(2021, 1, 1), 10, SAT_SUN, [44207, 44208, 44214]));
    expect(call('WORKDAY.INTL', d(2021, 1, 1), 10, undefined, hol)).toBe(d(2021, 1, 20));
    expect(call('WORKDAY.INTL', d(2021, 1, 20), -5, 1, [44207, 44208, 44214])).toBe(walk(44216, -5, SAT_SUN, [44207, 44208, 44214]));
    expect(call('WORKDAY.INTL', d(2021, 1, 1), 1, 1, [44200, 44200, 44201])).toBe(d(2021, 1, 6)); // 4 and 5 January, 4 twice
    expect(call('WORKDAY.INTL', d(2021, 1, 1), 1, 1, [44198])).toBe(d(2021, 1, 4)); // a Saturday holiday changes nothing
    expect(call('WORKDAY.INTL', d(2021, 1, 4), 1, 1, d(2021, 1, 5))).toBe(d(2021, 1, 6));
  });

  it('moves past a run of holidays in both directions', () => {
    const week = [44201, 44202, 44203, 44204, 44205, 44206, 44207]; // 5-11 January 2021
    expect(call('WORKDAY.INTL', d(2021, 1, 4), 1, 1, week)).toBe(d(2021, 1, 12));
    expect(call('WORKDAY.INTL', d(2021, 1, 12), -1, 1, week)).toBe(d(2021, 1, 4));
    expect(call('WORKDAY.INTL', d(2021, 1, 12), -2, 1, week)).toBe(d(2021, 1, 1));
  });

  it('agrees with stepping day by day for many starts, lengths, weekends and holidays', () => {
    let seed = 12345;
    const next = (n: number): number => {
      seed = (seed * 1103515245 + 12345) % 2147483648;
      return seed % n;
    };
    const codes = [1, 2, 7, 11, 17, '0000011', '1010100', '0111111'];
    const tables = [[6, 0], [0, 1], [5, 6], [0], [6], [6, 0], [1, 3, 5], [2, 3, 4, 5, 6, 0]];
    for (let i = 0; i < 300; i++) {
      const k = next(codes.length);
      const start = 40000 + next(2000);
      const days = next(120) - 60;
      const holidays = Array.from({ length: next(8) }, () => start + next(100) - 50);
      expect(call('WORKDAY.INTL', start, days, codes[k], holidays)).toBe(walk(start, days, tables[k], holidays));
    }
  });

  it('rejects bad weekends, dates out of range, text holidays and arity', () => {
    expect(err('WORKDAY.INTL', d(2021, 1, 1), 10, '1111111')).toBe('#VALUE!');
    expect(err('WORKDAY.INTL', d(2021, 1, 1), 10, '01')).toBe('#VALUE!');
    expect(err('WORKDAY.INTL', d(2021, 1, 1), 10, '0000012')).toBe('#VALUE!');
    expect(err('WORKDAY.INTL', d(2021, 1, 1), 10, 8)).toBe('#NUM!');
    expect(err('WORKDAY.INTL', d(2021, 1, 1), 10, 0)).toBe('#NUM!');
    expect(err('WORKDAY.INTL', d(2021, 1, 1), 10, 18)).toBe('#NUM!');
    expect(err('WORKDAY.INTL', d(2021, 1, 1), 10, 'x')).toBe('#VALUE!');
    expect(err('WORKDAY.INTL', d(2021, 1, 1), 10, null)).toBe('#NUM!');
    expect(err('WORKDAY.INTL', 10, -100)).toBe('#NUM!');
    expect(err('WORKDAY.INTL', d(9999, 12, 30), 10)).toBe('#NUM!');
    expect(err('WORKDAY.INTL', d(2021, 1, 1), 10, 1, ['x'])).toBe('#VALUE!');
    expect(err('WORKDAY.INTL', d(2021, 1, 1), 10, 1, [-5])).toBe('#NUM!');
    expect(err('WORKDAY.INTL', d(2021, 1, 1), 'x')).toBe('#VALUE!');
    expect(err('WORKDAY.INTL', 1)).toBe('#VALUE!');
    expect(err('WORKDAY.INTL', 1, 2, 3, 4, 5)).toBe('#VALUE!');
  });

  it('is constant time for a huge number of days and lifts over arrays', () => {
    const started = Date.now();
    expect(call('WORKDAY.INTL', d(2021, 1, 1), 1_000_000)).toBe(d(2021, 1, 1) + 1_400_000);
    expect(err('WORKDAY.INTL', d(2021, 1, 1), 100_000_000)).toBe('#NUM!');
    expect(err('WORKDAY.INTL', d(2021, 1, 1), -100_000_000)).toBe('#NUM!');
    expect(Date.now() - started).toBeLessThan(500);
    expect(call('WORKDAY.INTL', [[44200], [44300]], 1, 17)).toEqual([[44201], [44301]]);
  });
});

describe('NETWORKDAYS.INTL', () => {
  it('counts the working days, both ends included', () => {
    expect(call('NETWORKDAYS.INTL', d(2021, 1, 1), d(2021, 1, 31))).toBe(21);
    expect(call('NETWORKDAYS.INTL', d(2021, 1, 31), d(2021, 1, 1))).toBe(-21);
    expect(call('NETWORKDAYS.INTL', d(2021, 1, 2), d(2021, 1, 3))).toBe(0);
    expect(call('NETWORKDAYS.INTL', d(2021, 1, 4), d(2021, 1, 4))).toBe(1);
    expect(call('NETWORKDAYS.INTL', d(2021, 1, 1), d(2021, 1, 31), 11)).toBe(26);
    expect(call('NETWORKDAYS.INTL', d(2021, 1, 1), d(2021, 1, 31), 7)).toBe(21);
    expect(call('NETWORKDAYS.INTL', d(2021, 1, 1), d(2021, 1, 31), '0000011')).toBe(21);
    expect(call('NETWORKDAYS.INTL', d(2021, 1, 1), d(2021, 1, 31), '1010101')).toBe(13);
    expect(call('NETWORKDAYS.INTL', 1, 61)).toBe(call('NETWORKDAYS.INTL', 0, 61) as number);
  });

  it('subtracts holidays that fall on working days, once each', () => {
    const hol = [[44207, 44208, 44207, 44213]]; // 11, 12, 11 and 17 January 2021 (a Sunday)... 44213 is a Sunday
    expect(call('NETWORKDAYS.INTL', d(2021, 1, 1), d(2021, 1, 31), 1, hol)).toBe(19);
    expect(call('NETWORKDAYS.INTL', d(2021, 1, 1), d(2021, 1, 31), undefined, hol)).toBe(19);
    expect(call('NETWORKDAYS.INTL', d(2021, 1, 1), d(2021, 1, 10), 1, [44207])).toBe(6); // the holiday is outside the span
    expect(call('NETWORKDAYS.INTL', d(2021, 1, 31), d(2021, 1, 1), 1, [44207])).toBe(-20);
  });

  it('is 0 when every day is a weekend day', () => {
    expect(call('NETWORKDAYS.INTL', d(2021, 1, 1), d(2021, 1, 31), '1111111')).toBe(0);
    expect(call('NETWORKDAYS.INTL', d(2021, 1, 31), d(2021, 1, 1), '1111111', [44207])).toBe(0);
  });

  it('rejects bad weekends, dates and arity', () => {
    expect(err('NETWORKDAYS.INTL', 1, 2, 0)).toBe('#NUM!');
    expect(err('NETWORKDAYS.INTL', 1, 2, '12')).toBe('#VALUE!');
    expect(err('NETWORKDAYS.INTL', 1, 'x')).toBe('#VALUE!');
    expect(err('NETWORKDAYS.INTL', -1, 2)).toBe('#NUM!');
    expect(err('NETWORKDAYS.INTL', 1, 2, 1, ['x'])).toBe('#VALUE!');
    expect(err('NETWORKDAYS.INTL', 1)).toBe('#VALUE!');
    expect(call('NETWORKDAYS.INTL', [[44200], [44300]], 44400, 11)).toEqual([[call('NETWORKDAYS.INTL', 44200, 44400, 11)], [call('NETWORKDAYS.INTL', 44300, 44400, 11)]]);
  });

  it('is constant time over the whole date range', () => {
    const started = Date.now();
    const days = call('NETWORKDAYS.INTL', 0, 2958465) as number;
    expect(days).toBeGreaterThan(2_000_000);
    expect(days).toBeLessThan(2_200_000);
    expect(Date.now() - started).toBeLessThan(200);
  });
});

describe('weekend codes agree with day-by-day counting', () => {
  it('for every code and a span of a few years', () => {
    const codes: unknown[] = [1, 2, 3, 4, 5, 6, 7, 11, 12, 13, 14, 15, 16, 17, '0101010', '1000000', '0000001'];
    for (const code of codes) {
      for (const [from, to] of [
        [44000, 44017],
        [44001, 44001],
        [44100, 45200],
      ]) {
        const walked: number[] = [];
        let cursor = from;
        for (;;) {
          const next = call('WORKDAY.INTL', cursor, 1, code) as number;
          if (next > to) break;
          walked.push(next);
          cursor = next;
        }
        const first = call('WORKDAY.INTL', from, 0, code) as number;
        const startCounts = call('NETWORKDAYS.INTL', first, first, code) as number;
        expect((call('NETWORKDAYS.INTL', from, to, code) as number) - startCounts).toBe(walked.length);
      }
    }
  });
});

describe('EDATE, EOMONTH and DATEDIF in the 1900 calendar', () => {
  const f = (name: string, ...args: unknown[]): unknown => dateFunctions[name](args, ctx);

  it('EDATE keeps 29 February 1900 and lands on it from month ends', () => {
    expect(f('EDATE', 60, 0)).toBe(60);
    expect(f('EDATE', 31, 1)).toBe(60); // 30 January + 1 month: the last day of Excel's 29-day February
    expect(f('EDATE', 61, -1)).toBe(32);
    expect(f('EDATE', 59, 1)).toBe(d(1900, 3, 28));
    expect(f('EDATE', d(2020, 2, 29), 12)).toBe(d(2021, 2, 28));
    expect(f('EDATE', d(2021, 1, 31), 1)).toBe(d(2021, 2, 28));
    expect(f('EOMONTH', 60, 0)).toBe(60);
    expect(f('EOMONTH', 32, 0)).toBe(60);
  });

  it('DATEDIF counts February 1900 as 29 days', () => {
    expect(f('DATEDIF', 1, 61, 'D')).toBe(60);
    expect(f('DATEDIF', 59, 91, 'MD')).toBe(3); // 28 February to 31 March
    expect(f('DATEDIF', 60, 90, 'MD')).toBe(1); // 29 to 30 March
    expect(f('DATEDIF', d(2020, 12, 20), d(2021, 1, 10), 'MD')).toBe(21); // the month before January is December: 31 days
    expect(f('DATEDIF', d(2021, 1, 20), d(2021, 3, 10), 'MD')).toBe(18); // February 2021: 28 days
    expect(f('DATEDIF', 60, 61, 'YD')).toBe(1);
    expect(f('DATEDIF', d(2020, 3, 1), d(2021, 2, 28), 'YD')).toBe(364);
    expect(f('DATEDIF', d(2020, 3, 1), d(2021, 3, 1), 'YD')).toBe(0);
  });
});
