// Date functions, batch B: DAYS, DAYS360, YEARFRAC, ISOWEEKNUM, WORKDAY.INTL and NETWORKDAYS.INTL.
//
// Dates are Excel serials (1900 date system, 1900-02-29 is serial 60). Every function here lifts over arrays
// (`=DAYS(A1:A3, B1)` spills), takes the whole-day part of a date, and refuses a serial outside 0..2958465.

import type { EagerFn } from './helpers';
import { checkArity, flatten, liftScalar, toBoolean, toNum } from './helpers';
import { dateParts, toSerial, ymdToSerial } from './date';

/** 9999-12-31, the last date Excel has. */
const MAX_SERIAL = 2_958_465;

/** The whole-day serial of a date argument; a date before 1900-01-00 or after 9999-12-31 is `#NUM!`. */
function dayOf(v: unknown): number {
  const day = Math.floor(toSerial(v));
  if (day < 0 || day > MAX_SERIAL) throw new Error('#NUM!');
  return day;
}

/** Like `dayOf`, but a number keeps its time of day (DAYS of two date-times is fractional, as in LibreOffice). */
function serialOf(v: unknown): number {
  const day = dayOf(v);
  return typeof v === 'number' ? v : day;
}

/** Excel's calendar has a 29 February 1900, so 1900 is a leap year here (YEARFRAC basis 1 on 1900 dates is unverified). */
function isLeap(year: number): boolean {
  return year === 1900 || (year % 4 === 0 && year % 100 !== 0) || year % 400 === 0;
}

function lastDayOfFebruary(year: number): number {
  return isLeap(year) ? 29 : 28;
}

interface Ymd {
  year: number;
  month: number;
  day: number;
}

/** The 30/360 day count of two dates from their (possibly adjusted) days. */
function count360(a: Ymd, b: Ymd, d1: number, d2: number): number {
  return (b.year - a.year) * 360 + (b.month - a.month) * 30 + (d2 - d1);
}

const isLastOfFebruary = (d: Ymd): boolean => d.month === 2 && d.day === lastDayOfFebruary(d.year);

/** DAYS360 with method FALSE: the US (NASD) rule, last days of February included. */
function days360Us(a: Ymd, b: Ymd): number {
  let d1 = a.day;
  let d2 = b.day;
  if (isLastOfFebruary(a)) {
    if (isLastOfFebruary(b)) d2 = 30;
    d1 = 30;
  }
  if (d2 === 31 && d1 >= 30) d2 = 30;
  if (d1 === 31) d1 = 30;
  return count360(a, b, d1, d2);
}

/** DAYS360 with method TRUE, and YEARFRAC basis 4: the 31st of a month counts as the 30th. */
function days360Eu(a: Ymd, b: Ymd): number {
  return count360(a, b, Math.min(a.day, 30), Math.min(b.day, 30));
}

/**
 * YEARFRAC basis 0. Excel's rule differs from the NASD one in its handling of the last day of February (the end
 * date only moves when the start is also one), and LibreOffice and Apache POI reproduce it.
 */
function days360Basis0(a: Ymd, b: Ymd): number {
  let d1 = a.day;
  let d2 = b.day;
  if (d1 === 31 && d2 === 31) {
    d1 = 30;
    d2 = 30;
  } else if (d1 === 31) {
    d1 = 30;
  } else if (d1 === 30 && d2 === 31) {
    d2 = 30;
  } else if (isLastOfFebruary(a)) {
    if (isLastOfFebruary(b)) d2 = 30;
    d1 = 30;
  }
  return count360(a, b, d1, d2);
}

/** The year length actual/actual uses: 366 when a 29 February falls in the span (or the span is within a leap year). */
function actualYearLength(a: Ymd, b: Ymd): number {
  if (a.year === b.year) return isLeap(a.year) ? 366 : 365;
  const oneYearOrLess =
    a.year + 1 === b.year && (a.month > b.month || (a.month === b.month && a.day >= b.day));
  if (!oneYearOrLess) {
    // More than a year: the average length of the calendar years the span touches.
    let total = 0;
    for (let y = a.year; y <= b.year; y++) total += isLeap(y) ? 366 : 365;
    return total / (b.year - a.year + 1);
  }
  const feb29 = (y: number) => ymdToSerial(y, 2, 29);
  const from = ymdToSerial(a.year, a.month, a.day);
  const to = ymdToSerial(b.year, b.month, b.day);
  const covers =
    (isLeap(a.year) && from <= feb29(a.year) && feb29(a.year) <= to) ||
    (isLeap(b.year) && from <= feb29(b.year) && feb29(b.year) <= to);
  return covers ? 366 : 365;
}

/** The weekend as a table by weekday (0 = Sunday): a code of 1-7 and 11-17, or a 7-character Monday-first string. */
function weekendTable(spec: unknown): boolean[] {
  const table: boolean[] = Array(7).fill(false);
  if (spec === undefined) {
    table[6] = table[0] = true; // code 1: Saturday and Sunday
    return table;
  }
  if (typeof spec === 'string') {
    if (!/^[01]{7}$/.test(spec)) throw new Error('#VALUE!');
    for (let i = 0; i < 7; i++) table[(i + 1) % 7] = spec[i] === '1';
  } else {
    const code = Math.trunc(toNum(spec));
    if (code >= 1 && code <= 7) {
      // 1 = Saturday + Sunday, 2 = Sunday + Monday ... 7 = Friday + Saturday
      table[(code + 5) % 7] = table[(code + 6) % 7] = true;
    } else if (code >= 11 && code <= 17) {
      table[code - 11] = true; // 11 = Sunday ... 17 = Saturday
    } else {
      throw new Error('#NUM!');
    }
  }
  return table;
}

/** Weekday of a serial, 0 = Sunday (serial 1 is a Sunday in Excel's calendar). */
function weekdayOf(serial: number): number {
  return (serial + 6) % 7;
}

/** The distinct working-day holidays of a holidays argument, sorted. */
function holidayList(arg: unknown, weekend: boolean[]): number[] {
  if (arg === undefined) return [];
  const days = new Set<number>();
  for (const v of flatten([arg])) {
    if (v === null || v === '') continue; // an empty cell
    const day = Math.floor(toSerial(v));
    if (day < 0) throw new Error('#NUM!');
    if (!weekend[weekdayOf(day)]) days.add(day);
  }
  return [...days].sort((x, y) => x - y);
}

/** Index of the first element of the sorted `list` that is >= `value`. */
function lowerBound(list: number[], value: number): number {
  let lo = 0;
  let hi = list.length;
  while (lo < hi) {
    const mid = (lo + hi) >> 1;
    if (list[mid] < value) lo = mid + 1;
    else hi = mid;
  }
  return lo;
}

/** Working days in serials 0..x (x >= -1), in constant time. */
function workdaysThrough(x: number, weekend: boolean[], perWeek: number): number {
  const weeks = Math.floor((x + 1) / 7);
  let count = weeks * perWeek;
  for (let d = weeks * 7; d <= x; d++) if (!weekend[weekdayOf(d)]) count++;
  return count;
}

/** The `n`-th working day after (or, with `step` -1, before) `start`, ignoring holidays; n >= 1. */
function nthWorkday(start: number, n: number, step: 1 | -1, weekend: boolean[], perWeek: number): number {
  let day = start + 7 * step * Math.floor((n - 1) / perWeek);
  let left = ((n - 1) % perWeek) + 1;
  while (left > 0) {
    day += step;
    if (!weekend[weekdayOf(day)]) left--;
  }
  return day;
}

export const dateExtFunctions: Record<string, EagerFn> = {
  DAYS: (args) => {
    checkArity(args, 2);
    return liftScalar(args, ([end, start]) => serialOf(end) - serialOf(start));
  },

  DAYS360: (args) => {
    checkArity(args, 2, 3);
    return liftScalar(args, ([start, end, method]) => {
      const a = dateParts(dayOf(start));
      const b = dateParts(dayOf(end));
      return method !== undefined && toBoolean(method) ? days360Eu(a, b) : days360Us(a, b);
    });
  },

  YEARFRAC: (args) => {
    checkArity(args, 2, 3);
    return liftScalar(args, ([startArg, endArg, basisArg]) => {
      let start = dayOf(startArg);
      let end = dayOf(endArg);
      const basis = basisArg === undefined ? 0 : Math.trunc(toNum(basisArg));
      if (basis < 0 || basis > 4) throw new Error('#NUM!');
      if (start > end) [start, end] = [end, start];
      const a = dateParts(start);
      const b = dateParts(end);
      switch (basis) {
        case 0:
          return days360Basis0(a, b) / 360;
        case 1:
          return (end - start) / actualYearLength(a, b);
        case 2:
          return (end - start) / 360;
        case 3:
          return (end - start) / 365;
        default:
          return days360Eu(a, b) / 360;
      }
    });
  },

  ISOWEEKNUM: (args) => {
    checkArity(args, 1);
    return liftScalar(args, ([date]) => {
      const day = dayOf(date);
      const weekday = weekdayOf(day) || 7; // Monday = 1 ... Sunday = 7
      const thursday = day + 4 - weekday; // the Thursday of this ISO week decides its year
      const year = dateParts(thursday).year;
      return Math.floor((thursday - ymdToSerial(year, 1, 1)) / 7) + 1;
    });
  },

  'WORKDAY.INTL': (args) => {
    checkArity(args, 2, 4);
    return liftScalar(args.slice(0, 3), ([startArg, daysArg, weekendArg]) => {
      const start = dayOf(startArg);
      const days = Math.trunc(toNum(daysArg));
      const weekend = weekendTable(weekendArg);
      const holidays = holidayList(args[3], weekend);
      const perWeek = 7 - weekend.filter(Boolean).length;
      if (perWeek === 0) throw new Error('#VALUE!'); // "1111111": no working day at all
      if (days === 0) return start;
      const step = days > 0 ? 1 : -1;
      // Holidays push the answer out: find the day with `held` extra working days skipped, until it stops changing.
      let held = 0;
      let result = start;
      for (;;) {
        result = nthWorkday(start, Math.abs(days) + held, step, weekend, perWeek);
        if (result < 0 || result > MAX_SERIAL) throw new Error('#NUM!');
        const inSpan =
          step === 1
            ? lowerBound(holidays, result + 1) - lowerBound(holidays, start + 1)
            : lowerBound(holidays, start) - lowerBound(holidays, result);
        if (inSpan === held) return result;
        held = inSpan;
      }
    });
  },

  'NETWORKDAYS.INTL': (args) => {
    checkArity(args, 2, 4);
    return liftScalar(args.slice(0, 3), ([startArg, endArg, weekendArg]) => {
      const a = dayOf(startArg);
      const b = dayOf(endArg);
      const weekend = weekendTable(weekendArg);
      const holidays = holidayList(args[3], weekend);
      const from = Math.min(a, b);
      const to = Math.max(a, b);
      const perWeek = 7 - weekend.filter(Boolean).length;
      const held = lowerBound(holidays, to + 1) - lowerBound(holidays, from);
      const count = workdaysThrough(to, weekend, perWeek) - workdaysThrough(from - 1, weekend, perWeek) - held;
      return a > b && count !== 0 ? -count : count; // a weekend of "1111111" has no working days: 0
    });
  },
};
