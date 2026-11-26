// Date and time functions.
//
// Dates are Excel-style serial numbers: day 0 is 1899-12-30, so 1900-03-01
// onward matches Excel exactly. The fractional part is the time of day
// (0.5 = noon). All calendar maths is done in UTC to stay DST-proof.

import type { EagerFn } from './helpers';
import { flatten, toNum, toText, isErrorValue } from './helpers';
import { dateToSerial, serialToDate } from '../../number-format/dates';

const EPOCH = Date.UTC(1899, 11, 30);
const MS_PER_DAY = 86_400_000;

/**
 * Serial number for a year/month/day (month and day overflow roll over). Follows Excel's calendar, which has a
 * 29 February 1900 (serial 60): 1900-01-01 is 1, `DATE(1900,2,30)` is 61 and `DATE(1900,3,0)` is 60. A date it
 * cannot place (before 1900, after 9999) falls back to the plain calendar.
 */
export function ymdToSerial(year: number, month: number, day: number): number {
  const m0 = month - 1;
  const y = year + Math.floor(m0 / 12);
  const m = (((m0 % 12) + 12) % 12) + 1;
  // the serial of the first of the month, then the days on from it (this is what makes the overflow follow Excel's 1900)
  const first = y === 1900 && m < 3 ? (m === 1 ? 1 : 32) : dateToSerial(y, m, 1);
  return first === null ? (Date.UTC(year, month - 1, day) - EPOCH) / MS_PER_DAY : first + day - 1;
}

/** Calendar parts of the whole-day portion of a serial (Excel's 1900 calendar: serial 60 is 1900-02-29, 0 is 1900-01-00). */
export function dateParts(serial: number): {
  year: number;
  month: number;
  day: number;
  weekday: number;
} {
  const excel = serialToDate(serial);
  if (excel) return excel;
  const d = new Date(EPOCH + Math.floor(serial) * MS_PER_DAY);
  return {
    year: d.getUTCFullYear(),
    month: d.getUTCMonth() + 1,
    day: d.getUTCDate(),
    weekday: d.getUTCDay(),
  };
}

/** Clock parts of the fractional portion of a serial. */
function timeParts(serial: number): { hour: number; minute: number; second: number } {
  let total = Math.round((serial - Math.floor(serial)) * 86_400);
  if (total >= 86_400) total -= 86_400;
  return {
    hour: Math.floor(total / 3600),
    minute: Math.floor((total % 3600) / 60),
    second: total % 60,
  };
}

/** Parse a date string (ISO or M/D/Y), optionally with a trailing time. */
function parseDateString(text: string): number | null {
  const trimmed = text.trim();
  const spaceIdx = trimmed.indexOf(' ');
  const datePart = spaceIdx >= 0 ? trimmed.slice(0, spaceIdx) : trimmed;
  const timePart = spaceIdx >= 0 ? trimmed.slice(spaceIdx + 1) : '';

  let year: number;
  let month: number;
  let day: number;
  let m = datePart.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
  if (m) {
    year = +m[1];
    month = +m[2];
    day = +m[3];
  } else if ((m = datePart.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/))) {
    month = +m[1];
    day = +m[2];
    year = +m[3];
  } else {
    return null;
  }
  if (month < 1 || month > 12 || day < 1 || day > 31) return null;

  let serial = ymdToSerial(year, month, day);
  if (timePart) {
    const t = parseTimeString(timePart);
    if (t !== null) serial += t;
  }
  return serial;
}

/** Parse a time string (`HH:MM`, `HH:MM:SS`, optional AM/PM) to a day fraction. */
function parseTimeString(text: string): number | null {
  const m = text.trim().match(/^(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?\s*([AaPp][Mm])?$/);
  if (!m) return null;
  let hour = +m[1];
  const minute = +m[2];
  const second = m[3] ? +m[3] : 0;
  const meridiem = m[4]?.toUpperCase();
  if (meridiem === 'PM' && hour < 12) hour += 12;
  if (meridiem === 'AM' && hour === 12) hour = 0;
  if (minute > 59 || second > 59) return null;
  return ((hour * 3600 + minute * 60 + second) % 86_400) / 86_400;
}

/** Coerce a value to a date serial, parsing date/time strings when needed. */
export function toSerial(v: unknown): number {
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (v == null || v === '') return 0;
  if (typeof v === 'string') {
    if (isErrorValue(v)) throw new Error(v);
    const t = v.trim();
    if (t !== '' && !isNaN(Number(t))) return Number(t);
    const asDate = parseDateString(t);
    if (asDate !== null) return asDate;
    const asTime = parseTimeString(t);
    if (asTime !== null) return asTime;
  }
  throw new Error('#VALUE!');
}

/** Days in a month of Excel's calendar (February 1900 has 29). */
function daysInMonth(year: number, month: number): number {
  return ymdToSerial(year, month + 1, 1) - ymdToSerial(year, month, 1);
}

/** Number of weekday positions to shift for a WEEKNUM start-day type. */
function weeknumStartDay(type: number): number {
  if (type === 2 || type === 11) return 1;
  if (type >= 12 && type <= 16) return type - 10;
  return 0; // type 1, 17 and the default: weeks start on Sunday
}

export const dateFunctions: Record<string, EagerFn> = {
  TODAY: () => {
    const now = new Date();
    return ymdToSerial(now.getFullYear(), now.getMonth() + 1, now.getDate());
  },

  NOW: () => {
    const now = new Date();
    const day = ymdToSerial(now.getFullYear(), now.getMonth() + 1, now.getDate());
    const time = (now.getHours() * 3600 + now.getMinutes() * 60 + now.getSeconds()) / 86_400;
    return day + time;
  },

  DATE: (args) => {
    let year = Math.trunc(toNum(args[0]));
    const month = Math.trunc(toNum(args[1]));
    const day = Math.trunc(toNum(args[2]));
    if (year >= 0 && year < 1900) year += 1900;
    return ymdToSerial(year, month, day);
  },

  TIME: (args) => {
    const total = toNum(args[0]) * 3600 + toNum(args[1]) * 60 + toNum(args[2]);
    if (total < 0) throw new Error('#NUM!');
    return (total % 86_400) / 86_400;
  },

  YEAR: (args) => dateParts(toSerial(args[0])).year,
  MONTH: (args) => dateParts(toSerial(args[0])).month,
  DAY: (args) => dateParts(toSerial(args[0])).day,
  HOUR: (args) => timeParts(toSerial(args[0])).hour,
  MINUTE: (args) => timeParts(toSerial(args[0])).minute,
  SECOND: (args) => timeParts(toSerial(args[0])).second,

  WEEKDAY: (args) => {
    const wd = dateParts(toSerial(args[0])).weekday;
    const type = args[1] !== undefined ? Math.trunc(toNum(args[1])) : 1;
    if (type === 1) return wd + 1;
    if (type === 2) return ((wd + 6) % 7) + 1;
    if (type === 3) return (wd + 6) % 7;
    if (type >= 11 && type <= 17) {
      return ((wd - ((type - 10) % 7) + 7) % 7) + 1;
    }
    throw new Error('#NUM!');
  },

  WEEKNUM: (args) => {
    const serial = Math.floor(toSerial(args[0]));
    const type = args[1] !== undefined ? Math.trunc(toNum(args[1])) : 1;
    const startDay = weeknumStartDay(type);
    const { year } = dateParts(serial);
    const jan1 = ymdToSerial(year, 1, 1);
    const jan1Offset = (dateParts(jan1).weekday - startDay + 7) % 7;
    return Math.floor((serial - jan1 + jan1Offset) / 7) + 1;
  },

  EDATE: (args) => {
    const { year, month, day } = dateParts(toSerial(args[0]));
    const total = year * 12 + (month - 1) + Math.trunc(toNum(args[1]));
    const targetYear = Math.floor(total / 12);
    const targetMonth = total % 12;
    return ymdToSerial(targetYear, targetMonth + 1, Math.min(day, daysInMonth(targetYear, targetMonth + 1)));
  },

  EOMONTH: (args) => {
    const { year, month } = dateParts(toSerial(args[0]));
    const total = year * 12 + (month - 1) + Math.trunc(toNum(args[1]));
    const targetYear = Math.floor(total / 12);
    const targetMonth = total % 12;
    return ymdToSerial(targetYear, targetMonth + 2, 0);
  },

  DATEDIF: (args) => {
    const start = Math.floor(toSerial(args[0]));
    const end = Math.floor(toSerial(args[1]));
    const unit = toText(args[2]).toUpperCase();
    if (end < start) throw new Error('#NUM!');
    const s = dateParts(start);
    const e = dateParts(end);

    switch (unit) {
      case 'D':
        return end - start;
      case 'Y': {
        let years = e.year - s.year;
        if (e.month < s.month || (e.month === s.month && e.day < s.day)) years--;
        return years;
      }
      case 'M': {
        let months = (e.year - s.year) * 12 + (e.month - s.month);
        if (e.day < s.day) months--;
        return months;
      }
      case 'MD': {
        let days = e.day - s.day;
        if (days < 0) days += e.month === 1 ? 31 : daysInMonth(e.year, e.month - 1); // the month before the end date
        return days;
      }
      case 'YM': {
        let months = (e.year - s.year) * 12 + (e.month - s.month);
        if (e.day < s.day) months--;
        return ((months % 12) + 12) % 12;
      }
      case 'YD': {
        let anchor = ymdToSerial(s.year, e.month, e.day);
        if (anchor < start) anchor = ymdToSerial(s.year + 1, e.month, e.day);
        return anchor - start;
      }
      default:
        throw new Error('#NUM!');
    }
  },

  NETWORKDAYS: (args) => {
    let start = Math.floor(toSerial(args[0]));
    let end = Math.floor(toSerial(args[1]));
    let sign = 1;
    if (start > end) {
      [start, end] = [end, start];
      sign = -1;
    }
    const holidays = new Set(
      args[2] !== undefined ? flatten([args[2]]).map((h) => Math.floor(toSerial(h))) : []
    );
    let count = 0;
    for (let s = start; s <= end; s++) {
      const wd = dateParts(s).weekday;
      if (wd !== 0 && wd !== 6 && !holidays.has(s)) count++;
    }
    return count * sign;
  },

  WORKDAY: (args) => {
    let serial = Math.floor(toSerial(args[0]));
    const days = Math.trunc(toNum(args[1]));
    const holidays = new Set(
      args[2] !== undefined ? flatten([args[2]]).map((h) => Math.floor(toSerial(h))) : []
    );
    const step = days >= 0 ? 1 : -1;
    let remaining = Math.abs(days);
    while (remaining > 0) {
      serial += step;
      const wd = dateParts(serial).weekday;
      if (wd !== 0 && wd !== 6 && !holidays.has(serial)) remaining--;
    }
    return serial;
  },

  DATEVALUE: (args) => {
    const serial = parseDateString(toText(args[0]));
    if (serial === null) throw new Error('#VALUE!');
    return Math.floor(serial);
  },

  TIMEVALUE: (args) => {
    const fraction = parseTimeString(toText(args[0]));
    if (fraction === null) throw new Error('#VALUE!');
    return fraction;
  },
};
