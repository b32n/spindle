// Excel serial dates. The 1900 system counts from 1900-01-01 = 1 and wrongly has a 29 February 1900 (serial 60),
// so serials below 60 are one day further from the real calendar than serials from 61; serial 0 is "1900-01-00".
// The 1904 system (older Mac workbooks) counts from 1904-01-01 = 0 and has no such day.

export type DateSystem = 1900 | 1904;

export interface CalendarDate {
  year: number;
  month: number; // 1..12
  day: number; // 0 only for serial 0 of the 1900 system
  weekday: number; // 0 = Sunday
}

const DAY_MS = 86_400_000;
const MAX_SERIAL: Record<DateSystem, number> = { 1900: 2_958_465, 1904: 2_957_003 }; // 9999-12-31

function utc(year: number, month: number, day: number): number {
  const d = new Date(0);
  d.setUTCFullYear(year, month - 1, day);
  return d.getTime();
}

/** Calendar date of the whole-day part of a serial, or null when Excel shows #### (negative or after 9999-12-31). */
export function serialToDate(serial: number, system: DateSystem = 1900): CalendarDate | null {
  const day = Math.floor(serial);
  if (!(day >= 0) || day > MAX_SERIAL[system]) return null;
  if (system === 1904) {
    const d = new Date(utc(1904, 1, 1) + day * DAY_MS);
    return { year: d.getUTCFullYear(), month: d.getUTCMonth() + 1, day: d.getUTCDate(), weekday: d.getUTCDay() };
  }
  const weekday = (((day - 1) % 7) + 7) % 7; // Excel's calendar says serial 1 was a Sunday
  if (day === 0) return { year: 1900, month: 1, day: 0, weekday };
  if (day === 60) return { year: 1900, month: 2, day: 29, weekday };
  const d = new Date(utc(1899, 12, day < 60 ? 31 : 30) + day * DAY_MS);
  return { year: d.getUTCFullYear(), month: d.getUTCMonth() + 1, day: d.getUTCDate(), weekday };
}

/** Serial of a calendar date; the inverse of serialToDate (1900-02-29 is serial 60). Null when out of range. */
export function dateToSerial(year: number, month: number, day: number, system: DateSystem = 1900): number | null {
  const time = utc(year, month, day);
  let serial: number;
  if (system === 1904) {
    serial = (time - utc(1904, 1, 1)) / DAY_MS;
  } else if (year === 1900 && month === 2 && day === 29) {
    serial = 60;
  } else {
    serial = (time - utc(1899, 12, time < utc(1900, 3, 1) ? 31 : 30)) / DAY_MS;
  }
  return Number.isInteger(serial) && serial >= 0 && serial <= MAX_SERIAL[system] ? serial : null;
}
