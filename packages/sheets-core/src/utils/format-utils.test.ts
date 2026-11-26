import { parseDateString, excelDateToJS, formatNumber, formatCustomPattern, formatDate, formatTime, formatDateTime, formatScientific, formatFraction, formatDuration, formatPlainNumber, formatCurrency, formatAccounting, formatPercentage, dateStringToExcelSerial, dateTimeStringToExcelSerial, jsToExcelDate, formatJSDate, areFormatsEqual, getDefaultFormatForType } from './format-utils';

function ymd(serial: number | null): [number, number, number] | null {
  if (serial === null) return null;
  const d = excelDateToJS(serial);
  return [d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate()];
}

describe('parseDateString', () => {
  it('parses a slash date as MM/DD/YYYY (US convention)', () => {
    // The 12th of January 2020, not the 1st of December.
    expect(ymd(parseDateString('1/12/2020'))).toEqual([2020, 1, 12]);
  });

  it('parses a dash date as DD-MM-YYYY (European convention)', () => {
    expect(ymd(parseDateString('1-12-2020'))).toEqual([2020, 12, 1]);
  });

  it('parses a dot date as DD.MM.YYYY (European convention)', () => {
    expect(ymd(parseDateString('1.12.2020'))).toEqual([2020, 12, 1]);
  });

  it('parses YYYY-MM-DD ISO form', () => {
    expect(ymd(parseDateString('2020-01-15'))).toEqual([2020, 1, 15]);
  });

  it('parses YYYY/MM/DD with slashes', () => {
    expect(ymd(parseDateString('2020/01/15'))).toEqual([2020, 1, 15]);
  });

  it('accepts single-digit month and day', () => {
    expect(ymd(parseDateString('3/4/2021'))).toEqual([2021, 3, 4]);
  });

  it('rejects an invalid month', () => {
    expect(parseDateString('13/15/2020')).toBeNull();
  });

  it('rejects an invalid day (Feb 30)', () => {
    expect(parseDateString('2-30-2021')).toBeNull();
  });

  it('rejects a year outside the supported range', () => {
    expect(parseDateString('1/1/1800')).toBeNull();
  });

  it('rejects a non-date string', () => {
    expect(parseDateString('hello')).toBeNull();
  });

  it('rejects a date the calendar rolls over (Feb 30)', () => {
    expect(parseDateString('2/30/2021')).toBeNull();
  });

  it('rejects an incomplete date', () => {
    expect(parseDateString('1/15')).toBeNull();
  });
});

describe('formatNumber routes through the number-format engine', () => {
  it('shows an unformatted number in General, and a text-typed one as typed', () => {
    expect(formatNumber(0.1 + 0.2, {})).toBe('0.3');
    expect(formatNumber(1 / 3, {})).toBe('0.333333333');
    expect(formatNumber(123456789012, {})).toBe('1.23457E+11');
    expect(formatNumber(0.1 + 0.2, { type: 'text' })).toBe('0.30000000000000004');
  });

  it('renders custom patterns with the full code grammar', () => {
    expect(formatNumber(1234.5, { type: 'custom', pattern: '#,##0.00;[Red](#,##0.00)' })).toBe('1,234.50');
    expect(formatNumber(-1234.5, { type: 'custom', pattern: '#,##0.00;[Red](#,##0.00)' })).toBe('(1,234.50)');
    expect(formatNumber(46024.75, { type: 'custom', pattern: 'dddd, mmmm d, yyyy' })).toBe('Friday, January 2, 2026');
    expect(formatNumber(1500000, { type: 'custom', pattern: '0.0,,"M"' })).toBe('1.5M');
    expect(formatNumber(5, { type: 'custom' })).toBe('5');
    expect(formatCustomPattern(0.256, '0.0%')).toBe('25.6%');
    expect(formatCustomPattern(5, '0 x')).toBe('5');
  });

  it('renders the date, time and datetime types as Excel codes, with the old defaults', () => {
    expect(formatNumber(46024.75, { type: 'date', dateFormat: 'YYYY-MM-DD' })).toBe('2026-01-02');
    expect(formatNumber(46024.75, { type: 'date', dateFormat: 'Month DD YYYY' })).toBe('January 02 2026');
    expect(formatNumber(46024.75, { type: 'date' })).toBe('01/02/2026');
    expect(formatNumber(46024.75, { type: 'time', timeFormat: 'h:mm AM/PM' })).toBe('6:00 PM');
    expect(formatNumber(46024.75, { type: 'time' })).toBe('18:00:00');
    expect(formatNumber(46024.75, { type: 'datetime', dateFormat: 'DD-MM-YYYY', timeFormat: 'HH:mm' })).toBe('02-01-2026 18:00');
    expect(formatNumber(46024.75, { type: 'datetime' })).toBe('01/02/2026 18:00:00');
    expect(formatDate(45000, { dateFormat: 'MM/DD/YYYY' })).toBe('03/15/2023');
    expect(formatTime(0.5, { timeFormat: 'HH:mm' })).toBe('12:00');
    expect(formatDateTime(45000.5, {})).toBe('03/15/2023 12:00:00');
  });

  it('renders scientific, fraction and duration types', () => {
    expect(formatNumber(1234.5, { type: 'scientific', decimalPlaces: 2 })).toBe('1.23E+03');
    expect(formatNumber(1234.5, { type: 'scientific' })).toBe('1.23E+03');
    expect(formatScientific(0.000123, { decimalPlaces: 1 })).toBe('1.2E-04');
    expect(formatNumber(1234.5, { type: 'scientific', decimalPlaces: 500 })).toContain('E+03');
    expect(formatNumber(1.25, { type: 'fraction', fractionType: 'upToOne' })).toBe('1 1/4');
    expect(formatNumber(1.3, { type: 'fraction', fractionType: 'asQuarters' })).toBe('1 1/4');
    expect(formatFraction(0.5, { fractionType: 'asSixteenths' })).toBe(' 8/16');
    expect(formatNumber(1.5, { type: 'duration', durationFormat: 'hours' })).toBe('36:00:00');
    expect(formatNumber(1.5, { type: 'duration', durationFormat: 'minutes' })).toBe('2160:00');
    expect(formatNumber(1.5, { type: 'duration', durationFormat: 'seconds' })).toBe('129600');
    expect(formatDuration(0.0001, { durationFormat: 'milliseconds' })).toBe('08.640');
  });

  it('leaves the number, currency, accounting and percentage types as they were', () => {
    expect(formatNumber(1234.5, { type: 'number', decimalPlaces: 1 })).toBe('1,234.5');
    expect(formatNumber(1234.5, { type: 'currency', currencyCode: 'USD' })).toBe('$1,234.50');
    expect(formatNumber(0.256, { type: 'percentage', decimalPlaces: 1 })).toBe('25.6%');
  });
});

describe('structured formats', () => {
  it('formats plain numbers with fixed or automatic decimals, thousands and negative styles', () => {
    expect(formatPlainNumber(1234.5, { decimalPlaces: 2 })).toBe('1,234.50');
    expect(formatPlainNumber(1234.5, {})).toBe('1,234.5');
    expect(formatPlainNumber(1000, {})).toBe('1,000');
    expect(formatPlainNumber(1234.5, { decimalPlaces: 1, useThousandsSeparator: false })).toBe('1234.5');
    expect(formatPlainNumber(-5, { decimalPlaces: 0, negativeFormat: 'parentheses' })).toBe('(5)');
    expect(formatPlainNumber(-5, { decimalPlaces: 0, negativeFormat: 'red' })).toBe('-5');
    expect(formatPlainNumber(-5, { decimalPlaces: 0, negativeFormat: 'minus' })).toBe('-5');
    expect(formatPlainNumber(1.5, { decimalPlaces: Number.NaN })).toBe('1.50');
    expect(formatPlainNumber(1.5, { decimalPlaces: 1.9 })).toBe('1.5');
    expect(formatPlainNumber(1.5, { decimalPlaces: -4 })).toBe('2');
  });

  it('formats currency with Intl, its fallback, and the negative styles', () => {
    expect(formatCurrency(1234.5, { currencyCode: 'EUR', decimalPlaces: 1 })).toBe('€1,234.5');
    expect(formatCurrency(1234.5, {})).toBe('$1,234.50');
    expect(formatCurrency(-5, { negativeFormat: 'parentheses' })).toBe('($5.00)');
    expect(formatCurrency(-5, { negativeFormat: 'red' })).toBe('$5.00');
    expect(formatCurrency(-5, { negativeFormat: 'minus' })).toBe('-$5.00');
    expect(formatCurrency(5, { currencyCode: 'not-a-currency' })).toBe('not-a-currency5');
    expect(formatCurrency(5, { currencyCode: 'not-a-currency', currencySymbolPosition: 'suffix' })).toBe('5not-a-currency');
    expect(formatCurrency(5, { currencyCode: 'XYZ' })).toBe('XYZ\u00a05.00');
  });

  it('formats accounting and percentages', () => {
    expect(formatAccounting(5, { decimalPlaces: 2 })).toBe('$    5.00');
    expect(formatAccounting(-5, { currencyCode: 'CAD', negativeFormat: 'parentheses' })).toBe('(C$    5)');
    expect(formatAccounting(-5, { negativeFormat: 'red' })).toBe('$    5');
    expect(formatAccounting(-5, {})).toBe('$    (5)');
    expect(formatPercentage(0.256, { decimalPlaces: 1 })).toBe('25.6%');
    expect(formatPercentage(0.5, {})).toBe('50%');
    expect(formatPercentage(0.5, { decimalPlaces: 0 })).toBe('50%');
    expect(formatNumber(-0.256, { type: 'percentage' })).toBe('-25.6%');
    expect(formatNumber(5, { type: 'accounting' })).toBe('$    5');
    expect(formatNumber(5, { type: 'bogus' as never })).toBe('5');
  });
});

describe('date helpers', () => {
  it('converts between serials and dates in UTC', () => {
    expect(dateStringToExcelSerial('2026-01-02')).toBe(46024);
    expect(dateTimeStringToExcelSerial('2026-01-02T18:00')).toBe(46024.75);
    expect(dateTimeStringToExcelSerial('2026-01-02')).toBe(46024);
    expect(jsToExcelDate(new Date(Date.UTC(2026, 0, 2, 18)))).toBe(46024);
    expect(excelDateToJS(46024.75).toISOString()).toBe('2026-01-02T00:00:00.000Z');
  });

  it('formats a JS date with the legacy patterns', () => {
    const d = new Date(Date.UTC(2026, 0, 2));
    expect(formatJSDate(d, 'MM/DD/YYYY')).toBe('01/02/2026');
    expect(formatJSDate(d, 'DD-MM-YYYY')).toBe('02-01-2026');
    expect(formatJSDate(d, 'YYYY-MM-DD')).toBe('2026-01-02');
    expect(formatJSDate(d, 'Month DD YYYY')).toBe('January 2 2026');
    expect(typeof formatJSDate(d, 'other')).toBe('string');
  });
});

describe('format helpers', () => {
  it('compares formats and builds defaults for every type', () => {
    expect(areFormatsEqual(undefined, undefined)).toBe(true);
    expect(areFormatsEqual(undefined, {})).toBe(false);
    expect(areFormatsEqual({}, undefined)).toBe(false);
    expect(areFormatsEqual({ type: 'date' }, { type: 'date' })).toBe(true);
    expect(areFormatsEqual({ type: 'date' }, { type: 'time' })).toBe(false);
    const types = ['number', 'currency', 'accounting', 'percentage', 'scientific', 'fraction', 'date', 'time', 'datetime', 'duration', 'custom', 'text', undefined] as const;
    for (const type of types) expect(getDefaultFormatForType(type as never).type).toBe(type ?? 'text');
    expect(getDefaultFormatForType('custom').pattern).toBe('#,##0.00');
  });
});
