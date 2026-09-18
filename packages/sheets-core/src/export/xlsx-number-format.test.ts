import { fromExcelNumberFormatCode, toExcelNumberFormatCode } from './xlsx-number-format';

describe('toExcelNumberFormatCode', () => {
  it('formats a plain number with and without a thousands separator', () => {
    expect(toExcelNumberFormatCode({ type: 'number', decimalPlaces: 2 })).toBe('0.00');
    expect(toExcelNumberFormatCode({ type: 'number', decimalPlaces: 2, useThousandsSeparator: true })).toBe('#,##0.00');
    expect(toExcelNumberFormatCode({ type: 'number', decimalPlaces: 0 })).toBe('0');
  });

  it('applies a negative-number display style', () => {
    expect(toExcelNumberFormatCode({ type: 'number', decimalPlaces: 0, negativeFormat: 'parentheses' })).toBe('0;(0)');
    expect(toExcelNumberFormatCode({ type: 'number', decimalPlaces: 0, negativeFormat: 'red' })).toBe('0;[Red]-0');
    expect(toExcelNumberFormatCode({ type: 'number', decimalPlaces: 0, negativeFormat: 'minus' })).toBe('0');
  });

  it('places the currency symbol as a prefix or suffix', () => {
    expect(toExcelNumberFormatCode({ type: 'currency', currencyCode: 'USD', decimalPlaces: 2 })).toBe('"$"0.00');
    expect(
      toExcelNumberFormatCode({ type: 'currency', currencyCode: 'EUR', decimalPlaces: 2, currencySymbolPosition: 'suffix' })
    ).toBe('0.00"€"');
  });

  it('falls back to $ for an unknown currency code', () => {
    expect(toExcelNumberFormatCode({ type: 'currency', currencyCode: 'XYZ', decimalPlaces: 0 })).toBe('"$"0');
  });

  it('formats a percentage', () => {
    expect(toExcelNumberFormatCode({ type: 'percentage', decimalPlaces: 1 })).toBe('0.0%');
  });

  it('formats scientific notation', () => {
    expect(toExcelNumberFormatCode({ type: 'scientific', decimalPlaces: 2 })).toBe('0.00E+00');
  });

  it('maps each fraction type to an Excel denominator pattern', () => {
    expect(toExcelNumberFormatCode({ type: 'fraction', fractionType: 'asQuarters' })).toBe('# ?/4');
    expect(toExcelNumberFormatCode({ type: 'fraction' })).toBe('# ?/?'); // default: upToOne
  });

  it('maps each known date pattern, and falls back for an unrecognized one', () => {
    expect(toExcelNumberFormatCode({ type: 'date', dateFormat: 'MM/DD/YYYY' })).toBe('mm/dd/yyyy');
    expect(toExcelNumberFormatCode({ type: 'date', dateFormat: 'YYYY-MM-DD' })).toBe('yyyy-mm-dd');
    expect(toExcelNumberFormatCode({ type: 'date', dateFormat: 'Month DD YYYY' })).toBe('mmmm dd yyyy');
    expect(toExcelNumberFormatCode({ type: 'date' })).toBe('m/d/yyyy');
  });

  it('maps each known time pattern, and falls back for an unrecognized one', () => {
    expect(toExcelNumberFormatCode({ type: 'time', timeFormat: 'HH:mm:ss' })).toBe('hh:mm:ss');
    expect(toExcelNumberFormatCode({ type: 'time', timeFormat: 'h:mm AM/PM' })).toBe('h:mm AM/PM');
    expect(toExcelNumberFormatCode({ type: 'time' })).toBe('h:mm:ss AM/PM');
  });

  it('combines date and time codes for datetime', () => {
    expect(toExcelNumberFormatCode({ type: 'datetime', dateFormat: 'YYYY-MM-DD', timeFormat: 'HH:mm' })).toBe(
      'yyyy-mm-dd hh:mm'
    );
  });

  it('maps each duration unit', () => {
    expect(toExcelNumberFormatCode({ type: 'duration', durationFormat: 'hours' })).toBe('[hh]:mm:ss');
    expect(toExcelNumberFormatCode({ type: 'duration', durationFormat: 'minutes' })).toBe('[mm]:ss');
    expect(toExcelNumberFormatCode({ type: 'duration', durationFormat: 'seconds' })).toBe('[ss]');
    expect(toExcelNumberFormatCode({ type: 'duration', durationFormat: 'milliseconds' })).toBe('[ss].000');
  });

  it('passes a custom pattern through as-is, falling back to General if empty', () => {
    expect(toExcelNumberFormatCode({ type: 'custom', pattern: '0.0"x"' })).toBe('0.0"x"');
    expect(toExcelNumberFormatCode({ type: 'custom' })).toBe('General');
  });

  it('uses General for text and unset types', () => {
    expect(toExcelNumberFormatCode({ type: 'text' })).toBe('General');
    expect(toExcelNumberFormatCode({})).toBe('General');
  });

  it('produces a valid accounting format with the currency symbol threaded through', () => {
    const code = toExcelNumberFormatCode({ type: 'accounting', currencyCode: 'GBP', decimalPlaces: 2 });
    expect(code).toContain('£');
    expect(code).toContain('#,##0.00');
  });
});

describe('fromExcelNumberFormatCode', () => {
  it('maps General (and empty) to no format', () => {
    expect(fromExcelNumberFormatCode('General')).toEqual({});
    expect(fromExcelNumberFormatCode('')).toEqual({});
  });

  it('exactly reverses every date/time/datetime code our own writer produces', () => {
    for (const dateFormat of ['MM/DD/YYYY', 'DD-MM-YYYY', 'YYYY-MM-DD', 'Month DD YYYY'] as const) {
      const code = toExcelNumberFormatCode({ type: 'date', dateFormat });
      expect(fromExcelNumberFormatCode(code)).toEqual({ type: 'date', dateFormat });
    }
    for (const timeFormat of ['HH:mm:ss', 'h:mm AM/PM', 'HH:mm'] as const) {
      const code = toExcelNumberFormatCode({ type: 'time', timeFormat });
      expect(fromExcelNumberFormatCode(code)).toEqual({ type: 'time', timeFormat });
    }
    const dtCode = toExcelNumberFormatCode({ type: 'datetime', dateFormat: 'YYYY-MM-DD', timeFormat: 'HH:mm' });
    expect(fromExcelNumberFormatCode(dtCode)).toEqual({ type: 'datetime', dateFormat: 'YYYY-MM-DD', timeFormat: 'HH:mm' });
  });

  it('sniffs a date/time pattern from a real spreadsheet app we did not author', () => {
    expect(fromExcelNumberFormatCode('dd/mm/yy')).toMatchObject({ type: 'date' });
    expect(fromExcelNumberFormatCode('h:mm:ss AM/PM')).toMatchObject({ type: 'time' });
  });

  it('recovers a percentage with its decimal precision', () => {
    expect(fromExcelNumberFormatCode('0.00%')).toEqual({ type: 'percentage', decimalPlaces: 2 });
  });

  it('recovers scientific notation', () => {
    expect(fromExcelNumberFormatCode('0.00E+00')).toEqual({ type: 'scientific', decimalPlaces: 2 });
  });

  it('recovers a currency format with symbol, position, and thousands separator', () => {
    expect(fromExcelNumberFormatCode('"$"#,##0.00')).toEqual({
      type: 'currency',
      currencyCode: 'USD',
      currencySymbolPosition: 'prefix',
      decimalPlaces: 2,
      useThousandsSeparator: true,
      negativeFormat: 'minus',
    });
    expect(fromExcelNumberFormatCode('0.00"€"')).toMatchObject({
      type: 'currency',
      currencyCode: 'EUR',
      currencySymbolPosition: 'suffix',
    });
  });

  it('recovers negative-format style from a currency or number code', () => {
    expect(fromExcelNumberFormatCode('"$"0.00;("$"0.00)')).toMatchObject({ negativeFormat: 'parentheses' });
    expect(fromExcelNumberFormatCode('0;[Red]-0')).toMatchObject({ type: 'number', negativeFormat: 'red' });
  });

  it('recovers an accounting format', () => {
    const code = toExcelNumberFormatCode({ type: 'accounting', currencyCode: 'GBP', decimalPlaces: 2 });
    expect(fromExcelNumberFormatCode(code)).toMatchObject({ type: 'accounting', currencyCode: 'GBP', decimalPlaces: 2 });
  });

  it('recovers a plain number format with thousands separator', () => {
    expect(fromExcelNumberFormatCode('#,##0.00')).toEqual({
      type: 'number',
      decimalPlaces: 2,
      useThousandsSeparator: true,
      negativeFormat: 'minus',
    });
  });

  it('falls back to custom for a pattern it cannot classify, preserving the raw code', () => {
    expect(fromExcelNumberFormatCode('@ "units"')).toEqual({ type: 'custom', pattern: '@ "units"' });
  });
});
