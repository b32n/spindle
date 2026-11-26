import { compileFormat, dateToSerial, formatDisplay, formatGeneral, formatValue, isDateTimeFormatCode, isValidFormatCode, serialToDate } from './index';
import { generalDigits, roundToPlaces, toDecimal } from './digits';
import { selectSection } from './render';

const show = (value: unknown, code: string): string => formatDisplay(value, code);

describe('digit placeholders', () => {
  it('0 pads, # does not, ? pads with a space', () => {
    expect(show(5, '000')).toBe('005');
    expect(show(5, '###')).toBe('5');
    expect(show(5, '??0')).toBe('  5');
    expect(show(0, '#,###')).toBe('');
    expect(show(0.05, '.00')).toBe('.05');
    expect(show(12.5, '.00')).toBe('12.50');
  });
  it('shows digits beyond the placeholders and trims # decimals', () => {
    expect(show(123456, '00')).toBe('123456');
    expect(show(1.5, '0.0#')).toBe('1.5');
    expect(show(1.05, '0.0#')).toBe('1.05');
    expect(show(1.2, '0.??')).toBe('1.2 ');
    expect(show(5, '0.')).toBe('5.');
  });
  it('interleaves literals with the digits, right to left', () => {
    expect(show(1234, '000-00')).toBe('012-34');
    expect(show(5551234567, '(000) 000-0000')).toBe('(555) 123-4567');
  });
  it('groups thousands, also for padded and long numbers, and handles trailing commas as scaling', () => {
    expect(show(1234567.891, '#,##0.00')).toBe('1,234,567.89');
    expect(show(5, '0,000')).toBe('0,005');
    expect(show(1234.5, '0,0')).toBe('1,235');
    expect(show(1234567, '#,##0,')).toBe('1,235');
    expect(show(1500, '0,')).toBe('2');
    expect(show(1e6, '0.0,,"M"')).toBe('1.0M');
    expect(show(1234, '0.0,')).toBe('1.2');
    expect(show(5, ',0')).toBe(',5');
    expect(show(5, '0.0.0')).toBe('5.0.0');
  });
  it('rounds half away from zero on the 15 significant digits', () => {
    expect(show(1.005, '0.00')).toBe('1.01');
    expect(show(2.675, '0.00')).toBe('2.68');
    expect(show(0.5, '0')).toBe('1');
    expect(show(-0.5, '0')).toBe('-1');
    expect(show(2.5, '0')).toBe('3');
    expect(show(999.995, '0.00')).toBe('1000.00');
    expect(show(1e22, '0')).toBe('10000000000000000000000');
    expect(show(1.2345678901234568e17, '0')).toBe('123456789012346000');
  });
  it('keeps the minus sign of a negative that rounds to zero (Excel shows -0 and -0.00)', () => {
    expect(show(-0.001, '0.00')).toBe('-0.00');
    expect(show(-0.4, '0')).toBe('-0');
    expect(show(-0.6, '0')).toBe('-1');
    expect(show(-0.001, '$#,##0.00')).toBe('-$0.00');
    expect(show(-0.001, '0.00;(0.00)')).toBe('(0.00)');
  });
  it('percent multiplies by 100 per sign', () => {
    expect(show(0.256, '0.0%')).toBe('25.6%');
    expect(show(0.25, '0.00%%')).toBe('2500.00%%');
    expect(show(1.5, '0%')).toBe('150%');
  });
});

describe('scientific', () => {
  it('formats E+00, E-00, e+0 and engineering codes', () => {
    expect(show(12345.6789, '0.00E+00')).toBe('1.23E+04');
    expect(show(0.000123, '0.00E+00')).toBe('1.23E-04');
    expect(show(0, '0.00E+00')).toBe('0.00E+00');
    expect(show(1.5, '0.0E-0')).toBe('1.5E0');
    expect(show(0.015, '0.0E-0')).toBe('1.5E-2');
    expect(show(1234.5, '0.0e+0')).toBe('1.2e+3');
    expect(show(12345.6789, '##0.0E+0')).toBe('12.3E+3');
    expect(show(0.256, '##0.0E+0')).toBe('256.0E-3');
    expect(show(99999, '0.0E+00')).toBe('1.0E+05');
    expect(show(1e100, '0.0E+000')).toBe('1.0E+100');
    expect(show(-1234.5, '0.0E+00')).toBe('-1.2E+03');
    expect(show(1234, '.00E+00')).toBe('1.23E+03');
    expect(show(1234, 'E+00')).toBe('1E+03');
  });
});

describe('fractions', () => {
  it('uses the closest fraction with the smallest denominator', () => {
    expect(show(8.5, '# ?/?')).toBe('8 1/2');
    expect(show(1234.5678, '# ?/?')).toBe('1234 4/7');
    expect(show(0.333, '# ??/??')).toBe('  1/3 ');
    expect(show(0.5, '?/?')).toBe('1/2');
    expect(show(3, '?/?')).toBe('3/1');
    expect(show(-1.5, '# ?/2')).toBe('-1 1/2');
    expect(show(0.333, '# ???/???')).toBe(' 332/997');
  });
  it('supports fixed denominators and carries into the whole part', () => {
    expect(show(2.3, '# ?/4')).toBe('2 1/4');
    expect(show(0.5, '0/16')).toBe('8/16');
    expect(show(1.97, '# ?/4')).toBe('2   ');
    expect(show(0.123, '# ??/100')).toBe(' 12/100');
    expect(show(2.5, '# ?/10')).toBe('2 5/10');
    expect(show(0.5, '# 0/8')).toBe(' 4/8');
    expect(show(2, '# 0/8')).toBe('2 0/8');
  });
  it('blanks the fraction of a whole number and shows zero as 0', () => {
    expect(show(2, '# ?/?')).toBe('2    ');
    expect(show(2, '# ?/8')).toBe('2   ');
    expect(show(0, '# ?/?')).toBe('0    ');
    expect(show(7.9999, '# ?/?')).toBe('8    ');
    expect(show(0, '# #/#')).toBe('0  ');
    expect(show(-0.001, '# ?/?')).toBe('-0    ');
  });
});

describe('sections, conditions and colours', () => {
  it('picks positive; negative; zero; text', () => {
    expect(show(5, '0;(0);"zero"')).toBe('5');
    expect(show(-5, '0;(0);"zero"')).toBe('(5)');
    expect(show(0, '0;(0);"zero"')).toBe('zero');
    expect(show(0, '0;(0)')).toBe('0');
    expect(show(-5, '0;;')).toBe('');
    expect(show(5, ';;;')).toBe('');
    expect(show('text', '0;0;0;"t"@')).toBe('ttext');
    expect(show(-5, '0')).toBe('-5');
    expect(show(-5, '0;-0')).toBe('-5');
  });
  it('conditions pick the first matching section', () => {
    expect(show(150, '[>100]"big";"small"')).toBe('big');
    expect(show(50, '[>100]"big";"small"')).toBe('small');
    expect(show(-50, '[>100]"big";"small"')).toBe('small');
    expect(show(50, '[>100]"a";[<0]"b";"c"')).toBe('c');
    expect(show(-5, '[>100]"a";[<0]"b";"c"')).toBe('b');
    expect(show(-5, '[>100]"a";"b";"c"')).toBe('b');
    expect(show(0, '[>100]"a";"b";"c"')).toBe('c');
    expect(show(5, '[=5]"five";0')).toBe('five');
    expect(show(5, '[<>5]"no";"yes"')).toBe('yes');
    expect(show(5, '[>=5]"ge";"lt"')).toBe('ge');
    expect(show(5, '[<=5]"le";"gt"')).toBe('le');
    expect(show(5, '[<5]"lt";"ge"')).toBe('ge');
    expect(show(5, '0;[>3]0"x";0"y"')).toBe('5');
  });
  it('a negative section chosen by condition shows the magnitude, other sections keep the sign', () => {
    expect(show(-5, '[<0][Red]0.0;[Blue]0.0')).toBe('5.0');
    expect(show(-5, '[<=0]"n"0;0')).toBe('n5');
    expect(show(-5, '[<10]0.0;[>=100]0.00;0')).toBe('-5.0');
    expect(show(-5, '[>100]0"a";[<50]0"b";0"c"')).toBe('5b');
    expect(show(-5, '[<1]0;0')).toBe('-5');
    expect(show(-2, '[<>1]0;0')).toBe('-2');
  });
  it('a lone condition that is not met falls back to General', () => {
    expect(show(5, '[>100]0.00')).toBe('5');
    expect(show(-5, '[>100]0.00')).toBe('-5');
    expect(show(500, '[>100]0.00')).toBe('500.00');
  });
  it('reports the colour of the chosen section', () => {
    expect(formatValue(-5, '0;[Red]-0')).toEqual({ text: '-5', color: 'Red' });
    expect(formatValue(5, '[Blue]0.0;[Red]-0.0')).toEqual({ text: '5.0', color: 'Blue' });
    expect(formatValue(5, '[Color12]0')).toEqual({ text: '5', color: 'Color12' });
    expect(formatValue(5, '[CYAN]0').color).toBe('Cyan');
    expect(formatValue('x', '0;0;0;[Green]@').color).toBe('Green');
    expect(formatValue(5, '0').color).toBeUndefined();
  });
  it('keeps the locale of [$...] tokens on the section', () => {
    const [section] = compileFormat('[$€-407] #,##0.00').sections;
    expect(section.locale).toBe('407');
  });
});

describe('literals, locale and padding', () => {
  it('quotes, backslashes and the unquoted literal characters', () => {
    expect(show(5, '0 "kg"')).toBe('5 kg');
    expect(show(5, '"Total: "0')).toBe('Total: 5');
    expect(show(5, '\\$0.00')).toBe('$5.00');
    expect(show(5, '$#,##0')).toBe('$5');
    expect(show(5, '0-')).toBe('5-');
    expect(show(5, '(0)')).toBe('(5)');
    expect(show(5, '0 1')).toBe('5 1');
    expect(show(5, '+0 €')).toBe('+5 €');
    expect(show(5, 'General" units"')).toBe('5 units');
    expect(show(5, '"a;b"0')).toBe('a;b5');
    expect(show(5, '"literal only"')).toBe('literal only');
    expect(show(-5, '"literal only"')).toBe('-literal only');
    expect(show(5, 'GENERAL')).toBe('5');
    expect(show(5, '"x"General"y"')).toBe('x5y');
    expect(show(5, '0/')).toBe('5/');
    expect(show(5, ',')).toBe(',');
    expect(show(5, '%')).toBe('%');
    expect(show(5, '/')).toBe('/');
    expect(show(5, '1')).toBe('1');
    expect(show(5, 'General%,/1')).toBe('5%,/1');
    expect(show(5, '%0')).toBe('%5'.replace('5', '500'));
  });
  it('currency and locale tokens', () => {
    expect(show(1234.5, '[$€-407] #,##0.00')).toBe('€ 1,234.50');
    expect(show(1234.5, '[$USD] #,##0')).toBe('USD 1,235');
    expect(show(1234.5, '[$-409]#,##0')).toBe('1,235');
    expect(show(1234.5, '#,##0.00[$ kr-41D]')).toBe('1,234.50 kr');
    expect(show(-1234.5, '[$€-407] #,##0.00')).toBe('-€ 1,234.50');
  });
  it('the system date and time tokens stand for the en-US long forms', () => {
    expect(show(46024.75, '[$-F800]')).toBe('Friday, January 2, 2026');
    expect(show(46024.75, '[$-F400]')).toBe('6:00:00 PM');
    expect(show(46024.75, '[$-F800]dddd, mmmm dd, yyyy')).toBe('Friday, January 02, 2026');
    expect(show(0.75, '[$-F400]h:mm')).toBe('18:00');
  });
  it('_x is a space and *x a fill position that is not repeated', () => {
    expect(show(5, '_(0_)')).toBe(' 5 ');
    expect(formatValue(5, '*-0')).toEqual({ text: '5', fill: { index: 0, char: '-' } });
    expect(formatValue(5, '0*-')).toEqual({ text: '5', fill: { index: 1, char: '-' } });
    expect(formatValue(-5, '0*-;0*.').fill).toEqual({ index: 1, char: '.' });
    expect(formatValue(-5, '0*-').fill).toEqual({ index: 2, char: '-' });
    expect(formatValue(5, '_($* #,##0.00_)')).toEqual({ text: ' $5.00 ', fill: { index: 2, char: ' ' } });
    expect(formatValue(5, '0*-*+').fill).toEqual({ index: 1, char: '-' });
    expect(formatValue(5, '0').fill).toBeUndefined();
  });
});

describe('text and the @ section', () => {
  it('@ stands for the text of a text value', () => {
    expect(show('abc', '@')).toBe('abc');
    expect(show('abc', '@" (text)"')).toBe('abc (text)');
    expect(show('abc', '@@')).toBe('abcabc');
    expect(show('abc', '0;0;0;@"!"')).toBe('abc!');
    expect(show('abc', '0;"t"@')).toBe('tabc');
    expect(show('abc', '"a"@;"b"')).toBe('abc');
    expect(show('abc', '0.00')).toBe('abc');
    expect(show('abc', '0;0;0;"hidden"')).toBe('hidden');
    expect(show('abc', '0;0;0;')).toBe('');
    expect(show('abc', '0.0 "x" %/,.')).toBe('abc');
    expect(show('abc', '0;0;0;0.0"x"%')).toBe('x');
    expect(show('abc', '0;0;0;# ?/?"y"')).toBe(' y');
    expect(show('abc', '0;0;0;"-"_)*x@')).toBe('- abc');
    expect(formatValue('abc', '0;0;0;*-@').fill).toEqual({ index: 0, char: '-' });
    expect(formatValue('abc', '0;0;0;[Red]@').color).toBe('Red');
  });
  it('a number in a text-only section shows as General', () => {
    expect(show(5, '@')).toBe('5');
    expect(show(-1234.5678, '@" (text)"')).toBe('-1234.5678');
  });
  it('booleans, errors, empty values and non-finite numbers are not formatted', () => {
    expect(show(true, '0.00')).toBe('TRUE');
    expect(show(false, '0.00')).toBe('FALSE');
    expect(show('#N/A', '0.00')).toBe('#N/A');
    expect(show('#DIV/0!', '"x"@')).toBe('#DIV/0!');
    expect(show('#NAME? Function X not found', '@"!"')).toBe('#NAME? Function X not found');
    expect(show(null, '0.00')).toBe('');
    expect(show(undefined, '0.00')).toBe('');
    expect(show(NaN, '0')).toBe('#NUM!');
    expect(show(Infinity, '0')).toBe('#NUM!');
    expect(show({}, '0')).toBe('[object Object]');
  });
});

describe('General', () => {
  it('shows what fits into 11 characters', () => {
    expect(formatGeneral(1234.5678)).toBe('1234.5678');
    expect(formatGeneral(0.1 + 0.2)).toBe('0.3');
    expect(formatGeneral(1 / 3)).toBe('0.333333333');
    expect(formatGeneral(-1 / 3)).toBe('-0.33333333');
    expect(formatGeneral(Math.PI)).toBe('3.141592654');
    expect(formatGeneral(0)).toBe('0');
    expect(formatGeneral(100)).toBe('100');
    expect(formatGeneral(12345678901)).toBe('12345678901');
    expect(formatGeneral(99999999999.9)).toBe('1E+11');
    expect(formatGeneral(123456.7890123)).toBe('123456.789');
    expect(formatGeneral(0.0001)).toBe('0.0001');
    expect(formatGeneral(0.000123456789)).toBe('0.000123457');
  });
  it('switches to E notation from 1e11 and below 1e-4', () => {
    expect(formatGeneral(123456789012)).toBe('1.23457E+11');
    expect(formatGeneral(1e11)).toBe('1E+11');
    expect(formatGeneral(-12345678901)).toBe('-1.2346E+10');
    expect(formatGeneral(1e-5)).toBe('1E-05');
    expect(formatGeneral(1.5e-5)).toBe('1.5E-05');
    expect(formatGeneral(1e20)).toBe('1E+20');
    expect(formatGeneral(1.23456789e100)).toBe('1.2346E+100');
    expect(formatGeneral(9.9999999e-6)).toBe('1E-05');
    expect(formatGeneral(1e-300)).toBe('1E-300');
    expect(formatGeneral(NaN)).toBe('#NUM!');
  });
  it('is used for General codes and General sections', () => {
    expect(show(0.1 + 0.2, 'General')).toBe('0.3');
    expect(show(-5, 'General;General')).toBe('5');
    expect(show(5, '')).toBe('5');
  });
});

describe('dates and times', () => {
  it('formats the date tokens', () => {
    expect(show(46024.75, 'yyyy-mm-dd')).toBe('2026-01-02');
    expect(show(46024.75, 'yy')).toBe('26');
    expect(show(46024.75, 'm/d/yy')).toBe('1/2/26');
    expect(show(46024.75, 'dddd, mmmm d, yyyy')).toBe('Friday, January 2, 2026');
    expect(show(46024.75, 'ddd mmm dd')).toBe('Fri Jan 02');
    expect(show(46024.75, 'mmmmm')).toBe('J');
    expect(show(45000, 'mmmm')).toBe('March');
    expect(show(45000, 'MMM-YY')).toBe('Mar-23');
    expect(show(46024.75, 'd "of" mmmm')).toBe('2 of January');
    expect(show(46024.75, 'dd.mm.yyyy, ')).toBe('02.01.2026, ');
    expect(show(46024.75, 'yyyy%mm')).toBe('2026%01');
    expect(show(46024.75, 'yyyy mmmmmmm')).toBe('2026 January');
    expect(show(46024.75, 'mmmmm')).toBe('J');
    expect(show(46024.75, 'yyyyyy')).toBe('2026');
    expect(show(46024.75, 'dddddd')).toBe('Friday');
    expect(show(46024.75, '[$-x-sysdate]')).toBe('Friday, January 2, 2026');
    expect(show(0.75, '[$-x-systime]')).toBe('6:00:00 PM');
    expect(show(46024.75, 'yyyy 1')).toBe('2026 1');
  });
  it('tells minutes from months by the neighbouring tokens', () => {
    expect(show(46024.75, 'h:mm')).toBe('18:00');
    expect(show(46024.76, 'h:m')).toBe('18:14');
    expect(show(46024.76, 'mm:ss')).toBe('14:24');
    expect(show(46024.76, 'm:ss')).toBe('14:24');
    expect(show(46024.76, 'h "h" mm "m"')).toBe('18 h 14 m');
    expect(show(46024.76, 'mm')).toBe('01');
    expect(show(46024.76, 'm/d h:mm')).toBe('1/2 18:14');
    expect(show(46024.76, 'hh:mm:ss.0 mm')).toBe('18:14:24.0 01');
  });
  it('12-hour clock and the AM/PM spellings', () => {
    expect(show(0.5, 'h:mm AM/PM')).toBe('12:00 PM');
    expect(show(0, 'h:mm AM/PM')).toBe('12:00 AM');
    expect(show(0.25, 'hh:mm am/pm')).toBe('06:00 am');
    expect(show(0.75, 'h A/P')).toBe('6 P');
    expect(show(0.75, 'h a/p')).toBe('6 p');
    expect(show(0.75, 'h:mm')).toBe('18:00');
  });
  it('elapsed time', () => {
    expect(show(1.5, '[h]:mm')).toBe('36:00');
    expect(show(100.25, '[h]:mm:ss')).toBe('2406:00:00');
    expect(show(1.5, '[hh]:mm:ss')).toBe('36:00:00');
    expect(show(0.25, '[hh]:mm')).toBe('06:00');
    expect(show(0.5, '[mm]:ss')).toBe('720:00');
    expect(show(0.5, '[ss]')).toBe('43200');
    expect(show(0.5, '[m]')).toBe('720');
    expect(show(0.5, '[s]')).toBe('43200');
    expect(show(0.0001, '[ss].00')).toBe('08.64');
    expect(show(1.5, '[h] "hours"')).toBe('36 hours');
  });
  it('fractional seconds and rounding to the displayed second', () => {
    expect(show(0.0001, 'mm:ss.000')).toBe('00:08.640');
    expect(show(0.0001, 'mm:ss.0')).toBe('00:08.6');
    expect(show(0.123456789, 'h:mm:ss')).toBe('2:57:47');
    expect(show(0.99999999, 'h:mm:ss')).toBe('0:00:00');
    expect(show(0.999994213, 'h:mm')).toBe('23:59'); // no seconds shown: the time is cut, not rounded up
    expect(show(0.9999999, 'yyyy-mm-dd')).toBe('1900-01-00'); // a date-only code does not roll the date forward
    expect(show(0.99999999, 'h:mm:ss.000')).toBe('23:59:59.999');
    expect(show(1.999999999, 'yyyy-mm-dd h:mm:ss')).toBe('1900-01-02 0:00:00');
    expect(show(0.5, 'ss.0')).toBe('00.0');
    expect(show(0.5, 's.00')).toBe('0.00');
  });
  it('the 1900 date system has a 29 February 1900 and serial 0 is 1900-01-00', () => {
    expect(show(0, 'yyyy-mm-dd')).toBe('1900-01-00');
    expect(show(1, 'yyyy-mm-dd')).toBe('1900-01-01');
    expect(show(59, 'yyyy-mm-dd')).toBe('1900-02-28');
    expect(show(60, 'yyyy-mm-dd')).toBe('1900-02-29');
    expect(show(61, 'yyyy-mm-dd')).toBe('1900-03-01');
    expect(show(1, 'dddd')).toBe('Sunday');
    expect(show(0, 'dddd')).toBe('Saturday');
    expect(show(60, 'dddd')).toBe('Wednesday');
    expect(show(61, 'dddd')).toBe('Thursday');
    expect(show(2958465, 'yyyy-mm-dd')).toBe('9999-12-31');
  });
  it('the 1904 date system starts at 1904-01-01 and has no 29 February 1900', () => {
    expect(formatDisplay(0, 'yyyy-mm-dd', { dateSystem: 1904 })).toBe('1904-01-01');
    expect(formatDisplay(1461, 'yyyy-mm-dd', { dateSystem: 1904 })).toBe('1908-01-01');
    expect(formatDisplay(0, 'dddd', { dateSystem: 1904 })).toBe('Friday');
    expect(formatDisplay(2957003, 'yyyy-mm-dd', { dateSystem: 1904 })).toBe('9999-12-31');
    expect(formatDisplay(2957004, 'yyyy-mm-dd', { dateSystem: 1904 })).toBe('####');
  });
  it('shows #### for negative or out-of-range dates and times, unless a negative section handles them', () => {
    expect(show(-1, 'yyyy-mm-dd')).toBe('####');
    expect(show(-0.5, 'h:mm')).toBe('####');
    expect(show(-0.5, '[h]:mm')).toBe('####');
    expect(show(2958466, 'yyyy-mm-dd')).toBe('####');
    expect(show(-1.5, '[h]:mm;-[h]:mm')).toBe('-36:00');
    expect(show(2958466, 'h:mm')).toBe('0:00');
  });
  it('knows which codes are dates', () => {
    expect(isDateTimeFormatCode('yyyy-mm-dd')).toBe(true);
    expect(isDateTimeFormatCode('[h]:mm')).toBe(true);
    expect(isDateTimeFormatCode('0.00')).toBe(false);
    expect(isDateTimeFormatCode('General')).toBe(false);
    expect(isDateTimeFormatCode('"d"0')).toBe(false);
    expect(isDateTimeFormatCode('yyyy"')).toBe(false);
  });
});

describe('serial <-> calendar date', () => {
  it('maps the edge serials', () => {
    expect(serialToDate(0)).toEqual({ year: 1900, month: 1, day: 0, weekday: 6 });
    expect(serialToDate(60)).toEqual({ year: 1900, month: 2, day: 29, weekday: 3 });
    expect(serialToDate(61)).toEqual({ year: 1900, month: 3, day: 1, weekday: 4 });
    expect(serialToDate(-1)).toBeNull();
    expect(serialToDate(2958466)).toBeNull();
    expect(serialToDate(NaN)).toBeNull();
    expect(serialToDate(0, 1904)).toEqual({ year: 1904, month: 1, day: 1, weekday: 5 });
    expect(dateToSerial(1900, 2, 29)).toBe(60);
    expect(dateToSerial(1900, 3, 1)).toBe(61);
    expect(dateToSerial(1900, 1, 1)).toBe(1);
    expect(dateToSerial(1899, 12, 31)).toBe(0);
    expect(dateToSerial(1899, 12, 30)).toBeNull();
    expect(dateToSerial(10000, 1, 1)).toBeNull();
    expect(dateToSerial(1904, 1, 1, 1904)).toBe(0);
    expect(dateToSerial(1900, 1, 1, 1904)).toBeNull();
    expect(dateToSerial(NaN, 1, 1)).toBeNull();
  });
});

describe('invalid codes', () => {
  const invalid = [
    '"unterminated', '[Red', '0;0;0;0;0', 'abc', '0 x', '[xyz]0', '[Red][Blue]0', '[>1][>2]0', 'yyyy 0', '[DBNum1]0', '0E', '\\', '0_', '0*', '[color57]0',
    '[color0]0', '[Color 1]0', '0.0E+', 'h:mm 0', 'e', 'Generalx', 'General0', 'mm/dd/yy hh:mm:ss.0 E+0', '#,##0.00 [Red', 'ammm', '0 `', '0\t', '@ 0 a',
    'yyyy General', 'hh E+0', '0.0E+', '0E+x',
  ];
  it.each(invalid)('%j is invalid: numbers show as General, the code is reported', (code) => {
    expect(isValidFormatCode(code)).toBe(false);
    expect(compileFormat(code).error).toBeTruthy();
    expect(formatValue(1234.5, code)).toEqual({ text: '1234.5', invalid: true });
    expect(formatValue('x', code)).toEqual({ text: 'x', invalid: true });
  });
  it('accepts the codes Excel accepts', () => {
    for (const code of ['', 'General', '0', '@', ';;;', '0;;', '[Red]0', '"a;b"0', '0.0,,"M"', 'dd/mm/yyyy', '[$-409]mmm', '#,##0;[Red](#,##0)', '\\0', 'a/p', '[h]:mm:ss.000'])
      expect(isValidFormatCode(code)).toBe(true);
  });
});

describe('compiled-format cache', () => {
  it('returns the same compiled object for a repeated code and recovers after many codes', () => {
    const a = compileFormat('0.000');
    expect(compileFormat('0.000')).toBe(a);
    for (let i = 0; i < 600; i++) compileFormat(`0.${'0'.repeat(i % 7)}"${i}"`);
    expect(compileFormat('0.000').ok).toBe(true);
  });
});

describe('selectSection', () => {
  const sections = (code: string) => compileFormat(code).sections;
  it('chooses sections and the sign', () => {
    expect(selectSection(sections('0'), -1)).toEqual({ index: 0, minus: true });
    expect(selectSection(sections('0;0'), -1)).toEqual({ index: 1, minus: false });
    expect(selectSection(sections('0;0;0'), 0)).toEqual({ index: 2, minus: false });
    expect(selectSection(sections('0;0;0'), 3)).toEqual({ index: 0, minus: false });
    expect(selectSection(sections('[>1]0'), 0)).toEqual({ index: -1, minus: false });
    expect(selectSection(sections('[>1]0'), -2)).toEqual({ index: -1, minus: true });
  });
});

describe('digit helpers', () => {
  it('toDecimal keeps 15 significant digits', () => {
    expect(toDecimal(0)).toEqual({ digits: '0', exp: 0 });
    expect(toDecimal(0.1 + 0.2)).toEqual({ digits: '3', exp: -1 });
    expect(toDecimal(1234.5)).toEqual({ digits: '12345', exp: 3 });
    expect(toDecimal(5e-324)).toEqual({ digits: '494065645841247', exp: -324 });
  });
  it('roundToPlaces rounds half up on the digits and carries', () => {
    expect(roundToPlaces(toDecimal(9.995), 2)).toEqual({ int: '10', frac: '00' });
    expect(roundToPlaces(toDecimal(0.0004), 3)).toEqual({ int: '', frac: '000' });
    expect(roundToPlaces(toDecimal(0.0005), 3)).toEqual({ int: '', frac: '001' });
    expect(roundToPlaces(toDecimal(12), 2)).toEqual({ int: '12', frac: '00' });
    expect(roundToPlaces(toDecimal(1e20), 0)).toEqual({ int: '100000000000000000000', frac: '' });
  });
  it('generalDigits', () => {
    expect(generalDigits(0, false)).toBe('0');
    expect(generalDigits(5.5, true)).toBe('5.5');
  });
});
