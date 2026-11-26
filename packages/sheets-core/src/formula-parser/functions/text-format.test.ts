import { FormulaParser } from '../parser';
import type { EvaluationContext } from '../types';

const parser = new FormulaParser();
const ctx: EvaluationContext = { getCellValue: () => null, getRangeValues: () => [] };
const ev = (f: string): unknown => {
  const { ast, error } = parser.parse(f);
  if (error) throw new Error(error);
  return parser.evaluate(ast, ctx);
};

describe('TEXT respects positive;negative;zero format sections', () => {
  it('formats a negative value with the negative section (no auto-minus)', () => {
    expect(ev('=TEXT(-5,"0;0")')).toBe('5');
    expect(ev('=TEXT(-1234.5,"#,##0.00;#,##0.00")')).toBe('1,234.50');
  });

  it('formats a positive value with the first section', () => {
    expect(ev('=TEXT(5,"0;0")')).toBe('5');
    expect(ev('=TEXT(1234.5,"#,##0.00")')).toBe('1,234.50');
  });

  it('a single-section format keeps the sign', () => {
    expect(ev('=TEXT(-5,"0")')).toBe('-5');
    expect(ev('=TEXT(-5,"0.0")')).toBe('-5.0');
  });

  it('keeps the minus sign of a negative that rounds to zero, as Excel does ("-0")', () => {
    expect(ev('=TEXT(-0.4,"0")')).toBe('-0');
    expect(ev('=TEXT(-0.004,"0.00")')).toBe('-0.00');
    expect(ev('=TEXT(-0.6,"0")')).toBe('-1');
  });
});

describe('TEXT date and time codes', () => {
  it('formats dates with Excel codes', () => {
    expect(ev('=TEXT(DATE(2026,1,2),"yyyy-mm-dd")')).toBe('2026-01-02');
    expect(ev('=TEXT(DATE(2026,1,2),"dddd")')).toBe('Friday');
    expect(ev('=TEXT(DATE(2026,1,2),"mmm d, yyyy")')).toBe('Jan 2, 2026');
    expect(ev('=TEXT(DATE(2026,1,2),"dddd d mmmm")')).toBe('Friday 2 January');
    expect(ev('=TEXT(46024.75,"h:mm AM/PM")')).toBe('6:00 PM');
    expect(ev('=TEXT(46024.75,"m/d/yyyy hh:mm")')).toBe('1/2/2026 18:00');
    expect(ev('=TEXT(1.5,"[h]:mm")')).toBe('36:00');
    expect(ev('=TEXT(60,"yyyy-mm-dd")')).toBe('1900-02-29');
  });

  it('formats numbers with sections, conditions, scaling and fractions', () => {
    expect(ev('=TEXT(1234567,"#,##0.0,,""M""")')).toBe('1.2M');
    expect(ev('=TEXT(150,"[>100]""big"";""small""")')).toBe('big');
    expect(ev('=TEXT(0.75,"# ?/?")')).toBe(' 3/4');
    expect(ev('=TEXT(5,"000-00")')).toBe('000-05');
    expect(ev('=TEXT(1234.5,"$#,##0.00")')).toBe('$1,234.50');
    expect(ev('=TEXT(0,"0;-0;""zero""")')).toBe('zero');
    expect(ev('=TEXT(1234.5678,"General")')).toBe('1234.5678');
  });

  it('an empty code gives an empty string and an invalid code #VALUE!', () => {
    expect(ev('=TEXT(5,"")')).toBe('');
    expect(() => ev('=TEXT(5,"yyyy 0")')).toThrow('#VALUE!');
    expect(() => ev('=TEXT(5,"0 abc")')).toThrow('#VALUE!');
    expect(() => ev('=TEXT(5,"""open")')).toThrow('#VALUE!');
  });

  it('text, numeric text, booleans and blanks', () => {
    expect(ev('=TEXT("12.5","0.00")')).toBe('12.50');
    expect(ev('=TEXT("abc","0.00")')).toBe('abc');
    expect(ev('=TEXT("abc","@@")')).toBe('abcabc');
    expect(ev('=TEXT("","0.00")')).toBe('');
    expect(ev('=TEXT(TRUE,"0")')).toBe('TRUE');
    expect(ev('=TEXT(A1,"0.00")')).toBe('0.00');
  });
});

describe('TEXT reads text and General like Excel', () => {
  it('General has no column to fit: 15 significant digits', () => {
    expect(ev('=TEXT(355/113,"General")')).toBe('3.14159292035398');
    expect(ev('=TEXT(1/3,"General")')).toBe('0.333333333333333');
    expect(ev('=TEXT(1234.5678,"General")')).toBe('1234.5678');
    expect(ev('=TEXT(10^21,"General")')).toBe('1E+21');
    expect(ev('=TEXT(1/10^7,"General")')).toBe('1E-07');
    expect(ev('=TEXT(-1/3,"General")')).toBe('-0.333333333333333');
    expect(ev('=TEXT(355/113,"@")')).toBe('3.14159292035398');
    expect(ev('=TEXT(355/113,"[>10]0")')).toBe('3.14159292035398');
  });

  it('text that reads as a number or a date is formatted as one; other text passes through', () => {
    expect(ev('=TEXT("$5","0.00")')).toBe('5.00');
    expect(ev('=TEXT("1,234","#,##0.0")')).toBe('1,234.0');
    expect(ev('=TEXT("50%","0.0%")')).toBe('50.0%');
    expect(ev('=TEXT(" 12 ","0.0")')).toBe('12.0');
    expect(ev('=TEXT("2024-01-15","yyyy")')).toBe('2024');
    expect(ev('=TEXT("1/2/2020","d mmm")')).toBe('2 Jan');
    expect(ev('=TEXT("0x1F","0")')).toBe('0x1F');
    expect(ev('=TEXT("0b11","0")')).toBe('0b11');
    expect(ev('=TEXT("Infinity","0")')).toBe('Infinity');
    expect(ev('=TEXT("1,2,3","0")')).toBe('1,2,3');
    expect(ev('=TEXT("5%x","0")')).toBe('5%x');
    expect(ev('=TEXT("abc%","0")')).toBe('abc%');
    expect(ev('=TEXT("%","0")')).toBe('%');
    expect(ev('=TEXT("(5)","0")')).toBe('-5');
    expect(ev('=TEXT("($1,234.5)","0.0")')).toBe('-1234.5');
    expect(ev('=TEXT("(abc)","0")')).toBe('(abc)');
    expect(ev('=TEXT("((5))","0")')).toBe('((5))');
    expect(ev('=TEXT("(-5)","0")')).toBe('(-5)');
    expect(ev('=TEXT("(+5)","0")')).toBe('(+5)');
    expect(ev('=TEXT("(5)","@")')).toBe('(5)');
    expect(ev('=TEXT("1,234","@")')).toBe('1,234');
    expect(ev('=TEXT("1,234","@@")')).toBe('1,2341,234');
    expect(ev('=TEXT("12","""x""")')).toBe('12');
    expect(ev('=TEXT("1,234","0;0;0;@")')).toBe('1234');
  });
});
