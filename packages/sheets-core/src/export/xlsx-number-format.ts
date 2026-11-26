import type { CellFormat } from '../types';

/** Translate our CellFormat into an Excel-syntax number format code (e.g. "#,##0.00", "0.00%", "m/d/yyyy"). */
export function toExcelNumberFormatCode(format: CellFormat): string {
  switch (format.type) {
    case 'number':
      return applyNegativeFormat(numberCore(format), format.negativeFormat);
    case 'currency':
      return applyNegativeFormat(withCurrencySymbol(numberCore(format), format), format.negativeFormat);
    case 'accounting':
      return accountingCode(format);
    case 'percentage':
      return `${numberCore(format)}%`;
    case 'scientific': {
      const decimals = format.decimalPlaces ?? 2;
      return `0${decimals > 0 ? '.' + '0'.repeat(decimals) : ''}E+00`;
    }
    case 'fraction':
      return fractionCode(format.fractionType);
    case 'date':
      return dateCode(format.dateFormat);
    case 'time':
      return timeCode(format.timeFormat);
    case 'datetime':
      return `${dateCode(format.dateFormat)} ${timeCode(format.timeFormat)}`;
    case 'duration':
      return durationCode(format.durationFormat);
    case 'custom':
      return format.pattern || 'General';
    case 'text':
    default:
      return 'General';
  }
}

function numberCore(format: CellFormat): string {
  const decimals = format.decimalPlaces ?? 2;
  const intPart = format.useThousandsSeparator ? '#,##0' : '0';
  return decimals > 0 ? `${intPart}.${'0'.repeat(decimals)}` : intPart;
}

function applyNegativeFormat(positive: string, negativeFormat: CellFormat['negativeFormat']): string {
  switch (negativeFormat) {
    case 'parentheses':
      return `${positive};(${positive})`;
    case 'red':
      return `${positive};[Red]-${positive}`;
    case 'minus':
    default:
      return positive;
  }
}

const CURRENCY_SYMBOLS: Record<string, string> = {
  USD: '$', EUR: '€', GBP: '£', JPY: '¥', CAD: 'C$', AUD: 'A$',
  CHF: 'CHF', CNY: '¥', INR: '₹', KRW: '₩',
};

function currencySymbol(currencyCode: string | undefined): string {
  return CURRENCY_SYMBOLS[currencyCode ?? 'USD'] ?? '$';
}

function withCurrencySymbol(core: string, format: CellFormat): string {
  const symbol = currencySymbol(format.currencyCode);
  return format.currencySymbolPosition === 'suffix' ? `${core}"${symbol}"` : `"${symbol}"${core}`;
}

function accountingCode(format: CellFormat): string {
  const symbol = currencySymbol(format.currencyCode);
  const decimals = format.decimalPlaces ?? 2;
  const core = decimals > 0 ? `#,##0.${'0'.repeat(decimals)}` : '#,##0';
  return `_("${symbol}"* ${core}_);_("${symbol}"* \\(${core}\\);_("${symbol}"* "-"??_);_(@_)`;
}

const FRACTION_DENOMINATORS: Record<string, string> = {
  upToOne: '?/?',
  upToTwo: '?/??',
  upToThree: '?/???',
  asHalves: '?/2',
  asQuarters: '?/4',
  asEighths: '?/8',
  asSixteenths: '?/16',
  asTenths: '?/10',
  asHundredths: '?/100',
};

function fractionCode(fractionType: CellFormat['fractionType']): string {
  return `# ${FRACTION_DENOMINATORS[fractionType ?? 'upToOne']}`;
}

const DATE_CODES: Record<string, string> = {
  'MM/DD/YYYY': 'mm/dd/yyyy',
  'DD-MM-YYYY': 'dd-mm-yyyy',
  'YYYY-MM-DD': 'yyyy-mm-dd',
  'Month DD YYYY': 'mmmm dd yyyy',
};

function dateCode(dateFormat: string | undefined): string {
  return DATE_CODES[dateFormat ?? ''] ?? 'm/d/yyyy';
}

const TIME_CODES: Record<string, string> = {
  'HH:mm:ss': 'hh:mm:ss',
  'h:mm AM/PM': 'h:mm AM/PM',
  'HH:mm': 'hh:mm',
};

function timeCode(timeFormat: string | undefined): string {
  return TIME_CODES[timeFormat ?? ''] ?? 'h:mm:ss AM/PM';
}

function durationCode(durationFormat: CellFormat['durationFormat']): string {
  switch (durationFormat) {
    case 'minutes':
      return '[mm]:ss';
    case 'seconds':
      return '[ss]';
    case 'milliseconds':
      return '[ss].000';
    case 'hours':
    default:
      return '[hh]:mm:ss';
  }
}

function invert(map: Record<string, string>): Record<string, string> {
  return Object.fromEntries(Object.entries(map).map(([k, v]) => [v, k]));
}

const REVERSE_DATE_CODES = invert(DATE_CODES); // e.g. 'mm/dd/yyyy' -> 'MM/DD/YYYY'
const REVERSE_TIME_CODES = invert(TIME_CODES);

const REVERSE_FRACTION_CODES = invert(Object.fromEntries(Object.entries(FRACTION_DENOMINATORS).map(([k, v]) => [k, `# ${v}`])));
const REVERSE_DURATION_CODES: Record<string, NonNullable<CellFormat['durationFormat']>> = {
  '[hh]:mm:ss': 'hours', '[mm]:ss': 'minutes', '[ss]': 'seconds', '[ss].000': 'milliseconds',
};

/**
 * The code each built-in number format id stands for (ECMA-376 18.8.30, Part 1). Ids 14 to 22, 27 to 36 and 50 to 58
 * depend on the system locale in Excel; they are given as their en-US forms (the CJK era formats of 27 to 36 and
 * 50 to 58 as the en-US short date, long date or time they correspond to).
 */
export const BUILTIN_NUMBER_FORMAT_CODES: Readonly<Record<number, string>> = {
  0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00',
  5: '"$"#,##0_);\\("$"#,##0\\)', 6: '"$"#,##0_);[Red]\\("$"#,##0\\)', 7: '"$"#,##0.00_);\\("$"#,##0.00\\)', 8: '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)',
  9: '0%', 10: '0.00%', 11: '0.00E+00', 12: '# ?/?', 13: '# ??/??',
  14: 'm/d/yyyy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy', 18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss', 22: 'm/d/yyyy h:mm',
  27: 'm/d/yyyy', 28: 'mmmm d, yyyy', 29: 'mmmm d, yyyy', 30: 'm/d/yyyy', 31: 'mmmm d, yyyy', 32: 'h:mm', 33: 'h:mm:ss', 34: 'h:mm AM/PM', 35: 'h:mm:ss AM/PM', 36: 'm/d/yyyy',
  37: '#,##0_);(#,##0)', 38: '#,##0_);[Red](#,##0)', 39: '#,##0.00_);(#,##0.00)', 40: '#,##0.00_);[Red](#,##0.00)',
  41: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)', 42: '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)',
  43: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)', 44: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)',
  45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mm:ss.0', 48: '##0.0E+0', 49: '@',
  50: 'm/d/yyyy', 51: 'mmmm d, yyyy', 52: 'h:mm AM/PM', 53: 'h:mm:ss AM/PM', 54: 'mmmm d, yyyy', 55: 'h:mm AM/PM', 56: 'h:mm:ss AM/PM', 57: 'm/d/yyyy', 58: 'mmmm d, yyyy',
};

// exceljs resolves built-in ids itself and reports the ECMA table's spelling of ids 14 and 22 ("mm-dd-yy" and
// 'm/d/yy "h":mm'); Excel shows both in the system date format, which is m/d/yyyy [h:mm] in en-US.
const EXCELJS_SPELLING: Record<string, string> = {
  'mm-dd-yy': BUILTIN_NUMBER_FORMAT_CODES[14],
  'm/d/yy "h":mm': BUILTIN_NUMBER_FORMAT_CODES[22],
};

/** The code as the editor stores it: the spelling the importer library gives a built-in id is replaced by what the id means. */
export function normalizeExcelNumberFormatCode(code: string): string {
  return EXCELJS_SPELLING[code] ?? code;
}

const CURRENCY_CODE_FOR_SYMBOL: Record<string, string> = {
  '$': 'USD', '€': 'EUR', '£': 'GBP', '¥': 'JPY', '₹': 'INR', '₩': 'KRW',
};

function countDecimalZeros(code: string): number {
  const match = code.match(/\.(0+)/);
  return match ? match[1].length : 0;
}

function findCurrencySymbol(code: string): string | undefined {
  const quoted = code.match(/"([^"]+)"/);
  if (quoted && CURRENCY_CODE_FOR_SYMBOL[quoted[1]]) return quoted[1];
  const bare = code.match(/[$€£¥₹₩]/);
  return bare?.[0];
}

/** The structured format a code most likely means; only used when it writes the same code back (see below). */
function structuredCandidate(code: string): CellFormat | undefined {
  if (REVERSE_DATE_CODES[code]) return { type: 'date', dateFormat: REVERSE_DATE_CODES[code] };
  if (REVERSE_TIME_CODES[code]) return { type: 'time', timeFormat: REVERSE_TIME_CODES[code] };
  if (REVERSE_FRACTION_CODES[code]) return { type: 'fraction', fractionType: REVERSE_FRACTION_CODES[code] as CellFormat['fractionType'] };
  if (REVERSE_DURATION_CODES[code]) return { type: 'duration', durationFormat: REVERSE_DURATION_CODES[code] };

  const [datePart, timePart] = code.split(' ');
  if (timePart && REVERSE_DATE_CODES[datePart] && REVERSE_TIME_CODES[timePart]) {
    return { type: 'datetime', dateFormat: REVERSE_DATE_CODES[datePart], timeFormat: REVERSE_TIME_CODES[timePart] };
  }
  if (code.endsWith('%')) return { type: 'percentage', decimalPlaces: countDecimalZeros(code) };
  if (/E\+0+$/i.test(code)) return { type: 'scientific', decimalPlaces: countDecimalZeros(code) };

  if (code.startsWith('_(')) {
    const symbol = findCurrencySymbol(code);
    return { type: 'accounting', currencyCode: symbol ? CURRENCY_CODE_FOR_SYMBOL[symbol] : undefined, decimalPlaces: countDecimalZeros(code) };
  }

  const negativeFormat = code.includes('(') ? 'parentheses' : /\[red\]/i.test(code) ? 'red' : 'minus';
  const currencySymbol = findCurrencySymbol(code);
  if (currencySymbol) {
    return {
      type: 'currency',
      currencyCode: CURRENCY_CODE_FOR_SYMBOL[currencySymbol],
      currencySymbolPosition: code.startsWith('"') || code.startsWith(currencySymbol) ? 'prefix' : 'suffix',
      decimalPlaces: countDecimalZeros(code),
      useThousandsSeparator: code.includes(',0') || code.includes('#,##0'),
      negativeFormat,
    };
  }
  if (/^[#0]/.test(code)) {
    return { type: 'number', decimalPlaces: countDecimalZeros(code), useThousandsSeparator: code.includes(','), negativeFormat };
  }
  return undefined;
}

/**
 * Translate an Excel number format code into a CellFormat. A code that maps exactly onto a structured type (the
 * structured format writes the very same code back) keeps that type; every other code is stored verbatim as `custom`,
 * so the code survives import and export unchanged and the number-format engine decides how it looks. The spelling
 * exceljs gives built-in ids 14 and 22 is first replaced by the en-US code the id means.
 */
export function fromExcelNumberFormatCode(code: string): CellFormat {
  const normalized = normalizeExcelNumberFormatCode(code);
  if (normalized === '' || normalized === 'General') return {};
  const candidate = structuredCandidate(normalized);
  if (candidate && toExcelNumberFormatCode(candidate) === normalized) return candidate;
  return { type: 'custom', pattern: normalized };
}
