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

const CURRENCY_CODE_FOR_SYMBOL: Record<string, string> = {
  '$': 'USD', '€': 'EUR', '£': 'GBP', '¥': 'JPY', '₹': 'INR', '₩': 'KRW',
};

function countDecimalZeros(code: string): number {
  const match = code.match(/\.(0+)/);
  return match ? match[1].length : 0;
}

function findCurrencySymbol(code: string): string | undefined {
  // Only a symbol we actually recognize counts — an arbitrary quoted label
  // (e.g. custom pattern `0 "units"`) isn't a currency and must fall through
  // to the other checks instead of being misread as one.
  const quoted = code.match(/"([^"]+)"/);
  if (quoted && CURRENCY_CODE_FOR_SYMBOL[quoted[1]]) return quoted[1];
  const bare = code.match(/[$€£¥₹₩]/);
  return bare?.[0];
}

/**
 * Translate an Excel-syntax number format code back into a CellFormat.
 * Best-effort: exactly reverses whatever toExcelNumberFormatCode produces
 * (so our own round-trip is exact), and does reasonable pattern-sniffing
 * for codes from real spreadsheet apps we didn't generate ourselves —
 * genuinely exotic custom codes fall back to `type: 'custom'`, which
 * preserves the raw code (if imperfectly, since our own custom-pattern
 * renderer is a simple one) rather than losing it.
 */
export function fromExcelNumberFormatCode(code: string): CellFormat {
  const trimmed = code.trim();
  if (!trimmed || trimmed === 'General') return {};

  if (REVERSE_DATE_CODES[trimmed]) return { type: 'date', dateFormat: REVERSE_DATE_CODES[trimmed] };
  if (REVERSE_TIME_CODES[trimmed]) return { type: 'time', timeFormat: REVERSE_TIME_CODES[trimmed] };

  const [datePart, timePart] = trimmed.split(' ');
  if (timePart && REVERSE_DATE_CODES[datePart] && REVERSE_TIME_CODES[timePart]) {
    return { type: 'datetime', dateFormat: REVERSE_DATE_CODES[datePart], timeFormat: REVERSE_TIME_CODES[timePart] };
  }

  // Generic date/time sniffing for codes we didn't author (real Excel files).
  // No exact source pattern to recover, so these fall back to our defaults.
  // Time first: 'm' alone is ambiguous (month vs. minute in Excel's own
  // syntax), so check for the unambiguous time markers (AM/PM, a colon with
  // h/s) before treating a bare 'm' as a date's month.
  if (/am\/pm/i.test(trimmed) || (trimmed.includes(':') && /[hs]/i.test(trimmed))) {
    return { type: 'time', timeFormat: 'h:mm AM/PM' };
  }
  if (!/[#0]/.test(trimmed) && /[ymd].*[ymd]/i.test(trimmed)) {
    return { type: 'date', dateFormat: 'MM/DD/YYYY' };
  }

  if (trimmed.endsWith('%')) {
    return { type: 'percentage', decimalPlaces: countDecimalZeros(trimmed) };
  }

  if (/E\+0+$/i.test(trimmed)) {
    return { type: 'scientific', decimalPlaces: countDecimalZeros(trimmed) };
  }

  if (trimmed.startsWith('_(')) {
    const symbol = findCurrencySymbol(trimmed);
    return {
      type: 'accounting',
      currencyCode: symbol ? CURRENCY_CODE_FOR_SYMBOL[symbol] : undefined,
      decimalPlaces: countDecimalZeros(trimmed),
    };
  }

  const currencySymbol = findCurrencySymbol(trimmed);
  if (currencySymbol) {
    return {
      type: 'currency',
      currencyCode: CURRENCY_CODE_FOR_SYMBOL[currencySymbol],
      currencySymbolPosition: trimmed.trim().startsWith('"') || trimmed.trim().startsWith(currencySymbol) ? 'prefix' : 'suffix',
      decimalPlaces: countDecimalZeros(trimmed),
      useThousandsSeparator: trimmed.includes(',0') || trimmed.includes('#,##0'),
      negativeFormat: trimmed.includes('(') ? 'parentheses' : /\[red\]/i.test(trimmed) ? 'red' : 'minus',
    };
  }

  if (/^[#0]/.test(trimmed)) {
    return {
      type: 'number',
      decimalPlaces: countDecimalZeros(trimmed),
      useThousandsSeparator: trimmed.includes(','),
      negativeFormat: trimmed.includes('(') ? 'parentheses' : /\[red\]/i.test(trimmed) ? 'red' : 'minus',
    };
  }

  return { type: 'custom', pattern: code };
}
