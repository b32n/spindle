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
