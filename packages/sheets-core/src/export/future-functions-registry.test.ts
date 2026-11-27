import { FUTURE_FUNCTIONS } from '@b32nio/spindle-io/xlsx';
import { getRegisteredFunctionNames } from '../formula-parser/functions';

// S4: the table of functions that Excel stores with an `_xlfn.` prefix (packages/io/src/xlsx/future-functions.ts) against the
// engine's own function registry. A registry function missing from the table would be written without its prefix and show
// #NAME? in Excel, so a new engine function must be a conscious decision: it goes in the table, or in the list below.

/** Functions of the original file format (Excel 2007): stored with their plain name, so they are not in the table. */
const NO_PREFIX = [
  'ABS', 'ACOS', 'ACOSH', 'ADDRESS', 'AND', 'ASIN', 'ASINH', 'ATAN', 'ATAN2', 'ATANH', 'AVERAGE', 'AVERAGEIF', 'AVERAGEIFS',
  'CEILING', 'CELL', 'CHAR', 'CHOOSE', 'CLEAN', 'CODE', 'COLUMN', 'COLUMNS', 'COMBIN', 'CONCATENATE', 'CORREL', 'COS', 'COSH',
  'COUNT', 'COUNTA', 'COUNTBLANK', 'COUNTIF', 'COUNTIFS', 'COVAR', 'DATE', 'DATEDIF', 'DATEVALUE', 'DAY', 'DAYS360', 'DDB', 'DEGREES',
  'DOLLAR', 'EDATE', 'EFFECT', 'EOMONTH', 'ERROR.TYPE', 'EVEN', 'EXACT', 'EXP', 'FACT', 'FACTDOUBLE', 'FALSE', 'FIND', 'FIXED',
  'FLOOR', 'FORECAST', 'FREQUENCY', 'FV', 'FVSCHEDULE', 'GCD', 'GEOMEAN', 'HLOOKUP', 'HOUR', 'HYPERLINK', 'IF', 'IFERROR', 'INDEX',
  'INDIRECT', 'INT', 'INTERCEPT', 'IPMT', 'IRR', 'ISBLANK', 'ISERR', 'ISERROR', 'ISEVEN', 'ISLOGICAL', 'ISNA', 'ISNONTEXT',
  'ISNUMBER', 'ISODD', 'ISPMT', 'ISREF', 'ISTEXT', 'LARGE', 'LCM', 'LEFT', 'LEN', 'LN', 'LOG', 'LOG10', 'LOOKUP', 'LOWER', 'MATCH', 'MAX',
  'MEDIAN', 'MID', 'MIN', 'MINUTE', 'MIRR', 'MOD', 'MODE', 'MONTH', 'MROUND', 'N', 'NA', 'NETWORKDAYS', 'NETWORKDAYS.INTL', 'NOMINAL',
  'NOT', 'NOW', 'NPER', 'NPV', 'ODD', 'OFFSET', 'OR', 'PERCENTILE', 'PERMUT', 'PI', 'PMT', 'POWER', 'PPMT', 'PRODUCT', 'PROPER', 'PV',
  'QUARTILE', 'QUOTIENT', 'RADIANS', 'RAND', 'RANDBETWEEN', 'RANK', 'RATE', 'REPLACE', 'REPT', 'RIGHT', 'ROMAN', 'ROUND', 'ROUNDDOWN',
  'ROUNDUP', 'ROW', 'ROWS', 'RSQ', 'SEARCH', 'SECOND', 'SIGN', 'SIN', 'SINH', 'SLN', 'SLOPE', 'SMALL', 'SQRT', 'SQRTPI', 'STDEV',
  'STDEVP', 'SUBSTITUTE', 'SUBTOTAL', 'SUM', 'SUMIF', 'SUMIFS', 'SUMPRODUCT', 'SUMSQ', 'SUMX2MY2', 'SUMX2PY2', 'SUMXMY2', 'SYD', 'T',
  'TAN', 'TANH', 'TEXT', 'TIME', 'TIMEVALUE', 'TODAY', 'TRANSPOSE', 'TRIM', 'TRUE', 'TRUNC', 'TYPE', 'UPPER', 'VALUE', 'VAR', 'VARP',
  'VLOOKUP', 'WEEKDAY', 'WEEKNUM', 'WORKDAY', 'WORKDAY.INTL', 'XIRR', 'XNPV', 'YEAR', 'YEARFRAC',
];

/** Registry functions that are not Excel functions at all (nothing to prefix): `SPLIT` is a Google Sheets function. */
const NOT_EXCEL = ['SPLIT'];

/**
 * The names XlsxWriter 3.3.4 (worksheet.py, its `_xlfn` table, checked against Excel files) stores with `_xlfn.`, minus the two
 * syntax forms (`SINGLE`, `ANCHORARRAY`, handled by `fromFileFormula`/`toFileFormula`) and minus `QUERYSTRING`, which is not a
 * function in Excel's own documentation.
 */
const EXCEL_FUTURE = [
  'ACOT', 'ACOTH', 'AGGREGATE', 'ARABIC', 'ARRAYTOTEXT', 'BASE', 'BETA.DIST', 'BETA.INV', 'BINOM.DIST', 'BINOM.DIST.RANGE',
  'BINOM.INV', 'BITAND', 'BITLSHIFT', 'BITOR', 'BITRSHIFT', 'BITXOR', 'BYCOL', 'BYROW', 'CEILING.MATH', 'CEILING.PRECISE',
  'CHISQ.DIST', 'CHISQ.DIST.RT', 'CHISQ.INV', 'CHISQ.INV.RT', 'CHISQ.TEST', 'CHOOSECOLS', 'CHOOSEROWS', 'COMBINA', 'CONCAT',
  'CONFIDENCE.NORM', 'CONFIDENCE.T', 'COT', 'COTH', 'COVARIANCE.P', 'COVARIANCE.S', 'CSC', 'CSCH', 'DAYS', 'DECIMAL', 'DROP',
  'ERF.PRECISE', 'ERFC.PRECISE', 'EXPAND', 'EXPON.DIST', 'F.DIST', 'F.DIST.RT', 'F.INV', 'F.INV.RT', 'F.TEST', 'FILTER', 'FILTERXML',
  'FLOOR.MATH', 'FLOOR.PRECISE', 'FORECAST.ETS', 'FORECAST.ETS.CONFINT', 'FORECAST.ETS.SEASONALITY', 'FORECAST.ETS.STAT',
  'FORECAST.LINEAR', 'FORMULATEXT', 'GAMMA', 'GAMMA.DIST', 'GAMMA.INV', 'GAMMALN.PRECISE', 'GAUSS', 'HSTACK', 'HYPGEOM.DIST', 'IFNA',
  'IFS', 'IMAGE', 'IMCOSH', 'IMCOT', 'IMCSC', 'IMCSCH', 'IMSEC', 'IMSECH', 'IMSINH', 'IMTAN', 'ISFORMULA', 'ISOMITTED', 'ISOWEEKNUM',
  'LAMBDA', 'LET', 'LOGNORM.DIST', 'LOGNORM.INV', 'MAKEARRAY', 'MAP', 'MAXIFS', 'MINIFS', 'MODE.MULT', 'MODE.SNGL', 'MUNIT',
  'NEGBINOM.DIST', 'NORM.DIST', 'NORM.INV', 'NORM.S.DIST', 'NORM.S.INV', 'NUMBERVALUE', 'PDURATION', 'PERCENTILE.EXC',
  'PERCENTILE.INC', 'PERCENTRANK.EXC', 'PERCENTRANK.INC', 'PERMUTATIONA', 'PHI', 'POISSON.DIST', 'QUARTILE.EXC', 'QUARTILE.INC',
  'RANDARRAY', 'RANK.AVG', 'RANK.EQ', 'REDUCE', 'RRI', 'SCAN', 'SEC', 'SECH', 'SEQUENCE', 'SHEET', 'SHEETS', 'SKEW.P', 'SORT',
  'SORTBY', 'STDEV.P', 'STDEV.S', 'SWITCH', 'T.DIST', 'T.DIST.2T', 'T.DIST.RT', 'T.INV', 'T.INV.2T', 'T.TEST', 'TAKE', 'TEXTAFTER',
  'TEXTBEFORE', 'TEXTJOIN', 'TEXTSPLIT', 'TOCOL', 'TOROW', 'UNICHAR', 'UNICODE', 'UNIQUE', 'VALUETOTEXT', 'VAR.P', 'VAR.S', 'VSTACK',
  'WEBSERVICE', 'WEIBULL.DIST', 'WRAPCOLS', 'WRAPROWS', 'XLOOKUP', 'XMATCH', 'XOR', 'Z.TEST',
];

describe('the _xlfn. table against the engine registry', () => {
  const registry = [...getRegisteredFunctionNames()].map(String).sort();

  it('has every registry function that Excel stores with a prefix', () => {
    const missing = registry.filter((name) => EXCEL_FUTURE.includes(name) && !FUTURE_FUNCTIONS.has(name));
    expect(missing).toEqual([]);
  });

  it('accounts for every registry function: in the table, or a known original-format or non-Excel function', () => {
    const unaccounted = registry.filter((name) => !FUTURE_FUNCTIONS.has(name) && !NO_PREFIX.includes(name) && !NOT_EXCEL.includes(name));
    expect(unaccounted).toEqual([]);
  });

  it('does not list a function twice, nor a registry function both with and without a prefix', () => {
    expect(NO_PREFIX.filter((name) => FUTURE_FUNCTIONS.has(name))).toEqual([]);
    expect(NOT_EXCEL.filter((name) => FUTURE_FUNCTIONS.has(name))).toEqual([]);
    expect(new Set(NO_PREFIX).size).toBe(NO_PREFIX.length);
  });

  it('covers every name of the oracle', () => {
    expect(EXCEL_FUTURE.filter((name) => !FUTURE_FUNCTIONS.has(name))).toEqual([]);
  });
});
