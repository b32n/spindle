// Text functions.

import type { EagerFn } from './helpers';
import { flatten, toText, toNum, toBoolean, wildcardSearch } from './helpers';
import { formatValue, isValidFormatCode } from '../../number-format';
import { parseDateString } from '../../utils/format-utils';

const DECIMAL = /^[+-]?(\d+\.?\d*|\.\d+)(e[+-]?\d+)?$/i;
const GROUPED = /^[+-]?\d{1,3}(,\d{3})+(\.\d*)?$/;

/** What Excel's TEXT reads as a number from text: a decimal, with `$`, thousands commas or `%`, or a date. */
function numberFromText(text: string): number | undefined {
  let t = text.trim().replace(/^([+-]?)\$/, '$1');
  const percent = t.endsWith('%');
  if (percent) t = t.slice(0, -1);
  if (GROUPED.test(t)) t = t.replace(/,/g, '');
  if (DECIMAL.test(t)) return Number(t) / (percent ? 100 : 1);
  return percent ? undefined : (parseDateString(text) ?? undefined);
}

export const textFunctions: Record<string, EagerFn> = {
  CONCAT: (args) => flatten(args).map(toText).join(''),

  CONCATENATE: (args) => flatten(args).map(toText).join(''),

  TEXTJOIN: (args) => {
    const delimiter = toText(args[0]);
    const ignoreEmpty = toBoolean(args[1]);
    const parts = flatten(args.slice(2)).map(toText);
    return (ignoreEmpty ? parts.filter((p) => p !== '') : parts).join(delimiter);
  },

  LEFT: (args) => {
    const text = toText(args[0]);
    const n = args[1] !== undefined ? toNum(args[1]) : 1;
    if (n < 0) throw new Error('#VALUE!');
    return text.slice(0, n);
  },

  RIGHT: (args) => {
    const text = toText(args[0]);
    const n = args[1] !== undefined ? toNum(args[1]) : 1;
    if (n < 0) throw new Error('#VALUE!');
    return n === 0 ? '' : text.slice(-n);
  },

  MID: (args) => {
    const text = toText(args[0]);
    const start = toNum(args[1]);
    const length = toNum(args[2]);
    if (start < 1 || length < 0) throw new Error('#VALUE!');
    return text.slice(start - 1, start - 1 + length);
  },

  LEN: (args) => toText(args[0]).length,

  FIND: (args) => {
    const find = toText(args[0]);
    const within = toText(args[1]);
    const start = args[2] !== undefined ? toNum(args[2]) : 1;
    if (start < 1) throw new Error('#VALUE!');
    const idx = within.indexOf(find, start - 1);
    if (idx < 0) throw new Error('#VALUE!');
    return idx + 1;
  },

  SEARCH: (args) => {
    const find = toText(args[0]);
    const within = toText(args[1]);
    const start = args[2] !== undefined ? toNum(args[2]) : 1;
    if (start < 1) throw new Error('#VALUE!');
    const idx = wildcardSearch(within.slice(start - 1), find);
    if (idx < 0) throw new Error('#VALUE!');
    return start + idx;
  },

  SUBSTITUTE: (args) => {
    const text = toText(args[0]);
    const oldText = toText(args[1]);
    const newText = toText(args[2]);
    if (oldText === '') return text;
    if (args[3] === undefined) return text.split(oldText).join(newText);
    const instance = toNum(args[3]);
    if (instance < 1) throw new Error('#VALUE!');
    let pos = -1;
    let from = 0;
    for (let i = 0; i < instance; i++) {
      pos = text.indexOf(oldText, from);
      if (pos < 0) return text;
      from = pos + oldText.length;
    }
    return text.slice(0, pos) + newText + text.slice(pos + oldText.length);
  },

  REPLACE: (args) => {
    const oldText = toText(args[0]);
    const start = toNum(args[1]);
    const numChars = toNum(args[2]);
    const newText = toText(args[3]);
    if (start < 1 || numChars < 0) throw new Error('#VALUE!');
    return oldText.slice(0, start - 1) + newText + oldText.slice(start - 1 + numChars);
  },

  UPPER: (args) => toText(args[0]).toUpperCase(),

  LOWER: (args) => toText(args[0]).toLowerCase(),

  PROPER: (args) =>
    toText(args[0])
      .toLowerCase()
      .replace(/(^|[^a-zA-Z])([a-z])/g, (_m, sep: string, ch: string) => sep + ch.toUpperCase()),

  TRIM: (args) => toText(args[0]).replace(/ +/g, ' ').trim(),

  // TEXT(value, code): the Excel number-format engine. An empty code gives "", an invalid one #VALUE!; text that
  // reads as a number is formatted as that number, other text and booleans pass through (text only via an @ section).
  TEXT: (args) => {
    const code = toText(args[1]);
    if (code === '') return '';
    if (!isValidFormatCode(code)) throw new Error('#VALUE!');
    const value = args[0];
    const text = toText(value); // an error value propagates
    const options = { fullGeneral: true }; // no column to fit: General gives the 15 significant digits
    if (value === null || value === undefined) return formatValue(0, code, options).text;
    const number = typeof value === 'string' ? numberFromText(text) : undefined;
    return formatValue(number ?? value, code, options).text;
  },

  VALUE: (args) => {
    const value = args[0];
    if (typeof value === 'number') return value;
    const text = toText(value).trim();
    if (text === '') return 0;
    let str = text.replace(/[$,\s]/g, '');
    let percent = false;
    if (str.endsWith('%')) {
      percent = true;
      str = str.slice(0, -1);
    }
    const num = Number(str);
    if (str === '' || isNaN(num)) throw new Error('#VALUE!');
    return percent ? num / 100 : num;
  },

  REPT: (args) => {
    const text = toText(args[0]);
    const count = Math.floor(toNum(args[1]));
    if (count < 0) throw new Error('#VALUE!');
    return text.repeat(count);
  },

  CHAR: (args) => {
    const code = Math.trunc(toNum(args[0]));
    if (code < 1 || code > 255) throw new Error('#VALUE!');
    return String.fromCharCode(code);
  },

  CODE: (args) => {
    const text = toText(args[0]);
    if (text === '') throw new Error('#VALUE!');
    return text.charCodeAt(0);
  },

  EXACT: (args) => toText(args[0]) === toText(args[1]),
};
