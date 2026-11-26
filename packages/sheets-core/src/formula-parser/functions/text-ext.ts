// Text functions, batch B: CLEAN, NUMBERVALUE, UNICHAR, UNICODE, T, FIXED, DOLLAR, BASE, DECIMAL, ROMAN and ARABIC.
//
// Scalar functions lift over arrays like the rest of the engine. TEXTBEFORE, TEXTAFTER and TEXTSPLIT are deferred until
// Excel-computed files exist (plan S12, D8: LibreOffice has no such function, so there is no oracle).

import type { EagerFn } from './helpers';
import { checkArity, isErrorValue, liftScalar, toBoolean, toNum, toText } from './helpers';

/** A string of decimal digits plus one (`'199'` -> `'200'`, `''` -> `'1'`, `'99'` -> `'100'`). */
function increment(digits: string): string {
  const out = digits.split('');
  let i = out.length - 1;
  while (i >= 0 && out[i] === '9') out[i--] = '0';
  if (i < 0) return '1' + out.join('');
  out[i] = String(Number(out[i]) + 1);
  return out.join('');
}

/** Digits of `abs` (non-negative) rounded half up to `decimals` places: the decimal-exact value Excel shows. */
function roundedDigits(abs: number, decimals: number): { int: string; frac: string } {
  const [mantissa, exp] = abs.toExponential(14).split('e'); // 15 significant digits, as Excel keeps
  let digits = mantissa.replace('.', '');
  const keep = Number(exp) + 1 + decimals; // digits that survive
  if (keep < 0) {
    digits = '';
  } else if (keep < digits.length) {
    const roundUp = digits.charCodeAt(keep) >= 53; // '5'
    digits = digits.slice(0, keep);
    if (roundUp) digits = increment(digits);
  } else {
    digits += '0'.repeat(keep - digits.length);
  }
  // `digits` times 10^-decimals is the value
  if (decimals <= 0) return { int: (digits + '0'.repeat(-decimals)).replace(/^0+/, '') || '0', frac: '' };
  const padded = digits.padStart(decimals + 1, '0');
  return { int: padded.slice(0, padded.length - decimals).replace(/^0+(?=\d)/, ''), frac: padded.slice(-decimals) };
}

/** FIXED / DOLLAR core: the number rounded to `decimals` places, with thousands commas unless `plain`. */
function fixedText(n: number, decimals: number, plain: boolean): { text: string; negative: boolean } {
  if (decimals > 127) throw new Error('#VALUE!');
  const { int, frac } = roundedDigits(Math.abs(n), Math.max(decimals, -400));
  const grouped = plain ? int : int.replace(/\B(?=(\d{3})+$)/g, ',');
  const text = frac === '' ? grouped : `${grouped}.${frac}`;
  return { text, negative: n < 0 && /[1-9]/.test(int + frac) };
}

const ROMAN_CHARS = ['M', 'D', 'C', 'L', 'X', 'V', 'I'];
const ROMAN_VALUES = [1000, 500, 100, 50, 10, 5, 1];

/** A Roman numeral for 1..3999 in one of Excel's five forms (0 classic ... 4 most concise). */
function roman(value: number, form: number): string {
  let out = '';
  let rest = value;
  for (let i = 0; i <= 3; i++) {
    let index = 2 * i;
    const digit = Math.floor(rest / ROMAN_VALUES[index]);
    if (digit % 5 === 4) {
      const subtractFrom = digit === 4 ? index - 1 : index - 2;
      let steps = 0;
      while (steps < form && index < ROMAN_VALUES.length - 1) {
        steps++;
        if (ROMAN_VALUES[subtractFrom] - ROMAN_VALUES[index + 1] <= rest) index++;
        else steps = form;
      }
      out += ROMAN_CHARS[index] + ROMAN_CHARS[subtractFrom];
      rest += ROMAN_VALUES[index] - ROMAN_VALUES[subtractFrom];
    } else {
      let count = digit;
      if (count > 4) {
        out += ROMAN_CHARS[index - 1];
        count -= 5;
      }
      out += ROMAN_CHARS[index].repeat(count);
      rest %= ROMAN_VALUES[index];
    }
  }
  return out;
}

/** The value of one Roman letter, or 0 for anything else. */
function romanValue(ch: string): number {
  const at = ROMAN_CHARS.indexOf(ch);
  return at < 0 ? 0 : ROMAN_VALUES[at];
}

export const textExtFunctions: Record<string, EagerFn> = {
  CLEAN: (args) => {
    checkArity(args, 1);
    // Excel removes the first 32 characters of ASCII (0-31)
    return liftScalar(args, ([text]) => [...toText(text)].filter((ch) => ch.charCodeAt(0) >= 32).join(''));
  },

  NUMBERVALUE: (args) => {
    checkArity(args, 1, 3);
    return liftScalar(args, ([textArg, decimalArg, groupArg]) => {
      if (typeof textArg === 'boolean') throw new Error('#VALUE!');
      const decimal = decimalArg === undefined || toText(decimalArg) === '' ? undefined : [...toText(decimalArg)][0];
      const group = groupArg === undefined || toText(groupArg) === '' ? undefined : [...toText(groupArg)][0];
      const decimalSep = decimal ?? '.';
      // with only a decimal separator given, the default group separator moves out of its way
      const groupSep = group ?? (decimalSep === ',' ? '.' : ',');
      if (decimalSep === groupSep) throw new Error('#VALUE!');
      let text = toText(textArg).replace(/\s/g, '');
      let percents = 0;
      while (text.endsWith('%')) {
        text = text.slice(0, -1);
        percents++;
      }
      if (text === '') {
        if (percents > 0) throw new Error('#VALUE!');
        return 0;
      }
      const at = text.indexOf(decimalSep);
      let whole = at < 0 ? text : text.slice(0, at);
      const fraction = at < 0 ? '' : text.slice(at + 1);
      if (fraction.includes(decimalSep) || fraction.includes(groupSep)) throw new Error('#VALUE!');
      whole = whole.split(groupSep).join('');
      const number = `${whole}${at < 0 ? '' : '.'}${fraction}`;
      if (!/^[+-]?(\d+\.?\d*|\.\d+)(e[+-]?\d+)?$/i.test(number)) throw new Error('#VALUE!');
      return Number(number) / Math.pow(100, percents);
    });
  },

  UNICHAR: (args) => {
    checkArity(args, 1);
    return liftScalar(args, ([n]) => {
      const code = Math.trunc(toNum(n));
      if (code < 1 || code > 0x10ffff || (code >= 0xd800 && code <= 0xdfff)) throw new Error('#VALUE!');
      return String.fromCodePoint(code);
    });
  },

  UNICODE: (args) => {
    checkArity(args, 1);
    return liftScalar(args, ([text]) => {
      const s = toText(text);
      if (s === '') throw new Error('#VALUE!');
      return s.codePointAt(0);
    });
  },

  T: (args) => {
    checkArity(args, 1);
    return liftScalar(args, ([v]) => {
      if (typeof v !== 'string') return '';
      if (isErrorValue(v)) throw new Error(v);
      return v;
    });
  },

  FIXED: (args) => {
    checkArity(args, 1, 3);
    return liftScalar(args, ([n, decimals, noCommas]) => {
      const { text, negative } = fixedText(
        toNum(n),
        decimals === undefined ? 2 : Math.trunc(toNum(decimals)),
        noCommas !== undefined && toBoolean(noCommas)
      );
      return negative ? `-${text}` : text;
    });
  },

  DOLLAR: (args) => {
    checkArity(args, 1, 2);
    return liftScalar(args, ([n, decimals]) => {
      const { text, negative } = fixedText(toNum(n), decimals === undefined ? 2 : Math.trunc(toNum(decimals)), false);
      return negative ? `($${text})` : `$${text}`;
    });
  },

  BASE: (args) => {
    checkArity(args, 2, 3);
    return liftScalar(args, ([n, radixArg, minLength]) => {
      const value = Math.trunc(toNum(n));
      const radix = Math.trunc(toNum(radixArg));
      const length = minLength === undefined ? 0 : Math.trunc(toNum(minLength));
      if (value < 0 || value >= 2 ** 53 || radix < 2 || radix > 36 || length < 0 || length > 255) {
        throw new Error('#NUM!');
      }
      return value.toString(radix).toUpperCase().padStart(length, '0');
    });
  },

  DECIMAL: (args) => {
    checkArity(args, 2);
    return liftScalar(args, ([text, radixArg]) => {
      const radix = Math.trunc(toNum(radixArg));
      const s = toText(text).toUpperCase();
      if (radix < 2 || radix > 36 || s.length > 255) throw new Error('#NUM!');
      let value = 0;
      for (const ch of s) {
        const digit = parseInt(ch, 36);
        if (!/[0-9A-Z]/.test(ch) || digit >= radix) throw new Error('#NUM!');
        value = value * radix + digit;
      }
      return value;
    });
  },

  ROMAN: (args) => {
    checkArity(args, 1, 2);
    return liftScalar(args, ([n, formArg]) => {
      const value = Math.trunc(toNum(n));
      // TRUE is the classic form (0) and FALSE the simplified one (4)
      const form = formArg === undefined ? 0 : typeof formArg === 'boolean' ? (formArg ? 0 : 4) : Math.trunc(toNum(formArg));
      if (value < 0 || value > 3999 || form < 0 || form > 4) throw new Error('#VALUE!');
      return value === 0 ? '' : roman(value, form);
    });
  },

  ARABIC: (args) => {
    checkArity(args, 1);
    return liftScalar(args, ([text]) => {
      let s = toText(text).toUpperCase();
      if (s.length > 255) throw new Error('#VALUE!');
      const negative = s.startsWith('-');
      if (negative) s = s.slice(1);
      let total = 0;
      let previous = 0;
      for (let i = s.length - 1; i >= 0; i--) {
        const v = romanValue(s[i]);
        if (v === 0) throw new Error('#VALUE!');
        total += v < previous ? -v : v;
        previous = Math.max(previous, v);
      }
      return negative ? -total : total;
    });
  },
};
