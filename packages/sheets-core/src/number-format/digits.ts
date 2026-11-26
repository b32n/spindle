// Decimal digit handling for the number-format engine. Excel works on the 15 significant decimal digits of a
// double, not on its binary expansion: 1.005 shows as 1.01 with two decimals (toFixed gives 1.00), and 1e22 shows as
// 10000000000000000000000 with a zero format. Everything here is string arithmetic on those digits.

/** value = d0.d1d2... x 10^exp, `digits` has no trailing zeros ("0" for zero). */
export interface Decimal {
  digits: string;
  exp: number;
}

/** The 15 significant digits of a finite, non-negative number. */
export function toDecimal(value: number): Decimal {
  if (value === 0) return { digits: '0', exp: 0 };
  const [mantissa, exp] = value.toExponential(14).split('e');
  return { digits: mantissa.replace('.', '').replace(/0+$/, ''), exp: Number(exp) };
}

function increment(digits: string): string {
  const out = digits.split('');
  for (let i = out.length - 1; i >= 0; i--) {
    if (out[i] !== '9') {
      out[i] = String.fromCharCode(out[i].charCodeAt(0) + 1);
      return out.join('');
    }
    out[i] = '0';
  }
  return '1' + out.join('');
}

/**
 * Round to `places` decimals, half away from zero. `int` has no leading zeros (empty for a value below one half
 * unit of the last place), `frac` is exactly `places` digits long.
 */
export function roundToPlaces(dec: Decimal, places: number): { int: string; frac: string } {
  const { digits, exp } = dec;
  let intPart: string;
  let fracPart: string;
  if (exp >= 0) {
    intPart = digits.slice(0, exp + 1).padEnd(exp + 1, '0');
    fracPart = digits.slice(exp + 1);
  } else {
    intPart = '0';
    fracPart = '0'.repeat(-exp - 1) + digits;
  }
  if (fracPart.length <= places) return { int: intPart.replace(/^0+/, ''), frac: fracPart.padEnd(places, '0') };
  let all = intPart + fracPart.slice(0, places);
  if (fracPart.charCodeAt(places) >= 53) all = increment(all); // '5'
  const intLength = all.length - places;
  return { int: all.slice(0, intLength).replace(/^0+/, ''), frac: all.slice(intLength) };
}

/**
 * Excel's General format. Excel fits the number into the column; this uses the width of a default column,
 * 11 characters including the sign: whole numbers up to 11 digits, otherwise as many decimals as fit, and E+nn
 * when the number is 1e11 or more or below 1e-4. Trailing zeros are dropped.
 */
export function generalDigits(abs: number, negative: boolean, full = false): string {
  if (abs === 0) return '0';
  if (full) return fullDigits(abs);
  const width = negative ? 10 : 11;
  const dec = toDecimal(abs);
  if (dec.exp >= -4 && dec.exp < width) {
    const places = Math.max(0, width - 1 - Math.max(dec.exp + 1, 1));
    const { int, frac } = roundToPlaces(dec, places);
    if (int.length <= width) {
      const trimmed = frac.replace(/0+$/, '');
      return (int || '0') + (trimmed ? '.' + trimmed : '');
    }
  }
  return scientific(dec, width);
}

function scientific(dec: Decimal, width: number): string {
  let exp = dec.exp;
  const places = Math.max(0, width - 1 - (String(Math.abs(exp)).padStart(2, '0').length + 1) - 2);
  let { int, frac } = roundToPlaces({ digits: dec.digits, exp: 0 }, places);
  if (int.length > 1) {
    exp += 1;
    ({ int, frac } = roundToPlaces({ digits: '1', exp: 0 }, places));
  }
  const trimmed = frac.replace(/0+$/, '');
  return `${int}${trimmed ? '.' + trimmed : ''}E${exp < 0 ? '-' : '+'}${String(Math.abs(exp)).padStart(2, '0')}`;
}

/**
 * General without a column to fit (what `TEXT(x,"General")` and string building give): the 15 significant digits,
 * written out the way `String(number)` does, as `E+nn` from 1e21 and below 1e-6.
 */
function fullDigits(abs: number): string {
  const dec = toDecimal(abs);
  if (dec.exp >= 21 || dec.exp < -6) {
    const frac = dec.digits.slice(1);
    return `${dec.digits[0]}${frac ? '.' + frac : ''}E${dec.exp < 0 ? '-' : '+'}${String(Math.abs(dec.exp)).padStart(2, '0')}`;
  }
  const { int, frac } = roundToPlaces(dec, Math.max(0, dec.digits.length - dec.exp - 1));
  return (int || '0') + (frac ? '.' + frac : '');
}
