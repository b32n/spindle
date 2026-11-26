// Rendering of a parsed number format: pick the section for a value, then lay the value out into its tokens.

import { generalDigits, roundToPlaces, toDecimal } from './digits';
import { serialToDate, type DateSystem } from './dates';
import type { Condition, Section, Token } from './parse';

export interface FormatResult {
  text: string;
  /** Colour named by the section ([Red], [Color12]); the canvas may ignore it. */
  color?: string;
  /** Where a `*x` fill would repeat x to the cell width (the text itself holds no fill characters). */
  fill?: { index: number; char: string };
  /** The code is not a valid Excel format code and was shown as General. */
  invalid?: boolean;
}

export interface FormatOptions {
  dateSystem?: DateSystem;
}

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const WEEKDAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
const OVERFLOW = '####';

function test(c: Condition, v: number): boolean {
  switch (c.op) {
    case '=': return v === c.value;
    case '<>': return v !== c.value;
    case '>': return v > c.value;
    case '<': return v < c.value;
    case '>=': return v >= c.value;
    default: return v <= c.value;
  }
}

/**
 * Which section formats a number, and whether the sign is still to be written (a section for negative numbers
 * shows the magnitude; the literal text in it, such as "-" or parentheses, carries the sign). Without conditions:
 * one section for everything, two = positive and zero; negative, three = positive; negative; zero. With conditions:
 * the first section whose condition holds (a missing one in the first section always holds, in the second it
 * means "negative" when a third section exists, otherwise "anything else"). Index -1: a lone condition not met.
 */
export function selectSection(sections: Section[], value: number): { index: number; minus: boolean } {
  const n = sections.length;
  const c1 = sections[0].cond;
  const c2 = sections[1]?.cond;
  let index: number;
  if (!c1 && !c2) index = n >= 2 && value < 0 ? 1 : n >= 3 && value === 0 ? 2 : 0;
  else if (!c1 || test(c1, value)) index = 0;
  else if (n === 1) index = -1;
  else if (n === 2) index = 1;
  else index = (c2 ? test(c2, value) : value < 0) ? 1 : 2;
  const negativeFirst = !!c1 && (c1.op === '<' || c1.op === '<=') && c1.value === 0;
  return { index, minus: value < 0 && (index === -1 || (index === 0 && !negativeFirst)) };
}

const PAD = { '0': '0', '#': '', '?': ' ' } as const;
const digitChar = (t: Token): string => PAD[(t as { c: '0' | '#' | '?' }).c];

function base(t: Token): string {
  switch (t.t) {
    case 'lit': return t.s;
    case 'pad': return ' ';
    case 'point': return '.';
    case 'percent': return '%';
    case 'slash': return '/';
    default: return '';
  }
}

type Out = Map<number, string>;

/** Right-align `digits` (no leading zeros) into integer placeholders; extra digits go in front of the leftmost one. */
function placeInt(tokens: Token[], idx: number[], digits: string, grouping: boolean, anchor: number, out: Out): void {
  const rev = digits.split('').reverse();
  const count = idx.length;
  const charAt = (k: number): string => (k < rev.length ? rev[k] : digitChar(tokens[idx[count - 1 - k]]));
  const total = Math.max(count, rev.length);
  let last = -1;
  for (let k = 0; k < total; k++) if (charAt(k) !== '') last = k;
  let overflow = '';
  for (let k = total - 1; k >= 0; k--) {
    const text = charAt(k) + (grouping && k > 0 && k % 3 === 0 && k <= last ? ',' : '');
    if (k < count) out.set(idx[count - 1 - k], text);
    else overflow += text;
  }
  if (overflow) out.set(anchor, overflow + (out.get(anchor) ?? base(tokens[anchor])));
}

function placeFraction(tokens: Token[], idx: number[], frac: string, out: Out): void {
  let last = -1;
  for (let j = 0; j < frac.length; j++) if (frac[j] !== '0') last = j;
  idx.forEach((tokenIndex, j) => {
    const c = (tokens[tokenIndex] as { c: '0' | '#' | '?' }).c;
    out.set(tokenIndex, c === '0' || j <= last ? frac[j] : c === '#' ? '' : ' ');
  });
}

interface Body {
  text: string;
  zero: boolean;
  fill?: { index: number; char: string };
}

function compose(tokens: Token[], out: Out): Pick<Body, 'text' | 'fill'> {
  let text = '';
  let fill: Body['fill'];
  tokens.forEach((t, i) => {
    if (t.t === 'fill' && !fill) fill = { index: text.length, char: t.c };
    text += out.get(i) ?? base(t);
  });
  return { text, fill };
}

function scaled(section: Section, abs: number): number {
  return (abs * 100 ** section.percent) / 1000 ** section.scale;
}

function numberBody(s: Section, abs: number): Body {
  const out: Out = new Map();
  const anchor = s.int[0] ?? s.tokens.findIndex((t) => t.t === 'point');
  const { int, frac } = roundToPlaces(toDecimal(scaled(s, abs)), s.fracDigits.length);
  placeInt(s.tokens, s.int, int, s.grouping, Math.max(anchor, 0), out);
  placeFraction(s.tokens, s.fracDigits, frac, out);
  return { ...compose(s.tokens, out), zero: int === '' && !/[1-9]/.test(frac) };
}

function exponentBody(s: Section, abs: number): Body {
  const out: Out = new Map();
  const per = Math.max(1, s.int.length);
  const places = s.fracDigits.length;
  const dec = abs === 0 ? { digits: '0', exp: 0 } : toDecimal(scaled(s, abs));
  let exp = abs === 0 ? 0 : Math.floor(dec.exp / per) * per;
  let mantissa = roundToPlaces({ digits: dec.digits, exp: dec.exp - exp }, places);
  if (mantissa.int.length > per) {
    exp += per;
    mantissa = roundToPlaces({ digits: dec.digits, exp: dec.exp - exp }, places);
  }
  placeInt(s.tokens, s.int, mantissa.int, s.grouping, s.int[0] ?? 0, out);
  placeFraction(s.tokens, s.fracDigits, mantissa.frac, out);
  const e = s.tokens[s.exp!.idx] as Extract<Token, { t: 'exp' }>;
  const digits = String(Math.abs(exp)).padStart(s.exp!.digits.length, '0');
  out.set(s.exp!.idx, (out.get(s.exp!.idx) ?? '') + (e.lower ? 'e' : 'E') + (exp < 0 ? '-' : e.plus ? '+' : ''));
  s.exp!.digits.forEach((i, k) => out.set(i, k === 0 ? digits : ''));
  return { ...compose(s.tokens, out), zero: abs === 0 };
}

/** Best p/q for `frac` with q up to `max`: the smallest denominator among the closest. */
function nearest(frac: number, max: number): [number, number] {
  let best: [number, number] = [0, 1];
  let error = Infinity;
  for (let q = 1; q <= max; q++) {
    const p = Math.round(frac * q);
    const e = Math.abs(frac - p / q);
    if (e < error - 1e-13) {
      best = [p, q];
      error = e;
    }
  }
  return best;
}

function fractionBody(s: Section, abs: number): Body {
  const out: Out = new Map();
  const f = s.fraction!;
  const v = scaled(s, abs);
  const hasWhole = f.whole.length > 0;
  let whole = hasWhole ? Math.floor(v) : 0;
  const rest = v - whole;
  const [ratio, den] = f.fixed ? [Math.round(Number((rest * f.fixed).toPrecision(12))), f.fixed] : nearest(rest, 10 ** f.den.length - 1);
  let num = ratio;
  if (hasWhole && num >= den) {
    whole += 1;
    num = 0;
  }
  const zero = whole === 0 && num === 0;
  const anchor = f.whole[0] ?? f.num[0];
  placeInt(s.tokens, f.whole, whole === 0 && num > 0 ? '' : String(whole), s.grouping, anchor, out);
  const blank = hasWhole && num === 0 && !f.num.some((i) => (s.tokens[i] as { c: string }).c === '0');
  if (blank) {
    for (let i = f.denStart; i < f.denEnd; i++) out.set(i, (s.tokens[i] as { c?: string }).c === '#' || s.tokens[i].t === 'dlit' ? '' : ' ');
  } else {
    placeInt(s.tokens, f.num, String(num), false, f.num[0], out);
    if (f.fixed) for (let i = f.slash + 1; i < f.denEnd; i++) out.set(i, i === f.slash + 1 ? String(f.fixed) : '');
    else f.den.forEach((i, k) => out.set(i, k < String(den).length ? String(den)[k] : digitChar(s.tokens[i])));
  }
  return { ...compose(s.tokens, out), zero };
}

const pad2 = (n: number): string => String(n).padStart(2, '0');

/** Date and time tokens; null when Excel shows #### (a date before 1900 or after 9999). */
function dateBody(s: Section, abs: number, system: DateSystem): Body | null {
  const unit = 10 ** s.subDigits;
  const dayUnits = 86_400 * unit;
  let days = Math.floor(abs);
  let units = Math.round(Number(((abs - days) * dayUnits).toPrecision(15)));
  if (units >= dayUnits) {
    days += 1;
    units -= dayUnits;
  }
  const date = serialToDate(days, system);
  if (s.needsDate && !date) return null;
  const total = days * dayUnits + units;
  const hour = Math.floor(units / (3600 * unit));
  const out: Out = new Map();
  s.tokens.forEach((t, i) => {
    if (t.t === 'ampm') out.set(i, hour < 12 ? t.am : t.pm);
    else if (t.t === 'sub') out.set(i, '.' + String(units % unit).padStart(s.subDigits, '0').slice(0, t.n));
    else if (t.t === 'dt') out.set(i, dateToken(t, { date, hour, units, total, unit, twelve: s.twelveHour }));
  });
  return { ...compose(s.tokens, out), zero: false };
}

interface Clock {
  date: ReturnType<typeof serialToDate>;
  hour: number;
  units: number;
  total: number;
  unit: number;
  twelve: boolean;
}

function dateToken(t: Extract<Token, { t: 'dt' }>, c: Clock): string {
  const { n } = t;
  if (t.elapsed) {
    const per = { h: 3600, mi: 60, s: 1 }[t.unit as 'h' | 'mi' | 's'] * c.unit;
    return String(Math.floor(c.total / per)).padStart(n, '0');
  }
  const date = c.date;
  switch (t.unit) {
    case 'y': return n === 2 ? pad2(date!.year % 100) : String(date!.year).padStart(4, '0');
    case 'mo': return n <= 2 ? (n === 2 ? pad2(date!.month) : String(date!.month)) : n === 5 ? MONTHS[date!.month - 1][0] : n === 3 ? MONTHS[date!.month - 1].slice(0, 3) : MONTHS[date!.month - 1];
    case 'd': return n <= 2 ? (n === 2 ? pad2(date!.day) : String(date!.day)) : n === 3 ? WEEKDAYS[date!.weekday].slice(0, 3) : WEEKDAYS[date!.weekday];
    case 'h': {
      const h = c.twelve ? c.hour % 12 || 12 : c.hour;
      return n === 2 ? pad2(h) : String(h);
    }
    case 'mi': {
      const m = Math.floor(c.units / (60 * c.unit)) % 60;
      return n === 2 ? pad2(m) : String(m);
    }
    default: {
      const sec = Math.floor(c.units / c.unit) % 60;
      return n === 2 ? pad2(sec) : String(sec);
    }
  }
}

function generalBody(s: Section, abs: number, negative: boolean): Body {
  const out: Out = new Map();
  s.tokens.forEach((t, i) => {
    if (t.t === 'general') out.set(i, generalDigits(abs, negative));
  });
  return { ...compose(s.tokens, out), zero: abs === 0 };
}

/** The text of a number in one section, or null for ####. `minus` says the sign is still to be written. */
export function renderNumber(s: Section, value: number, minus: boolean, system: DateSystem): { text: string; fill?: Body['fill'] } | null {
  const abs = Math.abs(value);
  let body: Body | null;
  switch (s.kind) {
    case 'date': body = minus ? null : dateBody(s, abs, system); break;
    case 'general': body = generalBody(s, abs, minus); break;
    case 'text': body = { text: generalDigits(abs, minus), zero: abs === 0 }; break;
    case 'literal': body = { ...compose(s.tokens, new Map()), zero: false }; break;
    default: body = s.exp ? exponentBody(s, abs) : s.fraction ? fractionBody(s, abs) : numberBody(s, abs);
  }
  if (!body) return null;
  const sign = minus && !body.zero ? '-' : '';
  return { text: sign + body.text, fill: body.fill && { index: body.fill.index + sign.length, char: body.fill.char } };
}

export { OVERFLOW };

/** A text value in a section: `@` stands for the text. */
export function renderText(s: Section, text: string): { text: string; fill?: Body['fill'] } {
  const out: Out = new Map();
  s.tokens.forEach((t, i) => {
    if (t.t === 'text') out.set(i, text);
    else if (t.t === 'digit' || t.t === 'point' || t.t === 'percent' || t.t === 'slash') out.set(i, '');
  });
  return compose(s.tokens, out);
}
