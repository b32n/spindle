// Parser for Excel number-format codes (ECMA-376 Part 1, 18.8.31). A code is up to four sections separated by
// semicolons (positive; negative; zero; text); each section is a run of tokens. Anything Excel would reject (an
// unterminated quote or bracket, a letter that is not a token, more than four sections ...) makes the whole code
// invalid: `compile` never throws, it returns `ok: false` and the caller decides (General for display, #VALUE! for TEXT).

export type Op = '=' | '<>' | '>' | '<' | '>=' | '<=';

export interface Condition {
  op: Op;
  value: number;
}

/** Year, month, day, hour, minute, second. `m` is month or minute until `resolveMinutes` has looked at its neighbours. */
export type DateUnit = 'y' | 'mo' | 'd' | 'h' | 'mi' | 's';

export type Token =
  | { t: 'lit'; s: string }
  | { t: 'dlit'; s: string } // an unquoted digit 1-9, a literal except as the fixed denominator of a fraction
  | { t: 'digit'; c: '0' | '#' | '?' }
  | { t: 'point' }
  | { t: 'thousands' }
  | { t: 'scale' }
  | { t: 'percent' }
  | { t: 'exp'; plus: boolean; lower: boolean }
  | { t: 'slash' }
  | { t: 'text' }
  | { t: 'general' }
  | { t: 'fill'; c: string }
  | { t: 'pad'; c: string }
  | { t: 'dt'; unit: DateUnit | 'm'; n: number; elapsed: boolean }
  | { t: 'ampm'; am: string; pm: string }
  | { t: 'sub'; n: number }; // fractional seconds: .0 .00 .000

export type SectionKind = 'number' | 'date' | 'text' | 'general' | 'literal';

export interface FractionInfo {
  whole: number[]; // token indexes of the whole-part placeholders
  num: number[];
  den: number[]; // variable denominator placeholders
  fixed?: number; // fixed denominator (?/8)
  slash: number;
  denStart: number; // token index range [denStart, denEnd) (numerator to denominator) that is blanked when the fraction part is zero
  denEnd: number;
}

export interface Section {
  tokens: Token[];
  kind: SectionKind;
  cond?: Condition;
  color?: string;
  locale?: string; // the LCID of a [$...-LCID] token, kept as a record
  hasText: boolean; // contains @
  percent: number;
  scale: number;
  grouping: boolean;
  point: boolean;
  int: number[]; // token indexes of integer digit placeholders
  fracDigits: number[]; // token indexes of digit placeholders after the decimal point
  exp?: { idx: number; digits: number[] };
  fraction?: FractionInfo;
  twelveHour: boolean;
  subDigits: number; // fractional-second digits shown (0 when none)
  needsDate: boolean; // shows year, month or day (so a value Excel cannot turn into a date shows ####)
}

export interface CompiledFormat {
  code: string;
  ok: boolean;
  error?: string;
  sections: Section[];
}

class FormatSyntaxError extends Error {}

const fail = (message: string): never => {
  throw new FormatSyntaxError(message);
};

const COLORS = new Set(['black', 'blue', 'cyan', 'green', 'magenta', 'red', 'white', 'yellow']);
const LITERAL_CHARS = new Set(" $-+():!^&'~{}<>=".split(''));

function splitSections(code: string): string[] {
  const sections: string[] = [];
  let current = '';
  for (let i = 0; i < code.length; ) {
    const c = code[i];
    let end = i + 1;
    if (c === '"') {
      end = code.indexOf('"', i + 1) + 1 || fail('unterminated quote');
    } else if (c === '[') {
      end = code.indexOf(']', i + 1) + 1 || fail('unterminated bracket');
    } else if (c === '\\' || c === '_' || c === '*') {
      if (i + 1 >= code.length) fail(`dangling ${c}`);
      end = i + 2;
    } else if (c === ';') {
      sections.push(current);
      current = '';
      i++;
      continue;
    }
    current += code.slice(i, end);
    i = end;
  }
  sections.push(current);
  if (sections.length > 4) fail('more than four sections');
  return sections;
}

function parseCondition(body: string): Condition | undefined {
  const m = /^(<=|>=|<>|=|<|>)\s*(-?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?)$/i.exec(body);
  return m ? { op: m[1] as Op, value: Number(m[2]) } : undefined;
}

interface Pending {
  tokens: Token[];
  cond?: Condition;
  color?: string;
  locale?: string;
}

function readBracket(body: string, out: Pending): void {
  const lower = body.toLowerCase();
  const condition = parseCondition(body);
  const elapsed = /^(h+|m+|s+)$/.exec(lower);
  const color = /^color\s*(\d+)$/.exec(lower);
  if (COLORS.has(lower) || (color && +color[1] >= 1 && +color[1] <= 56)) {
    if (out.color) fail('two colours');
    out.color = color ? `Color${+color[1]}` : body[0].toUpperCase() + lower.slice(1);
  } else if (condition) {
    if (out.cond) fail('two conditions');
    out.cond = condition;
  } else if (elapsed) {
    out.tokens.push({ t: 'dt', unit: ({ h: 'h', m: 'mi', s: 's' } as const)[elapsed[1][0] as 'h' | 'm' | 's'], n: elapsed[1].length, elapsed: true });
  } else if (body[0] === '$') {
    const m = /^(.*)-([0-9a-f]{1,8})$/i.exec(body.slice(1));
    const symbol = m ? m[1] : body.slice(1);
    if (m) out.locale = m[2].toUpperCase();
    if (symbol) out.tokens.push({ t: 'lit', s: symbol });
  } else {
    fail(`unsupported [${body}]`);
  }
}

const isSeconds = (t: Token | undefined): boolean => t?.t === 'dt' && t.unit === 's';

function tokenize(source: string): Pending {
  const out: Pending = { tokens: [] };
  const tokens = out.tokens;
  for (let i = 0; i < source.length; ) {
    const c = source[i];
    const rest = source.slice(i);
    let m: RegExpExecArray | null;
    if (c === '"') {
      const end = source.indexOf('"', i + 1);
      tokens.push({ t: 'lit', s: source.slice(i + 1, end) });
      i = end + 1;
    } else if (c === '\\') {
      tokens.push({ t: 'lit', s: source[i + 1] });
      i += 2;
    } else if (c === '_' || c === '*') {
      tokens.push(c === '_' ? { t: 'pad', c: source[i + 1] } : { t: 'fill', c: source[i + 1] });
      i += 2;
    } else if (c === '[') {
      const end = source.indexOf(']', i);
      readBracket(source.slice(i + 1, end), out);
      i = end + 1;
    } else if ((m = /^general/i.exec(rest))) {
      tokens.push({ t: 'general' });
      i += m[0].length;
    } else if ((m = /^(am\/pm|a\/p)/i.exec(rest))) {
      const [am, pm] = m[0].split('/');
      tokens.push({ t: 'ampm', am, pm });
      i += m[0].length;
    } else if ((m = /^([ymdhs])\1*/i.exec(rest))) {
      const unit = m[0][0].toLowerCase();
      const n = m[0].length;
      const dt = { y: 'y', m: 'm', d: 'd', h: 'h', s: 's' }[unit] as DateUnit | 'm';
      tokens.push({ t: 'dt', unit: dt, n: dt === 'y' ? (n <= 2 ? 2 : 4) : dt === 'd' ? Math.min(n, 4) : dt === 'm' ? (n > 5 ? 4 : n) : Math.min(n, 2), elapsed: false });
      i += n;
    } else if ((m = /^e[+-]/i.exec(rest))) {
      tokens.push({ t: 'exp', plus: m[0][1] === '+', lower: m[0][0] === 'e' });
      i += 2;
    } else if (c === '.' && (m = /^\.(0+)/.exec(rest)) && isSeconds(tokens[tokens.length - 1])) {
      tokens.push({ t: 'sub', n: m[1].length });
      i += m[0].length;
    } else {
      tokens.push(simple(c));
      i++;
    }
  }
  return out;
}

function simple(c: string): Token {
  if (c === '0' || c === '#' || c === '?') return { t: 'digit', c };
  if (c >= '1' && c <= '9') return { t: 'dlit', s: c };
  if (c === '.') return { t: 'point' };
  if (c === ',') return { t: 'thousands' };
  if (c === '%') return { t: 'percent' };
  if (c === '/') return { t: 'slash' };
  if (c === '@') return { t: 'text' };
  if (LITERAL_CHARS.has(c) || c.charCodeAt(0) > 127) return { t: 'lit', s: c };
  return fail(`unexpected ${JSON.stringify(c)}`);
}

/** In a date section `m` after an hour or before a second means minutes. */
function resolveMinutes(tokens: Token[]): void {
  const dates = tokens.filter((t): t is Extract<Token, { t: 'dt' }> => t.t === 'dt');
  dates.forEach((t, i) => {
    if (t.unit !== 'm') return;
    const prev = dates[i - 1];
    const next = dates[i + 1];
    const minute = (prev?.unit === 'h') || (next?.unit === 's');
    t.unit = minute && t.n <= 2 ? 'mi' : 'mo';
  });
}

const isDigit = (t: Token | undefined): boolean => t?.t === 'digit';

function section(source: string): Section {
  const pending = tokenize(source);
  let tokens = pending.tokens;
  const isDate = tokens.some((t) => t.t === 'dt' || t.t === 'ampm' || t.t === 'sub');
  if (!isDate && pending.locale === 'F800') tokens = tokenize('dddd, mmmm d, yyyy').tokens.concat(tokens);
  if (!isDate && pending.locale === 'F400') tokens = tokenize('h:mm:ss AM/PM').tokens.concat(tokens);
  const date = tokens.some((t) => t.t === 'dt' || t.t === 'ampm');
  const s: Section = {
    tokens, kind: 'literal', cond: pending.cond, color: pending.color, locale: pending.locale, hasText: false, percent: 0, scale: 0,
    grouping: false, point: false, int: [], fracDigits: [], twelveHour: false, subDigits: 0, needsDate: false,
  };
  s.hasText = tokens.some((t) => t.t === 'text');
  if (date) return dateSection(s);
  if (tokens.some((t) => t.t === 'general')) {
    if (tokens.some((t) => isDigit(t) || t.t === 'point' || t.t === 'exp')) fail('General mixed with digits');
    s.kind = 'general';
    return literalise(s);
  }
  if (tokens.some((t) => t.t === 'exp' || isDigit(t) || t.t === 'point')) return numberSection(s);
  s.kind = s.hasText ? 'text' : 'literal';
  return literalise(s);
}

/** Outside a number or date, the characters that would be digit-related tokens are plain text. */
function literalise(s: Section): Section {
  s.tokens = s.tokens.map((t): Token => {
    if (t.t === 'thousands') return { t: 'lit', s: ',' };
    if (t.t === 'slash') return { t: 'lit', s: '/' };
    if (t.t === 'percent') return { t: 'lit', s: '%' };
    return t.t === 'dlit' ? { t: 'lit', s: t.s } : t;
  });
  return s;
}

function dateSection(s: Section): Section {
  s.kind = 'date';
  resolveMinutes(s.tokens);
  s.tokens = s.tokens.map((t): Token => {
    switch (t.t) {
      case 'digit': return fail('digit placeholder in a date format');
      case 'exp': return fail('exponent in a date format');
      case 'general': return fail('General in a date format');
      case 'point': return { t: 'lit', s: '.' };
      case 'thousands': return { t: 'lit', s: ',' };
      case 'percent': return { t: 'lit', s: '%' };
      case 'slash': return { t: 'lit', s: '/' };
      case 'dlit': return { t: 'lit', s: t.s };
      default: return t;
    }
  });
  s.twelveHour = s.tokens.some((t) => t.t === 'ampm');
  s.subDigits = Math.max(0, ...s.tokens.map((t) => (t.t === 'sub' ? t.n : 0)));
  s.needsDate = s.tokens.some((t) => t.t === 'dt' && !t.elapsed && (t.unit === 'y' || t.unit === 'mo' || t.unit === 'd'));
  return s;
}

function numberSection(s: Section): Section {
  s.kind = 'number';
  const tokens = s.tokens;
  const expAt = tokens.findIndex((t) => t.t === 'exp');
  const pointAt = tokens.findIndex((t) => t.t === 'point');
  tokens.forEach((t, i) => {
    if (t.t === 'point' && i !== pointAt) tokens[i] = { t: 'lit', s: '.' };
    if (t.t === 'percent') s.percent++;
  });
  // A comma between digit placeholders groups thousands; with no digit after it, it divides by 1000.
  tokens.forEach((t, i) => {
    if (t.t !== 'thousands') return;
    const next = tokens.slice(i + 1).find((x) => x.t !== 'thousands');
    if (!tokens.slice(0, i).some(isDigit)) tokens[i] = { t: 'lit', s: ',' };
    else if (isDigit(next)) s.grouping = true;
    else {
      s.scale++;
      tokens[i] = { t: 'scale' };
    }
  });
  s.point = pointAt >= 0;
  const limit = expAt >= 0 ? expAt : tokens.length;
  tokens.forEach((t, i) => {
    if (!isDigit(t) || i >= limit) return;
    (pointAt >= 0 && i > pointAt ? s.fracDigits : s.int).push(i);
  });
  if (expAt >= 0) {
    s.exp = { idx: expAt, digits: tokens.map((t, i) => (i > expAt && isDigit(t) ? i : -1)).filter((i) => i >= 0) };
    if (!s.exp.digits.length) fail('exponent without digits');
  }
  const slashAt = tokens.findIndex((t, i) => t.t === 'slash' && isDigit(tokens[i - 1]) && (isDigit(tokens[i + 1]) || tokens[i + 1]?.t === 'dlit'));
  if (slashAt >= 0 && expAt < 0 && pointAt < 0) s.fraction = fractionInfo(s, slashAt);
  tokens.forEach((t, i) => {
    if (t.t === 'slash' && i !== slashAt) tokens[i] = { t: 'lit', s: '/' };
    if (t.t === 'dlit' && !(s.fraction && i > slashAt && i < s.fraction.denEnd)) tokens[i] = { t: 'lit', s: t.s };
  });
  return s;
}

function fractionInfo(s: Section, slash: number): FractionInfo {
  const tokens = s.tokens;
  let numStart = slash;
  while (isDigit(tokens[numStart - 1])) numStart--;
  let denEnd = slash + 1;
  while (isDigit(tokens[denEnd]) || tokens[denEnd]?.t === 'dlit') denEnd++;
  const den: number[] = [];
  let fixed = '';
  for (let i = slash + 1; i < denEnd; i++) {
    const t = tokens[i];
    if (t.t === 'dlit') fixed += t.s;
    else if (t.t === 'digit' && fixed) fixed += '0'; // 0/10, 0/100: the zero belongs to the denominator
    else den.push(i);
  }
  const num = s.int.filter((i) => i >= numStart && i < slash);
  const whole = s.int.filter((i) => i < numStart);
  s.int = whole;
  s.fracDigits = [];
  return { whole, num, den: fixed ? [] : den, fixed: fixed ? Number(fixed) : undefined, slash, denStart: numStart, denEnd };
}

/** Parse a code. Never throws; an invalid code gives `ok: false` with one placeholder General section. */
export function parseFormat(code: string): CompiledFormat {
  try {
    const sources = code === '' ? ['General'] : splitSections(code);
    return { code, ok: true, sections: sources.map(section) };
  } catch (e) {
    return { code, ok: false, error: (e as Error).message, sections: [section('General')] };
  }
}
