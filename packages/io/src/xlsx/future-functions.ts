/**
 * Formula text as the file stores it versus as the editor shows it (S4).
 *
 * Excel writes a function newer than the 2007 file format with a prefix (`_xlfn.IFS(...)`, `_xlfn._xlws.SORT(...)`), the
 * parameters of LET and LAMBDA with `_xlpm.`, and two syntax forms as calls: the implicit-intersection operator `@x` is
 * `_xlfn.SINGLE(x)` and the spill reference `A1#` is `_xlfn.ANCHORARRAY(A1)`. `fromFileFormula` turns that into the text a user
 * types, `toFileFormula` does the reverse, so a formula survives open -> save -> open and a file Excel wrote can be read by
 * an engine that knows the function by its plain name. Pure string work over one formula (at most 8192 characters), linear
 * except for nested SINGLE/LAMBDA calls (bounded by the formula length); string literals, quoted sheet names and bracket
 * groups are never rewritten.
 */

/** The prefix a future function is written with, by upper-case plain name. */
export const FUTURE_FUNCTIONS: ReadonlyMap<string, string> = (() => {
  const plain = `ACOT ACOTH AGGREGATE ARABIC ARRAYTOTEXT BASE BETA.DIST BETA.INV BINOM.DIST BINOM.DIST.RANGE BINOM.INV BITAND BITLSHIFT
    BITOR BITRSHIFT BITXOR BYCOL BYROW CEILING.MATH CEILING.PRECISE CHISQ.DIST CHISQ.DIST.RT CHISQ.INV CHISQ.INV.RT CHISQ.TEST
    CHOOSECOLS CHOOSEROWS COMBINA CONCAT CONFIDENCE.NORM CONFIDENCE.T COT COTH COVARIANCE.P COVARIANCE.S CSC CSCH DAYS DECIMAL DROP
    ERF.PRECISE ERFC.PRECISE EXPAND EXPON.DIST F.DIST F.DIST.RT F.INV F.INV.RT F.TEST FILTERXML FLOOR.MATH FLOOR.PRECISE
    FORECAST.ETS FORECAST.ETS.CONFINT FORECAST.ETS.SEASONALITY FORECAST.ETS.STAT FORECAST.LINEAR FORMULATEXT GAMMA GAMMA.DIST
    GAMMA.INV GAMMALN.PRECISE GAUSS HSTACK HYPGEOM.DIST IFNA IFS IMCOSH IMCOT IMCSC IMCSCH IMSEC IMSECH IMSINH IMTAN ISFORMULA
    ISOMITTED ISOWEEKNUM LAMBDA LET LOGNORM.DIST LOGNORM.INV MAKEARRAY MAP MAXIFS MINIFS MODE.MULT MODE.SNGL MUNIT NEGBINOM.DIST
    NORM.DIST NORM.INV NORM.S.DIST NORM.S.INV NUMBERVALUE PDURATION PERCENTILE.EXC PERCENTILE.INC PERCENTRANK.EXC PERCENTRANK.INC
    PERMUTATIONA PHI POISSON.DIST QUARTILE.EXC QUARTILE.INC QUERYSTRING RANDARRAY RANK.AVG RANK.EQ REDUCE RRI SCAN SEC SECH SEQUENCE
    SHEET SHEETS SKEW.P SORTBY STDEV.P STDEV.S SWITCH T.DIST T.DIST.2T T.DIST.RT T.INV T.INV.2T T.TEST TAKE TEXTAFTER TEXTBEFORE
    TEXTJOIN TEXTSPLIT TOCOL TOROW UNICHAR UNICODE UNIQUE VAR.P VAR.S VSTACK WEBSERVICE WEIBULL.DIST WRAPCOLS WRAPROWS XLOOKUP
    XMATCH XOR Z.TEST`;
  const out = new Map<string, string>();
  for (const name of plain.split(/\s+/)) out.set(name, '_xlfn.');
  // the dynamic-array functions that Excel groups as "worksheet" functions carry a second prefix
  for (const name of ['FILTER', 'SORT']) out.set(name, '_xlfn._xlws.');
  return out;
})();

type Kind = 'str' | 'quoted' | 'bracket' | 'word' | 'ws' | 'op';
interface Tok {
  kind: Kind;
  text: string;
}

/** A word is a name, a cell reference, a number or a function name: letters, digits and `_ . $ \`, and any non-ASCII character. */
const isWordChar = (ch: string): boolean => /[A-Za-z0-9_.$\\]/.test(ch) || ch.charCodeAt(0) > 127;

/** Split a formula into tokens; the concatenation of the tokens is the formula. Strings, quoted names and `[...]` groups are single tokens. */
function tokenize(text: string): Tok[] {
  const out: Tok[] = [];
  let i = 0;
  while (i < text.length) {
    const ch = text[i];
    let j = i + 1;
    let kind: Kind = 'op';
    if (ch === '"' || ch === "'") {
      // doubled quotes are escapes; an unterminated literal runs to the end
      for (; j < text.length; j++) {
        if (text[j] === ch) {
          if (text[j + 1] === ch) j++;
          else {
            j++;
            break;
          }
        }
      }
      kind = ch === '"' ? 'str' : 'quoted';
    } else if (ch === '[') {
      // structured-reference and external-workbook text, nested groups included; inside it a quote is an escape, not a delimiter
      let depth = 1;
      for (; j < text.length && depth > 0; j++) {
        if (text[j] === '[') depth++;
        else if (text[j] === ']') depth--;
      }
      kind = 'bracket';
    } else if (isWordChar(ch)) {
      while (j < text.length && isWordChar(text[j])) j++;
      kind = 'word';
    } else if (ch === ' ' || ch === '\t' || ch === '\n' || ch === '\r') {
      while (j < text.length && /[ \t\n\r]/.test(text[j])) j++;
      kind = 'ws';
    }
    out.push({ kind, text: text.slice(i, j) });
    i = j;
  }
  return out;
}

/** Index of the `)` that closes the `(` at `open` (tokens.length when it never closes). */
function matchParen(toks: Tok[], open: number): number {
  let depth = 0;
  for (let i = open; i < toks.length; i++) {
    if (toks[i].kind !== 'op') continue;
    if (toks[i].text === '(') depth++;
    else if (toks[i].text === ')' && --depth === 0) return i;
  }
  return toks.length;
}

/** The tokens of `toks` split at the commas that are not inside parentheses or braces. */
function splitArgs(toks: Tok[]): Tok[][] {
  const args: Tok[][] = [[]];
  let depth = 0;
  for (const t of toks) {
    if (t.kind === 'op') {
      if (t.text === '(' || t.text === '{') depth++;
      else if (t.text === ')' || t.text === '}') depth--;
      else if (t.text === ',' && depth === 0) {
        args.push([]);
        continue;
      }
    }
    args[args.length - 1].push(t);
  }
  return args;
}

const join = (toks: Tok[]): string => toks.map((t) => t.text).join('');
const FUTURE_PREFIX = /^(?:_xlfn\.|_xlws\.)+/i;
const PARAM_PREFIX = /^_xlpm\./i;
const CELL_REF = /^\$?[A-Za-z]{1,3}\$?[0-9]+$/;
/** What `@` may be applied to and `SINGLE` may wrap: a reference, a name or a call, with no operator between its parts. */
const SIMPLE_OPERAND_OPS = new Set([':', '!', '(', ')', '@', '#']);

/** True when `toks` read as one reference-like operand (sheet-qualified cell, range, name, table reference or a call). */
function isSimpleOperand(toks: Tok[]): boolean {
  if (toks.length === 0) return false;
  let depth = 0;
  for (const t of toks) {
    if (t.kind === 'ws' && depth === 0) return false;
    if (t.kind !== 'op') continue;
    if (t.text === '(') depth++;
    else if (t.text === ')') depth--;
    else if (depth === 0 && !SIMPLE_OPERAND_OPS.has(t.text)) return false;
  }
  return depth === 0;
}

/** The formula as the editor holds it: no future-function, parameter or call-form prefixes. */
export function fromFileFormula(formula: string): string {
  return join(fromTokens(tokenize(formula)));
}

function fromTokens(toks: Tok[]): Tok[] {
  const out: Tok[] = [];
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    if (t.kind !== 'word') {
      out.push(t);
      continue;
    }
    const isCall = toks[i + 1]?.kind === 'op' && toks[i + 1].text === '(';
    if (PARAM_PREFIX.test(t.text) && t.text.length > 6) {
      out.push({ kind: 'word', text: t.text.replace(PARAM_PREFIX, '') });
      continue;
    }
    const name = t.text.replace(FUTURE_PREFIX, '');
    if (!isCall || name === t.text || name === '') {
      out.push(t);
      continue;
    }
    const upper = name.toUpperCase();
    if ((upper === 'SINGLE' || upper === 'ANCHORARRAY') && /^_xlfn\./i.test(t.text)) {
      const close = matchParen(toks, i + 1);
      if (close < toks.length) {
        const inner = fromTokens(toks.slice(i + 2, close));
        if (upper === 'SINGLE' && isSimpleOperand(inner)) out.push({ kind: 'op', text: '@' }, ...inner);
        else if (upper === 'ANCHORARRAY' && isAnchorOperand(inner)) out.push(...inner, { kind: 'op', text: '#' });
        // no short form for this operand: the call stays, without its prefix
        else out.push({ kind: 'word', text: name }, toks[i + 1], ...inner, toks[close]);
        i = close;
        continue;
      }
    }
    out.push({ kind: 'word', text: name });
  }
  return out;
}

/** A single cell reference, sheet-qualified or not (what `#` can follow). */
function isAnchorOperand(toks: Tok[]): boolean {
  const n = toks.length;
  if (n === 1) return toks[0].kind === 'word' && CELL_REF.test(toks[0].text);
  return (
    n === 3 &&
    (toks[0].kind === 'word' || toks[0].kind === 'quoted') &&
    toks[1].kind === 'op' &&
    toks[1].text === '!' &&
    toks[2].kind === 'word' &&
    CELL_REF.test(toks[2].text)
  );
}

/** The formula as the file stores it: future functions prefixed, `@x` and `A1#` written as calls, LET/LAMBDA parameters prefixed. */
export function toFileFormula(formula: string): string {
  return join(toTokens(tokenize(formula), new Set()));
}

/** The end (exclusive) of the operand that starts at `from`: `Sheet!A1:B2`, `'My sheet'!A1`, `INDEX(...)`, `Table[col]`; `from` when there is none. */
function operandEnd(toks: Tok[], from: number): number {
  let i = from;
  for (;;) {
    const t = toks[i];
    if (t === undefined || (t.kind !== 'word' && t.kind !== 'quoted')) return i;
    i++;
    if (toks[i]?.kind === 'bracket') i++;
    else if (toks[i]?.kind === 'op' && toks[i].text === '(' && t.kind === 'word') i = Math.min(matchParen(toks, i) + 1, toks.length);
    const next = toks[i];
    if (next?.kind === 'op' && (next.text === '!' || next.text === ':')) i++;
    else return i;
  }
}

function toTokens(toks: Tok[], params: ReadonlySet<string>): Tok[] {
  const out: Tok[] = [];
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    if (t.kind === 'op' && t.text === '@') {
      const end = operandEnd(toks, i + 1);
      if (end > i + 1) {
        out.push({ kind: 'word', text: '_xlfn.SINGLE' }, { kind: 'op', text: '(' }, ...toTokens(toks.slice(i + 1, end), params), {
          kind: 'op',
          text: ')',
        });
        i = end - 1;
        continue;
      }
    }
    if (t.kind === 'op' && t.text === '#') {
      // `A1#` / `Sheet1!A1#`: the operand already written to `out` (a cell reference directly before the `#`, not followed by a letter: `#REF!`)
      const next = toks[i + 1];
      const n = out.length;
      const ref = out[n - 1];
      if (ref?.kind === 'word' && CELL_REF.test(ref.text) && !(next?.kind === 'word' && /^[A-Za-z]/.test(next.text))) {
        const sheet = out[n - 2]?.kind === 'op' && out[n - 2].text === '!' && (out[n - 3]?.kind === 'word' || out[n - 3]?.kind === 'quoted');
        const start = sheet ? n - 3 : n - 1;
        const operand = out.splice(start, n - start);
        out.push({ kind: 'word', text: '_xlfn.ANCHORARRAY' }, { kind: 'op', text: '(' }, ...operand, { kind: 'op', text: ')' });
        continue;
      }
    }
    if (t.kind !== 'word') {
      out.push(t);
      continue;
    }
    const isCall = toks[i + 1]?.kind === 'op' && toks[i + 1].text === '(';
    if (params.has(t.text.toUpperCase())) {
      out.push({ kind: 'word', text: `_xlpm.${t.text}` });
      continue;
    }
    if (!isCall) {
      out.push(t);
      continue;
    }
    const upper = t.text.toUpperCase();
    const prefix = FUTURE_FUNCTIONS.get(upper);
    out.push(prefix !== undefined ? { kind: 'word', text: prefix + t.text } : t);
    if (upper === 'LET' || upper === 'LAMBDA') {
      // the parameter names of the call (LET: every other argument but the last; LAMBDA: every argument but the last) are written `_xlpm.name` wherever the call uses them
      const close = matchParen(toks, i + 1);
      const args = splitArgs(toks.slice(i + 2, close));
      const scoped = new Set(params);
      args.forEach((arg, k) => {
        const isName = upper === 'LET' ? k % 2 === 0 : true;
        const words = arg.filter((a) => a.kind !== 'ws');
        if (k < args.length - 1 && isName && words.length === 1 && words[0].kind === 'word') scoped.add(words[0].text.toUpperCase());
      });
      out.push({ kind: 'op', text: '(' }, ...toTokens(toks.slice(i + 2, close), scoped));
      if (close < toks.length) out.push(toks[close]);
      i = close;
    }
  }
  return out;
}
