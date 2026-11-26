/**
 * What a formula's text uses that the editor's formula engine cannot evaluate or keep. Pure string analysis:
 * the io package has no engine, so function support comes in as a set of names from the caller.
 */

export interface FormulaFacts {
  /** Function names as written in the file, `_xlfn.`/`_xlws.` stripped, upper-cased, in first-seen order, no duplicates. */
  functions: string[];
  structuredRef: boolean;
  /** A reference into another workbook (`[1]Sheet1!A1`). */
  externalRef: boolean;
  /** `A:A` or `1:1`, with or without a sheet prefix. */
  wholeColumnOrRow: boolean;
}

/** Future-function prefixes Excel writes in front of names newer than the 2007 file format. */
const PREFIX = /^(?:_xlfn\.|_xlws\.)+/i;

/** Functions that recalculate on every pass (the volatile set, plus the ones that read the environment). */
export const VOLATILE_FUNCTIONS: ReadonlySet<string> = new Set([
  'NOW',
  'TODAY',
  'RAND',
  'RANDBETWEEN',
  'RANDARRAY',
  'OFFSET',
  'INDIRECT',
  'INFO',
]);

/** `name` without the future-function prefixes, upper-cased: the key to look a function up by. */
export function normalizeFunctionName(name: string): string {
  return name.replace(PREFIX, '').toUpperCase();
}

/**
 * Blank out string literals (`"..."`) and, with `quoted`, quoted sheet names (`'...'`), each with doubled-quote
 * escapes, so what remains is only operators, names, references and calls. With `brackets` the `[...]` groups
 * (structured-reference and external-workbook text, nested ones included) are blanked too; inside them a quote
 * is an escape character, not a delimiter.
 */
function mask(text: string, quoted: boolean, brackets: boolean): string {
  if (!/["'[]/.test(text)) return text;
  let out = '';
  let depth = 0;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (depth === 0 && (ch === '"' || (quoted && ch === "'"))) {
      let j = i + 1;
      for (; j < text.length; j++) {
        if (text[j] === ch) {
          if (text[j + 1] === ch) j++;
          else break;
        }
      }
      out += ' ';
      i = j;
    } else if (brackets && ch === '[') {
      depth++;
      out += ' ';
    } else if (brackets && ch === ']' && depth > 0) {
      depth--;
      out += ' ';
    } else out += depth > 0 ? ' ' : ch;
  }
  return out;
}

const CALL = /(?<![A-Za-z0-9_.$])([A-Za-z_\\][A-Za-z0-9_.]*)\s*\(/g;
const WHOLE = /(?<![A-Za-z0-9_$.])(?:\$?[A-Z]{1,3}:\$?[A-Z]{1,3}|\$?\d+:\$?\d+)(?![A-Za-z0-9_(])/i;
const STRUCTURED = /(?:[A-Za-z_\\][A-Za-z0-9_.\\]*\[|\[#|\[@)/;
const EXTERNAL = /\[\d+\]/;

export function scanFormula(text: string): FormulaFacts {
  const noStrings = mask(text, false, false);
  const noQuoted = mask(text, true, false);
  const plain = mask(text, true, true);
  const functions: string[] = [];
  for (const m of plain.matchAll(CALL)) {
    // `_xlpm.name(...)` calls a LET/LAMBDA parameter, which is a name and not a function.
    if (/^_xlpm\./i.test(m[1])) continue;
    const fn = normalizeFunctionName(m[1]);
    if (!functions.includes(fn)) functions.push(fn);
  }
  return {
    functions,
    structuredRef: STRUCTURED.test(noQuoted),
    externalRef: EXTERNAL.test(noStrings),
    wholeColumnOrRow: WHOLE.test(plain),
  };
}
