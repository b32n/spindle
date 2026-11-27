import { isTrue } from '../xml';

// every branch starts differently (a digit run is followed by `.` or nothing, never by more digits), so a failed match cannot retry
// the same digits another way: the cost is linear in the text
const DECIMAL = /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/;

/** Longest number text read (characters): Excel writes at most 17 significant digits, an exponent and a sign; a longer text is not a number the editor holds. */
export const MAX_NUMBER_CHARS = 64;

/** A finite number written as a plain decimal (`12`, `-1.5`, `.5`, `1E-3`; surrounding blanks allowed), or `undefined`: never `Infinity`, `0x10`, `1_0`, an empty text or one over `MAX_NUMBER_CHARS`. */
export function parseDecimal(text: string): number | undefined {
  const t = text.trim();
  if (t.length > MAX_NUMBER_CHARS || !DECIMAL.test(t)) return undefined;
  const n = Number(t);
  return Number.isFinite(n) ? n : undefined;
}

/** A non-negative whole number written in digits only (a shared string or style index), or `undefined`. */
export function parseIndex(text: string): number | undefined {
  const t = text.trim();
  if (!/^\d{1,9}$/.test(t)) return undefined;
  return Number(t);
}

/** A finite number from an attribute, or `undefined` when it is absent, empty or not a number. */
export function numAttr(v: string | undefined): number | undefined {
  return v === undefined ? undefined : parseDecimal(v);
}

/** A whole number from an attribute, or `undefined`. */
export function intAttr(v: string | undefined): number | undefined {
  const n = numAttr(v);
  return n === undefined || !Number.isInteger(n) ? undefined : n;
}

/** An OOXML boolean attribute, or `undefined` when the attribute is absent (so "not written" survives). */
export const boolAttr = (v: string | undefined): boolean | undefined => (v === undefined ? undefined : isTrue(v));

/** Set `key` on `target` unless the value is `undefined` (keeps results free of `undefined` members, so they compare and serialise cleanly). */
export function put<T extends object, K extends keyof T>(target: T, key: K, value: T[K] | undefined): void {
  if (value !== undefined) target[key] = value as T[K];
}
