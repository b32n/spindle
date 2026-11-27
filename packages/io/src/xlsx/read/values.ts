import { isTrue } from '../xml';

/** A finite number from an attribute, or `undefined` when it is absent, empty or not a number. */
export function numAttr(v: string | undefined): number | undefined {
  if (v === undefined || v.trim() === '') return undefined;
  const n = Number(v);
  return Number.isFinite(n) ? n : undefined;
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
