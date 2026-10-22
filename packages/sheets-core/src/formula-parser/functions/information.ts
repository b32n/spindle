// Information functions.
//
// The IS* predicates and TYPE / ERROR.TYPE are lazy: they must report on an
// argument that errors (e.g. ISNUMBER(1/0) is FALSE, not an error), which
// means catching errors thrown while the argument is evaluated.
//
// ISREF, ISFORMULA and CELL are not here — they need the argument's
// reference rather than its value, which is a separate engine change.

import type { EagerFn, LazyFn } from './helpers';
import { isErrorValue } from './helpers';

/** Excel ERROR.TYPE codes, longest-prefix matched against the error text. */
const ERROR_CODES: Array<[string, number]> = [
  ['#NULL!', 1],
  ['#DIV/0!', 2],
  ['#VALUE!', 3],
  ['#REF!', 4],
  ['#NAME?', 5],
  ['#NUM!', 6],
  ['#N/A', 7],
];

export const informationEagerFunctions: Record<string, EagerFn> = {
  NA: () => {
    throw new Error('#N/A');
  },

  N: (args) => {
    const v = args[0];
    if (typeof v === 'number') return v;
    if (typeof v === 'boolean') return v ? 1 : 0;
    if (typeof v === 'string' && isErrorValue(v)) throw new Error(v);
    return 0;
  },
};

export const informationLazyFunctions: Record<string, LazyFn> = {
  ISBLANK: (thunks) => {
    try {
      const v = thunks[0]();
      return v == null; // only a truly empty cell; ="" is text, not blank
    } catch {
      return false;
    }
  },

  ISNUMBER: (thunks) => {
    try {
      const v = thunks[0]();
      return typeof v === 'number' && !isNaN(v);
    } catch {
      return false;
    }
  },

  ISTEXT: (thunks) => {
    try {
      const v = thunks[0]();
      return typeof v === 'string' && !isErrorValue(v);
    } catch {
      return false;
    }
  },

  ISLOGICAL: (thunks) => {
    try {
      return typeof thunks[0]() === 'boolean';
    } catch {
      return false;
    }
  },

  ISERROR: (thunks) => {
    try {
      return isErrorValue(thunks[0]());
    } catch {
      return true;
    }
  },

  ISERR: (thunks) => {
    try {
      const v = thunks[0]();
      return isErrorValue(v) && v !== '#N/A';
    } catch (e) {
      return !(e instanceof Error && e.message.startsWith('#N/A'));
    }
  },

  ISNA: (thunks) => {
    try {
      return thunks[0]() === '#N/A';
    } catch (e) {
      return e instanceof Error && e.message.startsWith('#N/A');
    }
  },

  TYPE: (thunks) => {
    let v: unknown;
    try {
      v = thunks[0]();
    } catch {
      return 16;
    }
    if (typeof v === 'number') return 1;
    if (typeof v === 'string') return isErrorValue(v) ? 16 : 2;
    if (typeof v === 'boolean') return 4;
    if (Array.isArray(v)) return 64;
    return 1;
  },

  'ERROR.TYPE': (thunks) => {
    let errorText: string | null = null;
    try {
      const v = thunks[0]();
      if (isErrorValue(v)) errorText = v as string;
    } catch (e) {
      if (e instanceof Error) errorText = e.message;
    }
    if (errorText !== null) {
      for (const [prefix, code] of ERROR_CODES) {
        if (errorText.startsWith(prefix)) return code;
      }
    }
    throw new Error('#N/A');
  },
};
