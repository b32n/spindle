import type { XlsxSharedString, XlsxTextRun } from '@b32nio/spindle-shared/export/xlsx';
import { kid, kids, type XmlElement } from '../xml';
import type { ColorContext } from './colors';
import { readFont } from './font';

/** `_xHHHH_` is how a control character is stored (`_x000D_` for CR); `_x005F_` is a literal underscore that guards a following `_x...`. */
const ESCAPE = /_x([0-9a-fA-F]{4})_/g;

/** Decode the `_xHHHH_` escapes of a string item. A lone surrogate half is left as written (it has no character of its own). */
export function decodeStringEscapes(text: string): string {
  if (!text.includes('_x')) return text;
  return text.replace(ESCAPE, (whole, hex: string) => {
    const code = parseInt(hex, 16);
    return code >= 0xd800 && code <= 0xdfff ? whole : String.fromCharCode(code);
  });
}

const textOf = (el: XmlElement): string => decodeStringEscapes(kids(el, 't').map((t) => t.text).join(''));

/** One `<si>`: plain text and, when it has `<r>` runs, the runs with their fonts. Phonetic runs (`rPh`) are not text. */
export function readSharedString(si: XmlElement, ctx: ColorContext): XlsxSharedString {
  const runEls = kids(si, 'r');
  if (runEls.length === 0) return { text: textOf(si) };
  const runs: XlsxTextRun[] = runEls.map((r) => {
    const run: XlsxTextRun = { text: textOf(r) };
    const font = readFont(kid(r, 'rPr'), ctx);
    if (font && Object.keys(font).length > 0) run.font = font;
    return run;
  });
  return { text: runs.map((r) => r.text).join(''), runs };
}

/**
 * Push every item of a shared-strings part to `visit` in index order, so a caller can build its own table (or stream into
 * one) without this reader holding the list. Items are the `<si>` children; anything else is skipped.
 */
export function forEachSharedString(root: XmlElement | undefined, ctx: ColorContext, visit: (index: number, item: XlsxSharedString) => void): void {
  kids(root, 'si').forEach((si, i) => visit(i, readSharedString(si, ctx)));
}

/** The items of a shared-strings part, indexed as cells refer to them (`<c t="s"><v>index</v>`). */
export function parseSharedStrings(root: XmlElement | undefined, ctx: ColorContext): XlsxSharedString[] {
  const items: XlsxSharedString[] = [];
  forEachSharedString(root, ctx, (_, item) => items.push(item));
  return items;
}
