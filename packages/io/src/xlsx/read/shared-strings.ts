import type { XlsxSharedString, XlsxTextRun } from '@b32nio/spindle-shared/export/xlsx';
import { kid, kids, type XmlElement } from '../xml';
import { flat } from './capped';
import { scanXml } from './sax';
import { TreeBuilder } from './tree';
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

const textOf = (el: XmlElement): string => flat(decodeStringEscapes(kids(el, 't').map((t) => t.text).join('')));

/** Most elements one rich-text run is read with (a run is `r`, its `rPr` and a few property elements: Excel writes up to about a dozen); the rest of a longer one is left out. */
export const MAX_RUN_ELEMENTS = 256;

/** One `<r>` as a run: its text and, when `rPr` sets anything, the font. */
function readRun(r: XmlElement, ctx: ColorContext): XlsxTextRun {
  const run: XlsxTextRun = { text: textOf(r) };
  const font = readFont(kid(r, 'rPr'), ctx);
  if (font && Object.keys(font).length > 0) run.font = font;
  return run;
}

export interface SharedStringOptions {
  /** Most items read (`ImportLimits.maxTotalCells`: the cells kept cannot name more); the scan ends at the first one over it. */
  maxItems?: number;
  /** Most rich-text runs kept over the part (`ImportLimits.maxPartElements`); the runs past it are joined into their item's text without their fonts. */
  maxRuns?: number;
  /** `damaged`: not well-formed (what precedes the damage was read); `items`/`runs`: a limit above was passed. */
  problem?: (what: 'damaged' | 'items' | 'runs') => void;
}

/**
 * Push every item of a shared-strings part to `visit` in index order, reading the text as events: memory is what `visit` keeps plus
 * one run at a time, never a tree of the part (a part of five million items is five million small objects, not several gigabytes).
 * An item is an `<si>` child of the root; anything else is skipped. An item is plain text (its direct `<t>`s) or, with `<r>`
 * runs, the runs and their fonts; phonetic runs (`rPh`) are not text. A cut-off item at the damage is not reported.
 */
export function forEachSharedString(xml: string | undefined, ctx: ColorContext, visit: (index: number, item: XlsxSharedString) => void, options: SharedStringOptions = {}): void {
  if (xml === undefined) return;
  const maxItems = options.maxItems ?? Infinity;
  let runsLeft = options.maxRuns ?? Infinity;
  let depth = 0;
  let index = 0;
  let inItem = false;
  let sink = false;
  let plain = '';
  let runs: XlsxTextRun[] = [];
  let runCount = 0;
  let joined = '';
  let run: TreeBuilder | undefined;
  let overItems = false;
  let overRuns = false;
  const wellFormed = scanXml(xml, {
    open(name, attrs) {
      depth++;
      if (run) run.start(name, attrs);
      else if (depth === 2 && name === 'si') {
        if (index >= maxItems) {
          overItems = true;
          return;
        }
        inItem = true;
        plain = '';
        runs = [];
        runCount = 0;
        joined = '';
      } else if (inItem && depth === 3) {
        if (name === 'r') {
          run = new TreeBuilder({ left: MAX_RUN_ELEMENTS });
          run.start(name, attrs);
        } else if (name === 't') sink = true;
      }
    },
    close(name) {
      const level = depth--;
      if (run) {
        run.end();
        if (run.open) return;
        const root = run.root as XmlElement;
        run = undefined;
        runCount++;
        if (runsLeft > 0) {
          runsLeft--;
          runs.push(readRun(root, ctx));
        } else {
          overRuns = true;
          joined += textOf(root);
        }
      } else if (inItem && level === 3 && name === 't') sink = false;
      else if (inItem && level === 2) {
        inItem = false;
        // with runs the item is its runs (the direct text of an item that has runs is not shown); runs all past the limit are plain text
        if (runCount === 0) visit(index++, { text: flat(decodeStringEscapes(plain)) });
        else if (runs.length === 0) visit(index++, { text: flat(joined) });
        else visit(index++, { text: flat(runs.map((r) => r.text).join('') + joined), runs });
      }
    },
    text(chunk) {
      if (run) run.text(chunk);
      else if (sink) plain += chunk;
    },
    halted: () => overItems,
  });
  if (overItems) options.problem?.('items');
  else if (!wellFormed) options.problem?.('damaged');
  if (overRuns) options.problem?.('runs');
}

/** The items of a shared-strings part, indexed as cells refer to them (`<c t="s"><v>index</v>`). */
export function parseSharedStrings(xml: string | undefined, ctx: ColorContext, options: SharedStringOptions = {}): XlsxSharedString[] {
  const items: XlsxSharedString[] = [];
  forEachSharedString(xml, ctx, (_, item) => items.push(item), options);
  return items;
}
