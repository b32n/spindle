import type { XlsxFont } from '@b32nio/spindle-shared/export/xlsx';
import { isOn, kid, type XmlElement } from '../xml';
import { readColor, type ColorContext } from './colors';
import { intAttr, numAttr, put } from './values';

const UNDERLINES = new Set(['single', 'double', 'singleAccounting', 'doubleAccounting']);
const VERTICAL_ALIGNS = new Set(['baseline', 'superscript', 'subscript']);

const flag = (font: XmlElement, name: string): boolean | undefined => {
  const el = kid(font, name);
  return el ? isOn(el.attrs.val) : undefined;
};

/**
 * A `<font>` of the styles part or an `<rPr>` of a rich-text run (the run form names the face `rFont`). Only what is written
 * is set, so a run's font says what the run overrides.
 */
export function readFont(el: XmlElement | undefined, ctx: ColorContext): XlsxFont | undefined {
  if (!el) return undefined;
  const font: XlsxFont = {};
  put(font, 'bold', flag(el, 'b'));
  put(font, 'italic', flag(el, 'i'));
  put(font, 'strike', flag(el, 'strike'));
  const u = kid(el, 'u');
  if (u) {
    const val = u.attrs.val ?? 'single';
    if (UNDERLINES.has(val)) font.underline = val as XlsxFont['underline'];
  }
  const va = kid(el, 'vertAlign')?.attrs.val;
  if (va !== undefined && VERTICAL_ALIGNS.has(va)) font.verticalAlign = va as XlsxFont['verticalAlign'];
  put(font, 'size', numAttr(kid(el, 'sz')?.attrs.val));
  put(font, 'name', (kid(el, 'name') ?? kid(el, 'rFont'))?.attrs.val);
  put(font, 'family', intAttr(kid(el, 'family')?.attrs.val));
  put(font, 'charset', intAttr(kid(el, 'charset')?.attrs.val));
  put(font, 'scheme', kid(el, 'scheme')?.attrs.val);
  put(font, 'color', readColor(kid(el, 'color'), ctx));
  return font;
}
