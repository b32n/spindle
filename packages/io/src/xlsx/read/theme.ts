import type { XlsxTheme } from '@b32nio/spindle-shared/export/xlsx';
import { kid, type XmlElement } from '../xml';
import { DEFAULT_THEME_COLORS, hlsToRgb, parseRgb } from './colors';
import { PRESET_COLORS } from './preset-colors';
import { numAttr } from './values';

/** Theme colour slots in the order of a cell colour's `theme` index (the first two pairs are swapped against the theme part's own order). */
const SLOTS = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];

/** `lastClr` is a system colour's value when last saved; these are the two the Office themes use when it is missing. */
const SYSTEM_COLORS: ReadonlyMap<string, string> = new Map([['windowText', '#000000'], ['window', '#ffffff']]);

/** A percentage attribute in 1000ths of a percent (`100000` is 100%) as 0 to 1, or `undefined`. */
function percent(v: string | undefined): number | undefined {
  const n = numAttr(v);
  return n === undefined ? undefined : Math.max(0, Math.min(1, n / 100000));
}

/** scRGB is linear light: the sRGB transfer function gives the byte Office shows. */
const scrgbByte = (linear: number): string => {
  const c = linear <= 0.0031308 ? linear * 12.92 : 1.055 * linear ** (1 / 2.4) - 0.055;
  return Math.round(c * 255).toString(16).padStart(2, '0');
};

/**
 * The colour a scheme slot gives: `srgbClr`, `sysClr` (its `lastClr`), `prstClr`, `scrgbClr` and `hslClr` resolve. `schemeClr`
 * (a reference to another slot, not meaningful inside the colour scheme itself) and anything unreadable do not.
 */
function slotColor(slot: XmlElement | undefined): string | undefined {
  const color = slot?.children[0];
  if (!color) return undefined;
  const a = color.attrs;
  switch (color.name) {
    case 'srgbClr':
      return parseRgb(a.val);
    case 'sysClr':
      return parseRgb(a.lastClr) ?? SYSTEM_COLORS.get(a.val ?? '');
    case 'prstClr':
      return PRESET_COLORS.get(a.val ?? '');
    case 'scrgbClr': {
      const [r, g, b] = [percent(a.r), percent(a.g), percent(a.b)];
      return r === undefined || g === undefined || b === undefined ? undefined : `#${scrgbByte(r)}${scrgbByte(g)}${scrgbByte(b)}`;
    }
    case 'hslClr': {
      const hue = numAttr(a.hue);
      const [s, l] = [percent(a.sat), percent(a.lum)];
      return hue === undefined || s === undefined || l === undefined ? undefined : hlsToRgb((((hue / 60000) % 360) + 360) % 360 / 360, l, s);
    }
    default:
      return undefined;
  }
}

/**
 * Read a theme part. A colour the part does not give is the Office default for its slot; so is one it gives unreadably, but that
 * is pushed to `warnings` (when given) because the colour is then not the theme's own. No part gives the Office theme.
 */
export function parseTheme(root: XmlElement | undefined, warnings?: string[]): XlsxTheme {
  const elements = kid(root, 'themeElements');
  const scheme = kid(elements, 'clrScheme');
  const fonts = kid(elements, 'fontScheme');
  const colors = SLOTS.map((slot, i) => {
    const el = kid(scheme, slot);
    const color = slotColor(el);
    if (color === undefined && el?.children[0]) warnings?.push(`The theme colour ${slot} (${el.children[0].name}) could not be read; the Office default is used.`);
    return color ?? DEFAULT_THEME_COLORS[i];
  });
  const theme: XlsxTheme = { colors };
  const major = kid(kid(fonts, 'majorFont'), 'latin')?.attrs.typeface;
  const minor = kid(kid(fonts, 'minorFont'), 'latin')?.attrs.typeface;
  if (major) theme.majorFont = major;
  if (minor) theme.minorFont = minor;
  return theme;
}
