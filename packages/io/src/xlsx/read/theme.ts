import type { XlsxTheme } from '@b32nio/spindle-shared/export/xlsx';
import { kid, type XmlElement } from '../xml';
import { DEFAULT_THEME_COLORS, parseRgb } from './colors';

/** Theme colour slots in the order of a cell colour's `theme` index (the first two pairs are swapped against the theme part's own order). */
const SLOTS = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];

/** `lastClr` is a system colour's value when last saved; these are the two the Office themes use when it is missing. */
const SYSTEM_COLORS: Record<string, string> = { windowText: '#000000', window: '#ffffff' };

function slotColor(slot: XmlElement | undefined): string | undefined {
  const color = slot?.children[0];
  if (!color) return undefined;
  if (color.name === 'srgbClr') return parseRgb(color.attrs.val);
  if (color.name === 'sysClr') return parseRgb(color.attrs.lastClr) ?? SYSTEM_COLORS[color.attrs.val ?? ''];
  return undefined;
}

/** Read a theme part. Any colour the part does not give (or gives unreadably) is the Office default for its slot; no part gives the Office theme. */
export function parseTheme(root: XmlElement | undefined): XlsxTheme {
  const elements = kid(root, 'themeElements');
  const scheme = kid(elements, 'clrScheme');
  const fonts = kid(elements, 'fontScheme');
  const theme: XlsxTheme = { colors: SLOTS.map((slot, i) => slotColor(kid(scheme, slot)) ?? DEFAULT_THEME_COLORS[i]) };
  const major = kid(kid(fonts, 'majorFont'), 'latin')?.attrs.typeface;
  const minor = kid(kid(fonts, 'minorFont'), 'latin')?.attrs.typeface;
  if (major) theme.majorFont = major;
  if (minor) theme.minorFont = minor;
  return theme;
}
