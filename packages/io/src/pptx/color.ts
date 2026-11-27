import type { PptxColor, PptxColorMap, PptxThemeSlot } from '@b32nio/spindle-shared/export/pptx';
import { kid, type XmlElement } from '../xlsx/xml';
import { pct } from './units';

export const THEME_SLOTS: readonly PptxThemeSlot[] = ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
const SLOT_SET: ReadonlySet<string> = new Set(THEME_SLOTS);

/** The map every master starts from: background and text roles point at the light and dark theme colours. */
export const DEFAULT_COLOR_MAP: PptxColorMap = {
  bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2',
  accent1: 'accent1', accent2: 'accent2', accent3: 'accent3', accent4: 'accent4', accent5: 'accent5', accent6: 'accent6',
  hlink: 'hlink', folHlink: 'folHlink',
};

/** `p:clrMap` of a master: every role it names a valid slot for; the others keep the default. */
export function parseColorMap(el: XmlElement | undefined): PptxColorMap {
  const map = { ...DEFAULT_COLOR_MAP };
  for (const role of Object.keys(map) as (keyof PptxColorMap)[]) {
    const slot = el?.attrs[role];
    if (slot !== undefined && SLOT_SET.has(slot)) map[role] = slot as PptxThemeSlot;
  }
  return map;
}

/** The names the 16 basic CSS colours use in `a:prstClr` (the rest of the preset names are rare and fall back to black). */
const PRESET_COLORS: Record<string, string> = {
  black: '000000', white: 'FFFFFF', red: 'FF0000', green: '008000', blue: '0000FF', yellow: 'FFFF00', gray: '808080', grey: '808080',
  orange: 'FFA500', purple: '800080', pink: 'FFC0CB', brown: 'A52A2A', cyan: '00FFFF', magenta: 'FF00FF', silver: 'C0C0C0', maroon: '800000',
  navy: '000080', lime: '00FF00', olive: '808000', teal: '008080', aqua: '00FFFF', fuchsia: 'FF00FF', dkGray: 'A9A9A9', ltGray: 'D3D3D3',
};

const HEX6 = /^[0-9a-fA-F]{6}$/;

/** `#RRGGBB` from a 6-digit hex string, or `undefined`. */
export const toHex = (v: string | undefined): string | undefined => (v !== undefined && HEX6.test(v.trim()) ? `#${v.trim().toUpperCase()}` : undefined);

/** The first child that is a colour element (`srgbClr`, `schemeClr`, `sysClr`, `prstClr`, `scrgbClr`, `hslClr`), or `undefined`. */
export function colorElement(parent: XmlElement | undefined): XmlElement | undefined {
  return parent?.children.find((c) => COLOR_NAMES.has(c.name));
}
const COLOR_NAMES = new Set(['srgbClr', 'schemeClr', 'sysClr', 'prstClr', 'scrgbClr', 'hslClr']);

/**
 * A colour element as a `PptxColor`. A `schemeClr` goes through the colour map (`bg1` -> `lt1`); `phClr` (the placeholder
 * colour of a style reference) and anything unreadable give `undefined`, so the caller falls back to what it inherits.
 * `lumMod`, `lumOff` and `alpha` are carried; other transforms (`tint`, `shade`, `satMod`...) are not interpreted.
 */
export function parseColor(el: XmlElement | undefined, colorMap: PptxColorMap = DEFAULT_COLOR_MAP): PptxColor | undefined {
  if (!el) return undefined;
  const alpha = pct(kid(el, 'alpha')?.attrs.val);
  const withAlpha = <T extends PptxColor>(c: T): T => (alpha !== undefined && alpha < 1 ? { ...c, alpha: Math.max(0, alpha) } : c);
  switch (el.name) {
    case 'srgbClr': {
      const hex = toHex(el.attrs.val);
      return hex ? withAlpha({ kind: 'rgb', hex }) : undefined;
    }
    case 'sysClr': {
      const hex = toHex(el.attrs.lastClr) ?? (el.attrs.val === 'window' ? '#FFFFFF' : el.attrs.val === 'windowText' ? '#000000' : undefined);
      return hex ? withAlpha({ kind: 'rgb', hex }) : undefined;
    }
    case 'prstClr': {
      const hex = PRESET_COLORS[el.attrs.val ?? ''];
      return withAlpha({ kind: 'rgb', hex: `#${hex ?? '000000'}` });
    }
    case 'schemeClr': {
      const role = el.attrs.val ?? '';
      const slot = SLOT_SET.has(role) ? (role as PptxThemeSlot) : (colorMap as Record<string, PptxThemeSlot | undefined>)[role];
      if (!slot) return undefined;
      const lumMod = pct(kid(el, 'lumMod')?.attrs.val);
      const lumOff = pct(kid(el, 'lumOff')?.attrs.val);
      return withAlpha({ kind: 'theme', slot, ...(lumMod !== undefined ? { lumMod } : {}), ...(lumOff !== undefined ? { lumOff } : {}) });
    }
    default:
      return undefined;
  }
}
