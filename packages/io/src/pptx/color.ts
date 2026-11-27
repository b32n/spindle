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

/** The `ST_PresetColorVal` names of `a:prstClr` as the CSS colour names in lower case (`dkBlue`, `ltGray`, `medBlue` are `darkblue`, `lightgray`, `mediumblue`). */
const PRESET_COLORS: Record<string, string> = {
  aliceblue: 'F0F8FF', antiquewhite: 'FAEBD7', aqua: '00FFFF', aquamarine: '7FFFD4', azure: 'F0FFFF', beige: 'F5F5DC', bisque: 'FFE4C4',
  black: '000000', blanchedalmond: 'FFEBCD', blue: '0000FF', blueviolet: '8A2BE2', brown: 'A52A2A', burlywood: 'DEB887', cadetblue: '5F9EA0',
  chartreuse: '7FFF00', chocolate: 'D2691E', coral: 'FF7F50', cornflowerblue: '6495ED', cornsilk: 'FFF8DC', crimson: 'DC143C', cyan: '00FFFF',
  darkblue: '00008B', darkcyan: '008B8B', darkgoldenrod: 'B8860B', darkgray: 'A9A9A9', darkgreen: '006400', darkkhaki: 'BDB76B',
  darkmagenta: '8B008B', darkolivegreen: '556B2F', darkorange: 'FF8C00', darkorchid: '9932CC', darkred: '8B0000', darksalmon: 'E9967A',
  darkseagreen: '8FBC8B', darkslateblue: '483D8B', darkslategray: '2F4F4F', darkturquoise: '00CED1', darkviolet: '9400D3', deeppink: 'FF1493',
  deepskyblue: '00BFFF', dimgray: '696969', dodgerblue: '1E90FF', firebrick: 'B22222', floralwhite: 'FFFAF0', forestgreen: '228B22',
  fuchsia: 'FF00FF', gainsboro: 'DCDCDC', ghostwhite: 'F8F8FF', gold: 'FFD700', goldenrod: 'DAA520', gray: '808080', green: '008000',
  greenyellow: 'ADFF2F', honeydew: 'F0FFF0', hotpink: 'FF69B4', indianred: 'CD5C5C', indigo: '4B0082', ivory: 'FFFFF0', khaki: 'F0E68C',
  lavender: 'E6E6FA', lavenderblush: 'FFF0F5', lawngreen: '7CFC00', lemonchiffon: 'FFFACD', lightblue: 'ADD8E6', lightcoral: 'F08080',
  lightcyan: 'E0FFFF', lightgoldenrodyellow: 'FAFAD2', lightgray: 'D3D3D3', lightgreen: '90EE90', lightpink: 'FFB6C1', lightsalmon: 'FFA07A',
  lightseagreen: '20B2AA', lightskyblue: '87CEFA', lightslategray: '778899', lightsteelblue: 'B0C4DE', lightyellow: 'FFFFE0', lime: '00FF00',
  limegreen: '32CD32', linen: 'FAF0E6', magenta: 'FF00FF', maroon: '800000', mediumaquamarine: '66CDAA', mediumblue: '0000CD',
  mediumorchid: 'BA55D3', mediumpurple: '9370DB', mediumseagreen: '3CB371', mediumslateblue: '7B68EE', mediumspringgreen: '00FA9A',
  mediumturquoise: '48D1CC', mediumvioletred: 'C71585', midnightblue: '191970', mintcream: 'F5FFFA', mistyrose: 'FFE4E1', moccasin: 'FFE4B5',
  navajowhite: 'FFDEAD', navy: '000080', oldlace: 'FDF5E6', olive: '808000', olivedrab: '6B8E23', orange: 'FFA500', orangered: 'FF4500',
  orchid: 'DA70D6', palegoldenrod: 'EEE8AA', palegreen: '98FB98', paleturquoise: 'AFEEEE', palevioletred: 'DB7093', papayawhip: 'FFEFD5',
  peachpuff: 'FFDAB9', peru: 'CD853F', pink: 'FFC0CB', plum: 'DDA0DD', powderblue: 'B0E0E6', purple: '800080', red: 'FF0000', rosybrown: 'BC8F8F',
  royalblue: '4169E1', saddlebrown: '8B4513', salmon: 'FA8072', sandybrown: 'F4A460', seagreen: '2E8B57', seashell: 'FFF5EE', sienna: 'A0522D',
  silver: 'C0C0C0', skyblue: '87CEEB', slateblue: '6A5ACD', slategray: '708090', snow: 'FFFAFA', springgreen: '00FF7F', steelblue: '4682B4',
  tan: 'D2B48C', teal: '008080', thistle: 'D8BFD8', tomato: 'FF6347', turquoise: '40E0D0', violet: 'EE82EE', wheat: 'F5DEB3', white: 'FFFFFF',
  whitesmoke: 'F5F5F5', yellow: 'FFFF00', yellowgreen: '9ACD32',
};

/** The hex of a preset colour name (case-insensitive, `dk`/`lt`/`med` and `grey` spellings included), or `undefined` for a name that is not in the schema. */
export function presetHex(name: string | undefined): string | undefined {
  const key = (name ?? '').toLowerCase().replace(/^dk/, 'dark').replace(/^lt/, 'light').replace(/^med(?!ium)/, 'medium').replace('grey', 'gray');
  return Object.prototype.hasOwnProperty.call(PRESET_COLORS, key) ? PRESET_COLORS[key] : undefined;
}

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
      const hex = presetHex(el.attrs.val);
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

/** A layout's `p:clrMapOvr/a:overrideClrMapping`: the roles it names a valid slot for replace the master's, the others stay. */
export function overrideColorMap(base: PptxColorMap, override: XmlElement | undefined): PptxColorMap {
  const map = { ...base };
  for (const role of Object.keys(map) as (keyof PptxColorMap)[]) {
    const slot = override?.attrs[role];
    if (slot !== undefined && SLOT_SET.has(slot)) map[role] = slot as PptxThemeSlot;
  }
  return map;
}
