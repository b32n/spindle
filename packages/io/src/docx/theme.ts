import { attrsOf, childrenOf, find, parseXml, tagOf, type XmlNode } from './xml-tree';

// ---------------------------------------------------------------------------
// Theme (word/theme/theme1.xml): the major/minor Latin fonts and the colour scheme that `asciiTheme`,
// `themeColor`, `themeFill` and their tint/shade modifiers point at.
// ---------------------------------------------------------------------------

export interface Theme {
  majorFont?: string;
  minorFont?: string;
  /** Scheme slot (`dk1`, `lt1`, `accent1`, `hlink`, ...) -> `#rrggbb`. */
  colors: Map<string, string>;
}

/** w:themeColor / w:themeFill value -> scheme slot (the default colour-scheme mapping: bg1 = lt1, tx1 = dk1, ...). */
const THEME_COLOR_SLOT: Record<string, string> = {
  dark1: 'dk1', text1: 'dk1',
  light1: 'lt1', background1: 'lt1',
  dark2: 'dk2', text2: 'dk2',
  light2: 'lt2', background2: 'lt2',
  accent1: 'accent1', accent2: 'accent2', accent3: 'accent3', accent4: 'accent4', accent5: 'accent5', accent6: 'accent6',
  hyperlink: 'hlink', followedHyperlink: 'folHlink',
};

export function parseTheme(xml: string | undefined): Theme | undefined {
  if (!xml) return undefined;
  let root: XmlNode[];
  try {
    root = parseXml(xml);
  } catch {
    return undefined;
  }
  const theme = find(root, 'a:theme');
  const elements = theme ? find(childrenOf(theme), 'a:themeElements') : undefined;
  if (!elements) return undefined;
  const out: Theme = { colors: new Map() };

  const scheme = find(childrenOf(elements), 'a:clrScheme');
  for (const slot of scheme ? childrenOf(scheme) : []) {
    const name = Object.keys(slot).find((k) => k.startsWith('a:'))?.slice(2);
    if (!name) continue;
    const color = childrenOf(slot).find((c) => tagOf(c) !== '#text');
    if (!color) continue;
    const a = attrsOf(color);
    const hex = a['@_val'] && /^[0-9a-fA-F]{6}$/.test(a['@_val']) ? a['@_val'] : a['@_lastClr'];
    if (hex && /^[0-9a-fA-F]{6}$/.test(hex)) out.colors.set(name, `#${hex.toLowerCase()}`);
  }

  const fonts = find(childrenOf(elements), 'a:fontScheme');
  const latin = (which: 'a:majorFont' | 'a:minorFont') => {
    const f = fonts ? find(childrenOf(fonts), which) : undefined;
    const l = f ? find(childrenOf(f), 'a:latin') : undefined;
    return l ? attrsOf(l)['@_typeface'] || undefined : undefined;
  };
  out.majorFont = latin('a:majorFont');
  out.minorFont = latin('a:minorFont');
  return out;
}

/** `majorHAnsi`, `minorAscii`, ... (a w:rFonts theme attribute) -> the theme's Latin font. */
export function themeFont(theme: Theme | undefined, value: string | undefined): string | undefined {
  if (!theme || !value) return undefined;
  if (value.startsWith('major')) return theme.majorFont;
  if (value.startsWith('minor')) return theme.minorFont;
  return undefined;
}

function rgb(hex: string): [number, number, number] {
  return [parseInt(hex.slice(1, 3), 16), parseInt(hex.slice(3, 5), 16), parseInt(hex.slice(5, 7), 16)];
}

function toHsl([r, g, b]: [number, number, number]): [number, number, number] {
  const R = r / 255, G = g / 255, B = b / 255;
  const max = Math.max(R, G, B), min = Math.min(R, G, B);
  const l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  const h = max === R ? (G - B) / d + (G < B ? 6 : 0) : max === G ? (B - R) / d + 2 : (R - G) / d + 4;
  return [h / 6, s, l];
}

function fromHsl([h, s, l]: [number, number, number]): string {
  const hue = (p: number, q: number, t: number) => {
    const u = t < 0 ? t + 1 : t > 1 ? t - 1 : t;
    if (u < 1 / 6) return p + (q - p) * 6 * u;
    if (u < 1 / 2) return q;
    if (u < 2 / 3) return p + (q - p) * (2 / 3 - u) * 6;
    return p;
  };
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  const channels = s === 0 ? [l, l, l] : [hue(p, q, h + 1 / 3), hue(p, q, h), hue(p, q, h - 1 / 3)];
  return '#' + channels.map((c) => Math.max(0, Math.min(255, Math.floor(c * 255 + 1e-6))).toString(16).padStart(2, '0')).join('');
}

/**
 * A theme colour with Word's tint (towards white) and shade (towards black) applied to its luminance, the way
 * `w:themeTint="BF"` / `w:themeShade="BF"` are specified (a value of 0-255, 255 meaning unchanged).
 * Returns undefined when the theme has no such colour.
 */
export function themeColor(theme: Theme | undefined, name: string | undefined, tintHex?: string, shadeHex?: string): string | undefined {
  const slot = name ? THEME_COLOR_SLOT[name] : undefined;
  const base = slot ? theme?.colors.get(slot) : undefined;
  if (!base) return undefined;
  const tint = tintHex ? parseInt(tintHex, 16) : NaN;
  const shade = shadeHex ? parseInt(shadeHex, 16) : NaN;
  if (Number.isNaN(tint) && Number.isNaN(shade)) return base;
  const [h, s, l] = toHsl(rgb(base));
  let lum = l;
  if (!Number.isNaN(shade)) lum = lum * (shade / 255);
  if (!Number.isNaN(tint)) lum = lum * (tint / 255) + (1 - tint / 255);
  return fromHsl([h, s, lum]);
}

/**
 * Word writes the colour it computed from a theme colour next to `w:themeColor` (as `w:val` / `w:fill`). Our
 * tint/shade arithmetic can be one step off in a channel, so when that recorded value is within a few steps of
 * the computed one it is the more exact answer; a recorded value that is far off (a producer that left the
 * unmodified base colour there) loses to the theme.
 */
export function preferRecorded(computed: string, recorded: string | undefined): string {
  if (!recorded || !/^[0-9a-fA-F]{6}$/.test(recorded)) return computed;
  const a = rgb(computed);
  const b = rgb(`#${recorded}`);
  return a.every((c, i) => Math.abs(c - b[i]) <= 3) ? `#${recorded.toLowerCase()}` : computed;
}
