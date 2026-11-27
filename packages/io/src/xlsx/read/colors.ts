import type { XlsxColor } from '@b32nio/spindle-shared/export/xlsx';
import type { XmlElement } from '../xml';
import { numAttr } from './values';

// ---------------------------------------------------------------------------
// Colours: `<color rgb|theme|indexed|auto tint>` to `#rrggbb`.
//
// - rgb: `AARRGGBB` (alpha ignored: Excel writes `FF`, and a cell colour has no transparency) or a bare `RRGGBB`.
// - theme: index into the theme's colour scheme in Excel's order (lt1, dk1, lt2, dk2, accent1..6, hlink, folHlink: the first
//   two pairs are swapped against the order the theme part lists them in).
// - indexed: the workbook's own palette if it has one (an index it does not reach uses the default palette), else the default palette; 64 and 65 are the system foreground
//   (black) and background (white).
// - tint: lightness adjustment in HLS space (ECMA-376 18.3.1.15): a negative tint darkens (L * (1 + tint)), a positive one
//   lightens (L * (1 - tint) + tint). Excel does this in integer HLS arithmetic; this is floating point and lands within 2 of
//   each channel it shows (the colour-picker values are the oracle in colors.test.ts).
// ---------------------------------------------------------------------------

/** The palette Excel uses when the workbook has no `indexedColors` (ECMA-376 18.8.27). */
export const DEFAULT_INDEXED_COLORS: readonly string[] = [
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080',
  '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
  '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF',
  '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
  '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696',
  '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333',
].map((c) => `#${c.toLowerCase()}`);

/** The Office theme colours (2013 and later), used when a workbook has no theme part or a theme without a colour scheme. */
export const DEFAULT_THEME_COLORS: readonly string[] = [
  '#ffffff', '#000000', '#e7e6e6', '#44546a', '#4472c4', '#ed7d31', '#a5a5a5', '#ffc000', '#5b9bd5', '#70ad47', '#0563c1', '#954f72',
];

const SYSTEM_FOREGROUND = 64;
const SYSTEM_BACKGROUND = 65;

export interface ColorContext {
  /** 12 `#rrggbb` entries in a cell colour's `theme` index order. */
  themeColors: readonly string[];
  /** The workbook's own palette, when it has one; a palette shorter than 64 falls back to the default per index. */
  indexedColors?: readonly string[];
}

const HEX6 = /^[0-9a-fA-F]{6}$/;
const HEX8 = /^[0-9a-fA-F]{8}$/;

/** `#rrggbb` (lower case) from `RRGGBB` or `AARRGGBB`, or `undefined` for anything else. */
export function parseRgb(value: string | undefined): string | undefined {
  if (value === undefined) return undefined;
  if (HEX6.test(value)) return `#${value.toLowerCase()}`;
  if (HEX8.test(value)) return `#${value.slice(2).toLowerCase()}`;
  return undefined;
}

function rgbToHls(hex: string): [number, number, number] {
  const r = parseInt(hex.slice(1, 3), 16) / 255;
  const g = parseInt(hex.slice(3, 5), 16) / 255;
  const b = parseInt(hex.slice(5, 7), 16) / 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return [0, l, 0];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h: number;
  if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
  else if (max === g) h = (b - r) / d + 2;
  else h = (r - g) / d + 4;
  return [h / 6, l, s];
}

export function hlsToRgb(h: number, l: number, s: number): string {
  const channel = (p: number, q: number, t0: number): number => {
    const t = t0 < 0 ? t0 + 1 : t0 > 1 ? t0 - 1 : t0;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  let r = l;
  let g = l;
  let b = l;
  if (s !== 0) {
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
    const p = 2 * l - q;
    r = channel(p, q, h + 1 / 3);
    g = channel(p, q, h);
    b = channel(p, q, h - 1 / 3);
  }
  const byte = (x: number) => Math.round(x * 255).toString(16).padStart(2, '0');
  return `#${byte(r)}${byte(g)}${byte(b)}`;
}

/** Apply a tint (-1 to 1, clamped) to `#rrggbb`. A tint of 0 returns the colour untouched. */
export function applyTint(hex: string, tint: number): string {
  if (tint === 0) return hex;
  const t = Math.max(-1, Math.min(1, tint));
  const [h, l, s] = rgbToHls(hex);
  return hlsToRgb(h, t < 0 ? l * (1 + t) : l * (1 - t) + t, s);
}

/** What a colour resolves to, or `undefined` (`auto`, an unreadable `rgb`, an index outside the palette). */
export function resolveColor(color: Pick<XlsxColor, 'rgb' | 'theme' | 'indexed' | 'auto' | 'tint'>, ctx: ColorContext): string | undefined {
  let base: string | undefined;
  if (color.rgb !== undefined) base = parseRgb(color.rgb);
  else if (color.theme !== undefined) base = ctx.themeColors[color.theme];
  else if (color.indexed !== undefined) {
    if (color.indexed === SYSTEM_FOREGROUND) base = '#000000';
    else if (color.indexed === SYSTEM_BACKGROUND) base = '#ffffff';
    else base = ctx.indexedColors?.[color.indexed] ?? DEFAULT_INDEXED_COLORS[color.indexed];
  }
  return base === undefined ? undefined : applyTint(base, color.tint ?? 0);
}

/** Read a `<color>`-shaped element (`color`, `fgColor`, `bgColor`) with its resolved value. `undefined` when there is no element. */
export function readColor(el: XmlElement | undefined, ctx: ColorContext): XlsxColor | undefined {
  if (!el) return undefined;
  const out: XlsxColor = {};
  const a = el.attrs;
  if (a.rgb !== undefined) out.rgb = a.rgb;
  const theme = numAttr(a.theme);
  if (theme !== undefined) out.theme = theme;
  const indexed = numAttr(a.indexed);
  if (indexed !== undefined) out.indexed = indexed;
  if (a.auto === '1' || a.auto === 'true') out.auto = true;
  const tint = numAttr(a.tint);
  if (tint !== undefined) out.tint = tint;
  const resolved = resolveColor(out, ctx);
  if (resolved !== undefined) out.resolved = resolved;
  return out;
}
