// Theme resolution — turn symbolic Colors/fonts into concrete CSS values at
// render time. Renderers call these on every paint, so a theme switch is a
// single re-render with no data migration.

import type { Color } from '../scene/types';
import type { ThemeData, PlaceholderStyle } from './types';

/** Parse '#RGB' or '#RRGGBB' into [r, g, b] (0–255). */
function hexToRgb(hex: string): [number, number, number] {
  let h = hex.replace('#', '').trim();
  if (h.length === 3) {
    h = h[0] + h[0] + h[1] + h[1] + h[2] + h[2];
  }
  const n = parseInt(h, 16);
  if (h.length !== 6 || Number.isNaN(n)) return [0, 0, 0];
  return [(n >> 16) & 0xff, (n >> 8) & 0xff, n & 0xff];
}

function clamp255(v: number): number {
  return Math.max(0, Math.min(255, Math.round(v)));
}

/** Resolve a Color to its red, green, blue (0–255) and alpha (0–1) channels. */
export function colorChannels(color: Color, theme: ThemeData): [number, number, number, number] {
  let rgb: [number, number, number];
  if (color.kind === 'rgb') {
    rgb = hexToRgb(color.hex);
  } else {
    rgb = hexToRgb(theme.colors[color.slot] ?? '#000000');
    if (color.lumMod !== undefined || color.lumOff !== undefined) {
      const mod = color.lumMod ?? 1;
      const off = (color.lumOff ?? 0) * 255;
      rgb = [
        clamp255(rgb[0] * mod + off),
        clamp255(rgb[1] * mod + off),
        clamp255(rgb[2] * mod + off),
      ];
    }
  }
  const alpha = color.alpha === undefined ? 1 : Math.max(0, Math.min(1, color.alpha));
  return [rgb[0], rgb[1], rgb[2], alpha];
}

/**
 * Resolve a Color to a CSS color string. Theme slots are looked up in the
 * theme; PPTX luminance transforms (lumMod scales, lumOff shifts toward white)
 * are applied per channel; alpha (0–1) produces an rgba() string.
 */
export function resolveColor(color: Color, theme: ThemeData): string {
  const [r, g, b, a] = colorChannels(color, theme);
  if (color.alpha !== undefined && color.alpha < 1) return `rgba(${r}, ${g}, ${b}, ${a})`;
  return `rgb(${r}, ${g}, ${b})`;
}

/**
 * Resolve a font reference to a CSS font-family. 'major'/'minor' map to the
 * theme's font pair; anything else is treated as a literal family.
 */
export function resolveFont(font: string | undefined, theme: ThemeData): string {
  if (font === 'major') return theme.fonts.major;
  if (font === 'minor' || font === undefined) return theme.fonts.minor;
  return font;
}

/**
 * Resolve a placeholder's default style into concrete CSS values. Used when
 * rendering placeholder prompt text and (Phase 4) when materializing a
 * placeholder into a real element.
 */
export function resolvePlaceholderStyle(
  style: PlaceholderStyle | undefined,
  theme: ThemeData
): { color: string; fontFamily: string; fontSize: number; bold: boolean } {
  return {
    color: style?.color ? resolveColor(style.color, theme) : resolveColor({ kind: 'theme', slot: 'dk1' }, theme),
    fontFamily: resolveFont(style?.font, theme),
    fontSize: style?.fontSize ?? 18,
    bold: style?.bold ?? false,
  };
}
