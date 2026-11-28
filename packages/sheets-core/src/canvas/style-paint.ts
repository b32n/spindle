// What the canvas draws for the style fields a file can bring: border lines, pattern fills and text decorations. Pure functions (no
// canvas), so what is drawn for a style is unit-tested without one.

import type { BorderKind, CellStyle } from '../types';
import { resolveBorderKind } from '../border-kinds';

/** How one border side is stroked: line width in pixels, dash pattern (empty: solid) and whether it is two thin lines. */
export interface BorderPaint {
  width: number;
  dash: readonly number[];
  double: boolean;
}

/** The stroke of each kind. Each of the thirteen is distinct; Excel's `hair` (finer dots) and `slantDashDot` (slanted dashes) are approximated. */
export const BORDER_PAINT: Readonly<Record<BorderKind, BorderPaint>> = {
  thin: { width: 1, dash: [], double: false },
  hair: { width: 1, dash: [1, 2], double: false },
  dotted: { width: 1, dash: [1, 1], double: false },
  dashed: { width: 1, dash: [4, 2], double: false },
  dashDot: { width: 1, dash: [4, 2, 1, 2], double: false },
  dashDotDot: { width: 1, dash: [4, 2, 1, 2, 1, 2], double: false },
  double: { width: 3, dash: [], double: true },
  medium: { width: 2, dash: [], double: false },
  mediumDashed: { width: 2, dash: [5, 3], double: false },
  mediumDashDot: { width: 2, dash: [5, 3, 1, 3], double: false },
  mediumDashDotDot: { width: 2, dash: [5, 3, 1, 3, 1, 3], double: false },
  slantDashDot: { width: 2, dash: [6, 2, 2, 2], double: false },
  thick: { width: 3, dash: [], double: false },
};

/** The kind drawn for a border shorthand (`2px dashed #000`) and the kind stored next to it (ignored when the shorthand was edited since). */
export function borderKindOf(border: string | undefined, kind: BorderKind | undefined): BorderKind | undefined {
  if (!border || border === 'none') return undefined;
  const parts = border.split(' ');
  const width = parseInt(parts[0], 10);
  if (parts.length < 2 || Number.isNaN(width)) return undefined;
  return resolveBorderKind(width, parts[1], kind);
}

/** Share of the pattern colour in the flat colour a pattern fill is drawn as (the pattern's average ink coverage). */
const PATTERN_COVERAGE: Readonly<Record<string, number>> = {
  gray0625: 0.0625, gray125: 0.125, lightGray: 0.25, mediumGray: 0.5, darkGray: 0.75,
  lightHorizontal: 0.25, lightVertical: 0.25, lightDown: 0.25, lightUp: 0.25, lightGrid: 0.35, lightTrellis: 0.35,
  darkHorizontal: 0.5, darkVertical: 0.5, darkDown: 0.5, darkUp: 0.5, darkGrid: 0.65, darkTrellis: 0.65,
};

function rgb(color: string): [number, number, number] | undefined {
  const m = /^#([0-9a-f]{3}|[0-9a-f]{6})$/i.exec(color);
  if (!m) return undefined;
  const hex = m[1].length === 3 ? [...m[1]].map((c) => c + c).join('') : m[1];
  return [parseInt(hex.slice(0, 2), 16), parseInt(hex.slice(2, 4), 16), parseInt(hex.slice(4, 6), 16)];
}

/**
 * The colour a cell's fill is drawn as, or `undefined` for none: a solid fill's colour, or for a pattern fill the pattern colour mixed into
 * the colour behind it (white when the file gives none) by the pattern's coverage. A colour that is not `#rgb` or `#rrggbb` is used as is.
 */
export function fillColorOf(style: CellStyle | undefined, behind = '#ffffff'): string | undefined {
  if (!style) return undefined;
  if (style.fillPattern === undefined || style.patternColor === undefined) return style.backgroundColor;
  const ink = rgb(style.patternColor);
  const paper = rgb(style.backgroundColor ?? behind);
  if (!ink || !paper) return style.patternColor;
  const t = PATTERN_COVERAGE[style.fillPattern] ?? 0.5;
  const mixed = ink.map((c, i) => Math.round(c * t + paper[i] * (1 - t)));
  return `#${mixed.map((c) => c.toString(16).padStart(2, '0')).join('')}`;
}

/** Pixels a level of indent moves the text (a level is three characters of the default font). */
export const INDENT_PX = 9;
