// The thirteen border line styles of the file format, and what the editor holds and draws for each (see `BorderKind` in types.ts).

import type { BorderKind } from './types';

/** The CSS shorthand parts the editor stores for each kind: its width in pixels and its CSS line style. */
export const CSS_OF_KIND: Readonly<Record<BorderKind, { width: number; style: 'solid' | 'dashed' | 'dotted' | 'double' }>> = {
  thin: { width: 1, style: 'solid' },
  hair: { width: 1, style: 'dotted' },
  dotted: { width: 1, style: 'dotted' },
  dashed: { width: 1, style: 'dashed' },
  dashDot: { width: 1, style: 'dashed' },
  dashDotDot: { width: 1, style: 'dashed' },
  double: { width: 3, style: 'double' },
  medium: { width: 2, style: 'solid' },
  mediumDashed: { width: 2, style: 'dashed' },
  mediumDashDot: { width: 2, style: 'dashed' },
  mediumDashDotDot: { width: 2, style: 'dashed' },
  slantDashDot: { width: 2, style: 'dashed' },
  thick: { width: 3, style: 'solid' },
};

export const isBorderKind = (value: unknown): value is BorderKind => typeof value === 'string' && Object.prototype.hasOwnProperty.call(CSS_OF_KIND, value);

/** The kind the shorthand alone says: what the editor's own border tools make. */
export function inferBorderKind(widthPx: number, cssStyle: string): BorderKind {
  if (cssStyle === 'dashed') return 'dashed';
  if (cssStyle === 'dotted') return 'dotted';
  if (cssStyle === 'double') return 'double';
  if (widthPx >= 3) return 'thick';
  if (widthPx === 2) return 'medium';
  return 'thin';
}

/** The kept kind when the shorthand still says what it did when the kind was stored (it may have been edited since), otherwise what the shorthand says. */
export function resolveBorderKind(widthPx: number, cssStyle: string, kind: BorderKind | undefined): BorderKind {
  return isBorderKind(kind) && CSS_OF_KIND[kind].width === widthPx && CSS_OF_KIND[kind].style === cssStyle ? kind : inferBorderKind(widthPx, cssStyle);
}

/** The kind to keep next to the shorthand made from `kind`: only one the shorthand alone would not give back. */
export function keptBorderKind(kind: BorderKind): BorderKind | undefined {
  const css = CSS_OF_KIND[kind];
  return inferBorderKind(css.width, css.style) === kind ? undefined : kind;
}

/** Kinds the canvas draws only approximately: Excel's `hair` has finer dots than a one-pixel dotted line, `slantDashDot` slants its dashes. */
export const APPROXIMATED_BORDER_KINDS: ReadonlySet<string> = new Set(['hair', 'slantDashDot']);
