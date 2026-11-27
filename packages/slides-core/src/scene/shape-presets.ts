// Shape preset metadata (K5): the picker categories, each preset's adjustment guides (defaults and limits), and the
// `unknown:<prst>` passthrough helpers. The geometry itself is in slides-react (`shapeGeom`); the guide semantics follow
// the OOXML preset definitions (an `adj` of 16667 is 0.16667 here), hand-built rather than interpreted (plan SD7).

import type { ShapeKind, ShapePreset, UnknownShape } from './types';

export type ShapeCategory = 'rectangles' | 'basic' | 'arrows' | 'stars' | 'callouts' | 'flowchart' | 'brackets' | 'math';

export const SHAPE_CATEGORIES: ReadonlyArray<{ id: ShapeCategory; label: string }> = [
  { id: 'rectangles', label: 'Rectangles' },
  { id: 'basic', label: 'Basic shapes' },
  { id: 'arrows', label: 'Block arrows' },
  { id: 'stars', label: 'Stars' },
  { id: 'callouts', label: 'Callouts' },
  { id: 'flowchart', label: 'Flowchart' },
  { id: 'brackets', label: 'Brackets and braces' },
  { id: 'math', label: 'Math and lines' },
];

/** One adjustment guide: its default and the range it is clamped to. */
export interface AdjustmentSpec {
  def: number;
  min: number;
  max: number;
}

export interface ShapePresetInfo {
  label: string;
  category: ShapeCategory;
  /** Guides in OOXML order; absent when the preset has none. */
  adj?: readonly AdjustmentSpec[];
}

const a = (def: number, min = 0, max = 0.5): AdjustmentSpec => ({ def, min, max });
/** A callout tail tip, relative to the shape centre in box fractions. */
const tip = (def: number): AdjustmentSpec => a(def, -1, 2);

const info = (label: string, category: ShapeCategory, ...adj: AdjustmentSpec[]): ShapePresetInfo =>
  adj.length ? { label, category, adj } : { label, category };

export const SHAPE_PRESET_INFO: Readonly<Record<ShapePreset, ShapePresetInfo>> = {
  rect: info('Rectangle', 'rectangles'),
  roundRect: info('Rounded rectangle', 'rectangles', a(0.15)),
  snip1Rect: info('Snip single corner rectangle', 'rectangles', a(1 / 6)),
  snip2SameRect: info('Snip same-side corner rectangle', 'rectangles', a(1 / 6), a(0)),
  snip2DiagRect: info('Snip diagonal corner rectangle', 'rectangles', a(0), a(1 / 6)),
  snipRoundRect: info('Snip and round single corner rectangle', 'rectangles', a(1 / 6), a(1 / 6)),
  round1Rect: info('Round single corner rectangle', 'rectangles', a(1 / 6)),
  round2SameRect: info('Round same-side corner rectangle', 'rectangles', a(1 / 6), a(0)),
  round2DiagRect: info('Round diagonal corner rectangle', 'rectangles', a(1 / 6), a(0)),
  foldedCorner: info('Folded corner', 'rectangles', a(1 / 6)),
  frame: info('Frame', 'rectangles', a(0.125)),
  plaque: info('Plaque', 'rectangles', a(1 / 6)),
  bevel: info('Bevel', 'rectangles', a(0.125)),
  can: info('Can', 'rectangles', a(0.25, 0, 2)),
  cube: info('Cube', 'rectangles', a(0.25, 0, 1)),

  ellipse: info('Oval', 'basic'),
  triangle: info('Isosceles triangle', 'basic'),
  rightTriangle: info('Right triangle', 'basic'),
  diamond: info('Diamond', 'basic'),
  pentagon: info('Regular pentagon', 'basic'),
  hexagon: info('Hexagon', 'basic'),
  heptagon: info('Heptagon', 'basic'),
  octagon: info('Octagon', 'basic'),
  decagon: info('Decagon', 'basic'),
  dodecagon: info('Dodecagon', 'basic'),
  parallelogram: info('Parallelogram', 'basic'),
  trapezoid: info('Trapezoid', 'basic'),
  plus: info('Cross', 'basic'),
  heart: info('Heart', 'basic'),
  cloud: info('Cloud', 'basic'),
  donut: info('Donut', 'basic', a(0.25)),
  noSmoking: info('No symbol', 'basic', a(0.1875)),
  teardrop: info('Teardrop', 'basic', a(1, 0, 2)),
  pie: info('Pie', 'basic', a(0, 0, 1), a(0.75, 0, 1)),
  chord: info('Chord', 'basic', a(0.125, 0, 1), a(0.75, 0, 1)),
  blockArc: info('Block arc', 'basic', a(0.5, 0, 1), a(0, 0, 1), a(0.25)),
  moon: info('Moon', 'basic', a(0.5, 0, 0.875)),
  smileyFace: info('Smiley face', 'basic'),
  lightningBolt: info('Lightning bolt', 'basic'),

  arrowRight: info('Arrow right', 'arrows'),
  arrowLeft: info('Arrow left', 'arrows'),
  rightArrow: info('Right arrow', 'arrows', a(0.5, 0, 1), a(0.5, 0, 1)),
  leftArrow: info('Left arrow', 'arrows', a(0.5, 0, 1), a(0.5, 0, 1)),
  upArrow: info('Up arrow', 'arrows', a(0.5, 0, 1), a(0.5, 0, 1)),
  downArrow: info('Down arrow', 'arrows', a(0.5, 0, 1), a(0.5, 0, 1)),
  leftRightArrow: info('Left-right arrow', 'arrows', a(0.5, 0, 1), a(0.5, 0, 1)),
  upDownArrow: info('Up-down arrow', 'arrows', a(0.5, 0, 1), a(0.5, 0, 1)),
  homePlate: info('Pentagon arrow', 'arrows', a(0.5, 0, 1)),
  chevron: info('Chevron', 'arrows'),

  star4: info('4-point star', 'stars', a(0.125)),
  star5: info('5-point star', 'stars'),
  star6: info('6-point star', 'stars', a(0.28868)),
  star8: info('8-point star', 'stars', a(0.375)),
  star12: info('12-point star', 'stars', a(0.375)),

  wedgeRectCallout: info('Rectangular callout', 'callouts', tip(-0.20833), tip(0.625)),
  wedgeRoundRectCallout: info('Rounded rectangular callout', 'callouts', tip(-0.20833), tip(0.625), a(1 / 6)),
  wedgeEllipseCallout: info('Oval callout', 'callouts', tip(-0.20833), tip(0.625)),

  flowChartProcess: info('Flowchart: process', 'flowchart'),
  flowChartAlternateProcess: info('Flowchart: alternate process', 'flowchart'),
  flowChartDecision: info('Flowchart: decision', 'flowchart'),
  flowChartInputOutput: info('Flowchart: data', 'flowchart'),
  flowChartTerminator: info('Flowchart: terminator', 'flowchart'),
  flowChartConnector: info('Flowchart: connector', 'flowchart'),
  flowChartPredefinedProcess: info('Flowchart: predefined process', 'flowchart'),
  flowChartDelay: info('Flowchart: delay', 'flowchart'),
  flowChartPreparation: info('Flowchart: preparation', 'flowchart'),
  flowChartOffpageConnector: info('Flowchart: off-page connector', 'flowchart'),
  flowChartMerge: info('Flowchart: merge', 'flowchart'),

  leftBracket: info('Left bracket', 'brackets', a(1 / 12)),
  rightBracket: info('Right bracket', 'brackets', a(1 / 12)),
  leftBrace: info('Left brace', 'brackets', a(1 / 12), a(0.5, 0, 1)),
  rightBrace: info('Right brace', 'brackets', a(1 / 12), a(0.5, 0, 1)),
  bracketPair: info('Double bracket', 'brackets', a(1 / 6)),

  mathPlus: info('Plus', 'math', a(0.2352, 0, 1)),
  mathMinus: info('Minus', 'math', a(0.2352, 0, 1)),
  lineInv: info('Inverted line', 'math'),
};

/** Every built-in preset, grouped by category in picker order. */
export const SHAPE_PRESETS: readonly ShapePreset[] = SHAPE_CATEGORIES.flatMap((c) =>
  (Object.keys(SHAPE_PRESET_INFO) as ShapePreset[]).filter((p) => SHAPE_PRESET_INFO[p].category === c.id),
);

export function isShapePreset(shape: string): shape is ShapePreset {
  return Object.prototype.hasOwnProperty.call(SHAPE_PRESET_INFO, shape);
}

const UNKNOWN_PREFIX = 'unknown:';

export function unknownShape(prst: string): UnknownShape {
  return `${UNKNOWN_PREFIX}${prst}`;
}

export function isUnknownShape(shape: string): shape is UnknownShape {
  return shape.startsWith(UNKNOWN_PREFIX) && shape.length > UNKNOWN_PREFIX.length;
}

/** The three presets that predate K5 under a name OOXML does not use, by the `prst` they stand for. */
const PRST_OF_LEGACY: Readonly<Record<string, string>> = { rightTriangle: 'rtTriangle', arrowRight: 'rightArrow', arrowLeft: 'leftArrow' };
const LEGACY_OF_PRST: Readonly<Record<string, ShapePreset>> = { rtTriangle: 'rightTriangle' };

/** The OOXML `prst` name of a shape: the original name of an `unknown:` shape, else the preset's `prst` value. */
export function shapePrst(shape: ShapeKind): string {
  if (isUnknownShape(shape)) return shape.slice(UNKNOWN_PREFIX.length);
  return PRST_OF_LEGACY[shape] ?? shape;
}

/** The shape to keep for an OOXML `prst`: the built-in preset when there is one, else an `unknown:<prst>` passthrough. */
export function shapeFromPrst(prst: string): ShapeKind {
  const known = LEGACY_OF_PRST[prst] ?? prst;
  return isShapePreset(known) ? known : unknownShape(prst);
}

/**
 * The adjustment values to draw a preset with: each guide takes the given value (clamped to its range; a missing or
 * non-finite value takes the default). Presets without guides, and unknown shapes, resolve to an empty list.
 */
export function resolveAdjustments(shape: ShapeKind, adjustments?: readonly number[]): number[] {
  if (isUnknownShape(shape) || !isShapePreset(shape)) return [];
  return (SHAPE_PRESET_INFO[shape].adj ?? []).map((spec, i) => {
    const v = adjustments?.[i];
    return typeof v === 'number' && Number.isFinite(v) ? Math.min(spec.max, Math.max(spec.min, v)) : spec.def;
  });
}
