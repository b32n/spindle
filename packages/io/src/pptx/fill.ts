import type { PptxColor, PptxColorMap, PptxFill, PptxGeometry, PptxLine } from '@b32nio/spindle-shared/export/pptx';
import type { Rel } from '../common/ooxml-package';
import { kid, kids, type XmlElement } from '../xlsx/xml';
import { colorElement, parseColor } from './color';
import { SUPPORTED_PRESETS } from './slide-scan';
import { emuToPx, num, pct } from './units';

// Fills, outlines and geometry of a shape (plan K1 and K5): what the element states, no theme style lookup (the shape
// reader adds `p:style` references on top).

const FILL_NAMES = new Set(['noFill', 'solidFill', 'gradFill', 'pattFill', 'blipFill', 'grpFill']);
const CAPS: Record<string, PptxLine['cap']> = { flat: 'flat', rnd: 'round', sq: 'square' };

export type Rels = ReadonlyMap<string, Rel>;

const colorOf = (parent: XmlElement | undefined, map: PptxColorMap): PptxColor | undefined => parseColor(colorElement(parent), map);

/** The first fill element among the children of `holder` (`a:spPr`, `a:ln`, `p:bgPr`), read; `undefined` when there is none, it is `grpFill` or its colour is unreadable. */
export function parseFill(holder: XmlElement | undefined, map: PptxColorMap, rels: Rels): PptxFill | undefined {
  const el = holder?.children.find((c) => FILL_NAMES.has(c.name));
  switch (el?.name) {
    case 'noFill':
      return { kind: 'none' };
    case 'solidFill': {
      const color = colorOf(el, map);
      return color ? { kind: 'solid', color } : undefined;
    }
    case 'gradFill':
      return gradient(el, map);
    case 'pattFill': {
      const fg = colorOf(kid(el, 'fgClr'), map);
      const bg = colorOf(kid(el, 'bgClr'), map);
      return { kind: 'pattern', prst: el.attrs.prst ?? 'pct50', ...(fg ? { fg } : {}), ...(bg ? { bg } : {}) };
    }
    case 'blipFill': {
      const blip = kid(el, 'blip');
      const rel = rels.get(blip?.rattrs?.embed ?? '');
      const alpha = pct(kid(blip, 'alphaModFix')?.attrs.amt);
      return {
        kind: 'picture',
        ...(rel && !rel.external ? { part: rel.target } : {}),
        ...(kid(el, 'tile') ? { tile: true } : {}),
        ...(alpha !== undefined && alpha < 1 ? { alpha: Math.max(0, alpha) } : {}),
      };
    }
    default:
      return undefined;
  }
}

function gradient(el: XmlElement, map: PptxColorMap): PptxFill | undefined {
  const stops: { pos: number; color: PptxColor }[] = [];
  for (const gs of kids(kid(el, 'gsLst'), 'gs')) {
    const color = colorOf(gs, map);
    const pos = pct(gs.attrs.pos);
    if (color && pos !== undefined) stops.push({ pos: Math.min(1, Math.max(0, pos)), color });
  }
  if (stops.length === 0) return undefined;
  const path = kid(el, 'path');
  const lin = kid(el, 'lin');
  const shape = path ? (path.attrs.path === 'circle' ? 'radial' : path.attrs.path === 'rect' ? 'rect' : 'shape') : 'linear';
  const angle = (num(lin?.attrs.ang) ?? 0) / 60000;
  return {
    kind: 'gradient',
    shape,
    ...(shape === 'linear' ? { angle } : {}),
    ...(lin?.attrs.scaled === '1' || lin?.attrs.scaled === 'true' ? { scaled: true } : {}),
    ...(el.attrs.rotWithShape === '1' || el.attrs.rotWithShape === 'true' ? { rotWithShape: true } : {}),
    stops,
  };
}

/** `a:ln`: width in px, its fill, dash, cap, join and line ends, only what it states; `undefined` for an absent element. */
export function parseLine(ln: XmlElement | undefined, map: PptxColorMap, rels: Rels): PptxLine | undefined {
  if (!ln) return undefined;
  const out: PptxLine = {};
  const width = emuToPx(ln.attrs.w);
  if (width !== undefined && width >= 0) out.width = width;
  const fill = parseFill(ln, map, rels);
  if (fill) out.fill = fill;
  const dash = kid(ln, 'prstDash')?.attrs.val;
  if (dash) out.dash = dash;
  else if (kid(ln, 'custDash')) out.dash = 'custom';
  const cap = CAPS[ln.attrs.cap ?? ''];
  if (cap) out.cap = cap;
  for (const j of ['round', 'bevel', 'miter'] as const) if (kid(ln, j)) out.join = j;
  for (const [key, name] of [['head', 'headEnd'], ['tail', 'tailEnd']] as const) {
    const end = kid(ln, name);
    if (end && end.attrs.type && end.attrs.type !== 'none')
      out[key] = { type: end.attrs.type, ...(end.attrs.w ? { w: end.attrs.w } : {}), ...(end.attrs.len ? { len: end.attrs.len } : {}) };
  }
  return out;
}

/** `a:prstGeom` (name and the `val` adjustments) or `a:custGeom`; `undefined` when `spPr` states neither. */
export function parseGeometry(spPr: XmlElement | undefined): PptxGeometry | undefined {
  if (kid(spPr, 'custGeom')) return { kind: 'custom' };
  const geom = kid(spPr, 'prstGeom');
  const prst = geom?.attrs.prst;
  if (!prst) return undefined;
  const adjustments: Record<string, number> = {};
  for (const gd of kids(kid(geom, 'avLst'), 'gd')) {
    const m = /^val\s+(-?\d+(?:\.\d+)?)$/.exec((gd.attrs.fmla ?? '').trim());
    if (gd.attrs.name && m) adjustments[gd.attrs.name] = Number(m[1]);
  }
  return { kind: 'preset', prst, ...(Object.keys(adjustments).length > 0 ? { adjustments } : {}), known: SUPPORTED_PRESETS.has(prst) };
}
