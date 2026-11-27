import type { PptxBullet, PptxColorMap, PptxFrame, PptxLayout, PptxLevelStyle, PptxLevelStyles, PptxPlaceholder } from '@b32nio/spindle-shared/export/pptx';
import { isTrue, kid, kids, type XmlElement } from '../xlsx/xml';
import { colorElement, parseColor } from './color';
import { emuToPx, num, pct } from './units';

// What a level, a frame and a placeholder inherit (plan K4): the pieces `layouts.ts` (masters and layouts) and `shapes.ts`
// (slide shapes) share, so neither depends on the other.

export const TITLE_TYPES = new Set(['title', 'ctrTitle']);
const LEVELS = 9;

// `just`, `justLow`, `dist` and `thaiDist` all read as `justify`: the distinction (distributed spacing) is not modelled.
const ALIGN: Record<string, PptxLevelStyle['align']> = { l: 'left', ctr: 'center', r: 'right', just: 'justify', justLow: 'justify', dist: 'justify', thaiDist: 'justify' };

/** Spacing properties that state one value in two units: stating either in `over` replaces both in `base`. */
const SPACING_GROUPS: (keyof PptxLevelStyle)[][] = [
  ['spaceBefore', 'spaceBeforePct'],
  ['spaceAfter', 'spaceAfterPct'],
  ['lineHeight', 'lineSpacingPts'],
];

/** The merged level styles `over` on `base`, level by level, a property of `over` winning. */
export function mergeLevels(base: PptxLevelStyles, over: PptxLevelStyles): PptxLevelStyles {
  return Array.from({ length: LEVELS }, (_, i) => mergeLevel(base[i], over[i]));
}

/** One level over another (see `mergeLevels`). */
export function mergeLevel(base: PptxLevelStyle | undefined, over: PptxLevelStyle | undefined): PptxLevelStyle {
  const merged = { ...base };
  for (const group of SPACING_GROUPS) if (group.some((k) => over?.[k] !== undefined)) for (const k of group) delete merged[k];
  return { ...merged, ...over };
}

function parseBullet(p: XmlElement): PptxBullet | undefined {
  if (kid(p, 'buNone')) return { kind: 'none' };
  const char = kid(p, 'buChar')?.attrs.char;
  if (char !== undefined && char !== '') return { kind: 'char', char };
  const auto = kid(p, 'buAutoNum');
  if (auto) {
    const startAt = num(auto.attrs.startAt);
    return { kind: 'auto', scheme: auto.attrs.type ?? 'arabicPeriod', ...(startAt !== undefined && startAt !== 1 ? { startAt } : {}) };
  }
  return undefined;
}

/** `a:spcBef` / `a:spcAft`: points (`spcPts`) or a fraction of the font size (`spcPct`), a negative value being unusable. */
function spacing(e: XmlElement | undefined): { pts?: number; pct?: number } | undefined {
  const pts = num(kid(e, 'spcPts')?.attrs.val);
  if (pts !== undefined && pts >= 0) return { pts: pts / 100 };
  const fraction = pct(kid(e, 'spcPct')?.attrs.val);
  return fraction !== undefined && fraction >= 0 ? { pct: fraction } : undefined;
}

/** One `a:lvlNpPr` (or `a:defPPr`): paragraph properties and its `a:defRPr`, only what the element states. */
export function parseLevel(p: XmlElement | undefined, colorMap: PptxColorMap): PptxLevelStyle {
  const out: PptxLevelStyle = {};
  if (!p) return out;
  const r = kid(p, 'defRPr');
  const size = num(r?.attrs.sz);
  if (size !== undefined && size > 0) out.fontSize = size / 100;
  if (r?.attrs.b !== undefined) out.bold = isTrue(r.attrs.b);
  if (r?.attrs.i !== undefined) out.italic = isTrue(r.attrs.i);
  const color = parseColor(colorElement(kid(r, 'solidFill')), colorMap);
  if (color) out.color = color;
  const face = (kid(r, 'latin')?.attrs.typeface ?? '').trim();
  if (face === '+mj-lt') out.font = 'major';
  else if (face === '+mn-lt') out.font = 'minor';
  else if (face !== '' && !face.startsWith('+')) out.font = face;
  const align = ALIGN[p.attrs.algn ?? ''];
  if (align) out.align = align;
  const marginLeft = emuToPx(p.attrs.marL);
  if (marginLeft !== undefined) out.marginLeft = marginLeft;
  const indent = emuToPx(p.attrs.indent);
  if (indent !== undefined) out.indent = indent;
  const line = kid(p, 'lnSpc');
  const lineHeight = pct(kid(line, 'spcPct')?.attrs.val);
  const linePts = num(kid(line, 'spcPts')?.attrs.val);
  if (lineHeight !== undefined && lineHeight > 0) out.lineHeight = lineHeight;
  else if (linePts !== undefined && linePts > 0) out.lineSpacingPts = linePts / 100;
  const before = spacing(kid(p, 'spcBef'));
  if (before?.pts !== undefined) out.spaceBefore = before.pts;
  else if (before?.pct !== undefined) out.spaceBeforePct = before.pct;
  const after = spacing(kid(p, 'spcAft'));
  if (after?.pts !== undefined) out.spaceAfter = after.pts;
  else if (after?.pct !== undefined) out.spaceAfterPct = after.pct;
  const bullet = parseBullet(p);
  if (bullet) out.bullet = bullet;
  const bulletColor = parseColor(colorElement(kid(p, 'buClr')), colorMap);
  if (bulletColor) out.bulletColor = bulletColor;
  const bulletSize = pct(kid(p, 'buSzPct')?.attrs.val);
  if (bulletSize !== undefined && bulletSize > 0) out.bulletSizePct = bulletSize;
  const bulletFont = (kid(p, 'buFont')?.attrs.typeface ?? '').trim();
  if (bulletFont !== '' && !bulletFont.startsWith('+')) out.bulletFont = bulletFont;
  const tabs = kids(kid(p, 'tabLst'), 'tab').flatMap((t) => {
    const pos = emuToPx(t.attrs.pos);
    return pos === undefined ? [] : [{ pos, align: t.attrs.algn ?? 'l' }];
  });
  if (tabs.length > 0) out.tabStops = tabs;
  return out;
}

/** `a:lstStyle`, `p:titleStyle`, `p:bodyStyle`, `p:otherStyle` or `p:defaultTextStyle`: levels 1 to 9. */
export function parseLevels(list: XmlElement | undefined, colorMap: PptxColorMap): PptxLevelStyles {
  return Array.from({ length: LEVELS }, (_, i) => parseLevel(kid(list, `lvl${i + 1}pPr`), colorMap));
}

export function frameOf(xfrm: XmlElement | undefined): PptxFrame | undefined {
  const off = kid(xfrm, 'off');
  const ext = kid(xfrm, 'ext');
  const x = emuToPx(off?.attrs.x);
  const y = emuToPx(off?.attrs.y);
  const w = emuToPx(ext?.attrs.cx);
  const h = emuToPx(ext?.attrs.cy);
  if (x === undefined || y === undefined || w === undefined || h === undefined || w < 0 || h < 0) return undefined;
  const rot = (num(xfrm?.attrs.rot) ?? 0) / 60000;
  return { x, y, w, h, ...(rot !== 0 ? { rot } : {}), ...(isTrue(xfrm?.attrs.flipH) ? { flipH: true } : {}), ...(isTrue(xfrm?.attrs.flipV) ? { flipV: true } : {}) };
}

const family = (type: string): string => (TITLE_TYPES.has(type) ? 'title' : type);

/**
 * The layout placeholder a slide placeholder inherits from: the one with the same `idx`, else the first of the same
 * type, else (only when the slide placeholder has no `idx` or `idx` 0) the first of the same kind (the two title types and
 * `subTitle` / `body` / `obj` are interchangeable, as in PowerPoint), else `undefined`.
 */
export function matchLayoutPlaceholder(layout: PptxLayout | undefined, ph: { type?: string; idx?: number }): PptxPlaceholder | undefined {
  if (!layout) return undefined;
  const type = ph.type || 'obj';
  // `idx` 0 is the title / first body slot and is matched like an absent idx, by type.
  const byIdx = ph.idx !== undefined && ph.idx !== 0 ? layout.placeholders.find((p) => p.idx === ph.idx) : undefined;
  if (byIdx) return byIdx;
  const exact = layout.placeholders.find((p) => p.type === type);
  if (exact) return exact;
  // A slide placeholder naming an `idx` the layout lacks inherits nothing beyond an exact type match, as in PowerPoint.
  if (ph.idx !== undefined && ph.idx !== 0) return undefined;
  const textual = (t: string) => t === 'body' || t === 'obj' || t === 'subTitle';
  return layout.placeholders.find((p) => family(p.type) === family(type) || (textual(p.type) && textual(type)));
}
