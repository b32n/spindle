import type {
  PptxBullet,
  PptxColor,
  PptxColorMap,
  PptxFrame,
  PptxLayout,
  PptxLevelStyle,
  PptxLevelStyles,
  PptxPlaceholder,
} from '@b32nio/spindle-shared/export/pptx';
import type { OoxmlPackage, Rel } from '../common/ooxml-package';
import { descendants, isTrue, kid, kids, type XmlElement } from '../xlsx/xml';
import { colorElement, overrideColorMap, parseColor, parseColorMap } from './color';
import { baseName } from './parts';
import { parseTheme } from './theme';
import { emuToPx, num, pct } from './units';

// Masters and layouts, flattened (plan K4, decision SD6): the master's text styles, the master placeholder and the layout
// placeholder are merged into the effective values, so nothing downstream follows inheritance again.

const LEVELS = 9;
const ALIGN: Record<string, PptxLevelStyle['align']> = { l: 'left', ctr: 'center', r: 'right', just: 'justify', justLow: 'justify', dist: 'justify', thaiDist: 'justify' };
const ANCHOR: Record<string, PptxPlaceholder['anchor']> = { t: 'top', ctr: 'middle', b: 'bottom' };
/** Shape elements that can carry a placeholder (`p:sp`, `p:pic`, `p:graphicFrame`) and every element that counts as a drawn shape. */
const PLACEHOLDER_HOLDERS = new Set(['sp', 'pic', 'graphicFrame']);
const SHAPES = new Set(['sp', 'pic', 'graphicFrame', 'cxnSp', 'grpSp']);
/** Placeholder types that are matched to the master by type alone. */
const FIXED_TYPES = new Set(['dt', 'ftr', 'sldNum', 'hdr']);
const TITLE_TYPES = new Set(['title', 'ctrTitle']);

/** The merged level styles `over` on `base`, level by level, a property of `over` winning. */
export function mergeLevels(base: PptxLevelStyles, over: PptxLevelStyles): PptxLevelStyles {
  return Array.from({ length: LEVELS }, (_, i) => ({ ...base[i], ...over[i] }));
}

const emptyLevels = (): PptxLevelStyles => Array.from({ length: LEVELS }, () => ({}));

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
  const lineHeight = pct(kid(kid(p, 'lnSpc'), 'spcPct')?.attrs.val);
  if (lineHeight !== undefined && lineHeight > 0) out.lineHeight = lineHeight;
  const before = num(kid(kid(p, 'spcBef'), 'spcPts')?.attrs.val);
  if (before !== undefined && before >= 0) out.spaceBefore = before / 100;
  const after = num(kid(kid(p, 'spcAft'), 'spcPts')?.attrs.val);
  if (after !== undefined && after >= 0) out.spaceAfter = after / 100;
  const bullet = parseBullet(p);
  if (bullet) out.bullet = bullet;
  return out;
}

/** `a:lstStyle`, `p:titleStyle`, `p:bodyStyle`, `p:otherStyle` or `p:defaultTextStyle`: levels 1 to 9. */
export function parseLevels(list: XmlElement | undefined, colorMap: PptxColorMap): PptxLevelStyles {
  return Array.from({ length: LEVELS }, (_, i) => parseLevel(kid(list, `lvl${i + 1}pPr`), colorMap));
}

function frameOf(xfrm: XmlElement | undefined): PptxFrame | undefined {
  const off = kid(xfrm, 'off');
  const ext = kid(xfrm, 'ext');
  const x = emuToPx(off?.attrs.x);
  const y = emuToPx(off?.attrs.y);
  const w = emuToPx(ext?.attrs.cx);
  const h = emuToPx(ext?.attrs.cy);
  if (x === undefined || y === undefined || w === undefined || h === undefined || w < 0 || h < 0) return undefined;
  const rot = (num(xfrm?.attrs.rot) ?? 0) / 60000;
  return { x, y, w, h, ...(rot !== 0 ? { rot } : {}) };
}

interface RawPlaceholder {
  type: string;
  idx?: number;
  name?: string;
  frame?: PptxFrame;
  anchor?: PptxPlaceholder['anchor'];
  lstStyle?: XmlElement;
  prompt?: string;
}

const phOf = (e: XmlElement): XmlElement | undefined => kid(kid(e.children.find((c) => c.name.startsWith('nv')), 'nvPr'), 'ph');

/** The shapes directly under the shape tree: its placeholders and the number of drawn non-placeholder shapes. */
function scanTree(root: XmlElement | undefined): { placeholders: RawPlaceholder[]; shapes: number } {
  const placeholders: RawPlaceholder[] = [];
  let shapes = 0;
  for (const e of kid(kid(root, 'cSld'), 'spTree')?.children ?? []) {
    if (!SHAPES.has(e.name)) continue;
    const ph = PLACEHOLDER_HOLDERS.has(e.name) ? phOf(e) : undefined;
    if (!ph) {
      shapes++;
      continue;
    }
    const idx = num(ph.attrs.idx);
    const name = (kid(e.children.find((c) => c.name.startsWith('nv')), 'cNvPr')?.attrs.name ?? '').trim();
    const body = kid(e, 'txBody');
    const prompt = kids(body, 'p').map((p) => descendants(p, 't').map((t) => t.text).join('')).join('\n').trim();
    const anchor = ANCHOR[kid(body, 'bodyPr')?.attrs.anchor ?? ''];
    const frame = frameOf(e.name === 'graphicFrame' ? kid(e, 'xfrm') : kid(kid(e, 'spPr'), 'xfrm'));
    placeholders.push({
      type: ph.attrs.type || 'obj',
      ...(idx !== undefined ? { idx } : {}),
      ...(name !== '' ? { name } : {}),
      ...(frame ? { frame } : {}),
      ...(anchor ? { anchor } : {}),
      ...(kid(body, 'lstStyle') ? { lstStyle: kid(body, 'lstStyle') } : {}),
      ...(prompt !== '' ? { prompt } : {}),
    });
  }
  return { placeholders, shapes };
}

/** The master placeholder a layout placeholder inherits from: by type for dates, footers and numbers, the title for both title types, else the body. */
function masterMatch(master: RawPlaceholder[], type: string): RawPlaceholder | undefined {
  const same = (t: string) => master.find((m) => m.type === t);
  if (FIXED_TYPES.has(type)) return same(type);
  if (TITLE_TYPES.has(type)) return same('title') ?? same('ctrTitle');
  return same('body');
}

/** Which master text style a placeholder type starts from. */
const styleList = (txStyles: XmlElement | undefined, type: string): XmlElement | undefined =>
  kid(txStyles, TITLE_TYPES.has(type) ? 'titleStyle' : FIXED_TYPES.has(type) ? 'otherStyle' : 'bodyStyle');

/** A solid background: `p:bg/p:bgPr` with a `solidFill`, or `p:bgRef` with its colour. Anything else (gradient, picture, pattern) is `undefined` with `unsupported` set. */
function readBackground(root: XmlElement | undefined, map: PptxColorMap): { present: boolean; color?: PptxColor; approximated?: string } {
  const bg = kid(kid(root, 'cSld'), 'bg');
  if (!bg) return { present: false };
  const props = kid(bg, 'bgPr');
  if (props) {
    const solid = kid(props, 'solidFill');
    if (solid) return { present: true, color: parseColor(colorElement(solid), map) };
    const other = props.children.find((c) => /Fill$/.test(c.name) && c.name !== 'noFill');
    return { present: true, ...(other ? { approximated: other.name.replace(/Fill$/, '') } : {}) };
  }
  const ref = kid(bg, 'bgRef');
  const color = parseColor(colorElement(ref), map);
  const idx = num(ref?.attrs.idx);
  return { present: true, color, ...(idx !== undefined && idx !== 1001 && color ? { approximated: 'theme style' } : {}) };
}

/** The layouts of a master in `p:sldLayoutIdLst` order (else relationship order). */
function layoutParts(master: XmlElement, rels: Rel[]): string[] {
  const layoutRels = rels.filter((r) => r.type === 'slideLayout' && !r.external);
  const listed = kids(kid(master, 'sldLayoutIdLst'), 'sldLayoutId')
    .map((e) => layoutRels.find((r) => r.id === e.rattrs?.id)?.target)
    .filter((t): t is string => t !== undefined);
  const parts = listed.length > 0 ? listed : layoutRels.map((r) => r.target);
  return [...new Set(parts)];
}

export interface ReadLayoutsResult {
  layouts: PptxLayout[];
}

/**
 * Every layout of the given masters (parts, in order), flattened. `defaultStyle` is the presentation's
 * `p:defaultTextStyle`, the bottom of every placeholder's inheritance chain. A master or layout that cannot be read is
 * skipped with a warning. Masters whose theme differs from the first one are reported: the deck keeps one theme.
 */
export async function readLayouts(pkg: OoxmlPackage, masterParts: string[], defaultStyle: XmlElement | undefined, warn: (m: string) => void): Promise<PptxLayout[]> {
  const layouts: PptxLayout[] = [];
  let firstTheme: string | undefined;
  for (const masterPart of masterParts) {
    const master = await pkg.xml(masterPart);
    if (!master || master.name !== 'sldMaster') {
      warn(`${baseName(masterPart)} could not be read as a slide master; its layouts are skipped.`);
      continue;
    }
    const mrels = await pkg.rels(masterPart);
    const themePart = mrels.find((r) => r.type === 'theme' && !r.external)?.target;
    const themeRoot = themePart === undefined ? undefined : await pkg.xml(themePart);
    if (themeRoot) {
      const sig = JSON.stringify(parseTheme(themeRoot));
      firstTheme ??= sig;
      if (sig !== firstTheme) warn(`${baseName(masterPart)} uses a different theme from the first master; the deck keeps one theme, so its colours and fonts follow the first.`);
    }
    const masterMap = parseColorMap(kid(master, 'clrMap'));
    const masterTree = scanTree(master);
    const masterName = (kid(master, 'cSld')?.attrs.name ?? '').trim();
    for (const part of layoutParts(master, mrels)) {
      const layout = await pkg.xml(part);
      if (!layout || layout.name !== 'sldLayout') {
        warn(`${baseName(part)} could not be read as a slide layout; it is skipped.`);
        continue;
      }
      layouts.push(flattenLayout(part, masterPart, masterName, master, masterMap, masterTree, layout, defaultStyle, warn));
    }
  }
  return layouts;
}

function flattenLayout(
  part: string,
  masterPart: string,
  masterName: string,
  master: XmlElement,
  masterMap: PptxColorMap,
  masterTree: ReturnType<typeof scanTree>,
  layout: XmlElement,
  defaultStyle: XmlElement | undefined,
  warn: (m: string) => void
): PptxLayout {
  const map = overrideColorMap(masterMap, kid(kid(layout, 'clrMapOvr'), 'overrideClrMapping'));
  const tree = scanTree(layout);
  const txStyles = kid(master, 'txStyles');
  const base = parseLevels(defaultStyle, map);
  const name = (kid(layout, 'cSld')?.attrs.name ?? '').trim() || baseName(part).replace(/\.xml$/i, '');
  const hideMasterShapes = layout.attrs.showMasterSp === '0' || layout.attrs.showMasterSp === 'false';

  const placeholders = tree.placeholders.map((ph): PptxPlaceholder => {
    const m = masterMatch(masterTree.placeholders, ph.type);
    let levels = mergeLevels(emptyLevels(), base);
    levels = mergeLevels(levels, parseLevels(styleList(txStyles, ph.type), map));
    levels = mergeLevels(levels, parseLevels(m?.lstStyle, map));
    levels = mergeLevels(levels, parseLevels(ph.lstStyle, map));
    const frame = ph.frame ?? m?.frame;
    const anchor = ph.anchor ?? m?.anchor;
    return {
      type: ph.type,
      ...(ph.idx !== undefined ? { idx: ph.idx } : {}),
      ...(ph.name ? { name: ph.name } : {}),
      ...(frame ? { frame } : {}),
      ...(anchor ? { anchor } : {}),
      levels,
      ...(ph.prompt ? { prompt: ph.prompt } : {}),
    };
  });

  const own = readBackground(layout, map);
  const bg = own.present ? own : readBackground(master, masterMap);
  if (bg.approximated) warn(`${name}: the background is a ${bg.approximated} fill; it is read as a ${bg.color ? 'solid colour' : 'plain background'}.`);
  return {
    partName: part,
    name,
    ...(layout.attrs.type ? { layoutType: layout.attrs.type } : {}),
    masterPart,
    ...(masterName !== '' ? { masterName } : {}),
    ...(bg.color ? { background: bg.color } : {}),
    ...(hideMasterShapes ? { hideMasterShapes: true } : {}),
    placeholders,
    staticShapes: tree.shapes + (hideMasterShapes ? 0 : masterTree.shapes),
  };
}

const family = (type: string): string => (TITLE_TYPES.has(type) ? 'title' : type);

/**
 * The layout placeholder a slide placeholder inherits from: the one with the same `idx`, else the first of the same
 * type (the two title types and `subTitle` / `body` / `obj` are interchangeable, as in PowerPoint), else `undefined`.
 */
export function matchLayoutPlaceholder(layout: PptxLayout | undefined, ph: { type?: string; idx?: number }): PptxPlaceholder | undefined {
  if (!layout) return undefined;
  const type = ph.type || 'obj';
  const byIdx = ph.idx !== undefined && ph.idx !== 0 ? layout.placeholders.find((p) => p.idx === ph.idx) : undefined;
  if (byIdx) return byIdx;
  const exact = layout.placeholders.find((p) => p.type === type);
  if (exact) return exact;
  const textual = (t: string) => t === 'body' || t === 'obj' || t === 'subTitle';
  return layout.placeholders.find((p) => family(p.type) === family(type) || (textual(p.type) && textual(type)));
}
