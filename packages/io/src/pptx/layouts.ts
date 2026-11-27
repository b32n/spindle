import type {
  PptxColor,
  PptxColorMap,
  PptxBodyProps,
  PptxFrame,
  PptxLayout,
  PptxLevelStyles,
  PptxPlaceholder,
} from '@b32nio/spindle-shared/export/pptx';
import type { OoxmlPackage, Rel } from '../common/ooxml-package';
import { descendants, kid, kids, type XmlElement } from '../xlsx/xml';
import { colorElement, overrideColorMap, parseColor, parseColorMap } from './color';
import { collectFonts } from './fonts';
import { TITLE_TYPES, frameOf, mergeLevels, parseLevels } from './levels';
import { baseName } from './parts';
import { readShapes } from './shapes';
import { mergeBodyProps, parseBodyProps } from './text';
import { parseTheme } from './theme';
import { num } from './units';

// Masters and layouts, flattened (plan K4, decision SD6): the master's text styles, the master placeholder and the layout
// placeholder are merged into the effective values, so nothing downstream follows inheritance again.

const emptyLevels = (): PptxLevelStyles => Array.from({ length: 9 }, () => ({}));
const ANCHOR: Record<string, PptxPlaceholder['anchor']> = { t: 'top', ctr: 'middle', b: 'bottom' };
/** Shape elements that can carry a placeholder (`p:sp`, `p:pic`, `p:graphicFrame`) and every element that counts as a drawn shape. */
const PLACEHOLDER_HOLDERS = new Set(['sp', 'pic', 'graphicFrame']);
const SHAPES = new Set(['sp', 'pic', 'graphicFrame', 'cxnSp', 'grpSp']);
/** Placeholder types that are matched to the master by type alone. */
const FIXED_TYPES = new Set(['dt', 'ftr', 'sldNum', 'hdr']);

interface RawPlaceholder {
  type: string;
  idx?: number;
  name?: string;
  frame?: PptxFrame;
  anchor?: PptxPlaceholder['anchor'];
  bodyProps?: PptxBodyProps;
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
    const bodyProps = parseBodyProps(kid(body, 'bodyPr'));
    const frame = frameOf(e.name === 'graphicFrame' ? kid(e, 'xfrm') : kid(kid(e, 'spPr'), 'xfrm'));
    placeholders.push({
      type: ph.attrs.type || 'obj',
      ...(idx !== undefined ? { idx } : {}),
      ...(name !== '' ? { name } : {}),
      ...(frame ? { frame } : {}),
      ...(anchor ? { anchor } : {}),
      ...(Object.keys(bodyProps).length > 0 ? { bodyProps } : {}),
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
 * skipped with a warning; `fonts` collects the typefaces they name. Masters whose theme differs from the first one are reported: the deck keeps one theme.
 */
export async function readLayouts(pkg: OoxmlPackage, masterParts: string[], defaultStyle: XmlElement | undefined, warn: (m: string) => void, fonts: Set<string>): Promise<PptxLayout[]> {
  const layouts: PptxLayout[] = [];
  const warnedBackgrounds = new Set<string>();
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
    collectFonts(master, fonts);
    const masterMap = parseColorMap(kid(master, 'clrMap'));
    const masterTree = scanTree(master);
    const masterName = (kid(master, 'cSld')?.attrs.name ?? '').trim();
    for (const part of layoutParts(master, mrels)) {
      const layout = await pkg.xml(part);
      if (!layout || layout.name !== 'sldLayout') {
        warn(`${baseName(part)} could not be read as a slide layout; it is skipped.`);
        continue;
      }
      collectFonts(layout, fonts);
      const lrels = new Map((await pkg.rels(part)).map((r) => [r.id, r]));
      layouts.push(flattenLayout(part, masterPart, masterName, master, masterMap, new Map(mrels.map((r) => [r.id, r])), masterTree, layout, lrels, defaultStyle, warn, warnedBackgrounds));
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
  masterRels: ReadonlyMap<string, Rel>,
  masterTree: ReturnType<typeof scanTree>,
  layout: XmlElement,
  layoutRels: ReadonlyMap<string, Rel>,
  defaultStyle: XmlElement | undefined,
  warn: (m: string) => void,
  warnedBackgrounds: Set<string>
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
    const bodyProps = mergeBodyProps(m?.bodyProps, ph.bodyProps ?? {});
    return {
      type: ph.type,
      ...(ph.idx !== undefined ? { idx: ph.idx } : {}),
      ...(ph.name ? { name: ph.name } : {}),
      ...(frame ? { frame } : {}),
      ...(anchor ? { anchor } : {}),
      ...(Object.keys(bodyProps).length > 0 ? { bodyProps } : {}),
      levels,
      ...(ph.prompt ? { prompt: ph.prompt } : {}),
    };
  });

  const defaultLevels = base;
  const elements = [
    ...(hideMasterShapes ? [] : readShapes(kid(kid(master, 'cSld'), 'spTree'), { map: masterMap, rels: masterRels, defaultLevels: parseLevels(defaultStyle, masterMap), warn }, false)),
    ...readShapes(kid(kid(layout, 'cSld'), 'spTree'), { map, rels: layoutRels, defaultLevels, warn }, false),
  ];
  const own = readBackground(layout, map);
  const bg = own.present ? own : readBackground(master, masterMap);
  // One warning per background source: a master gradient shared by eleven layouts is reported once, under the master.
  const source = own.present ? part : masterPart;
  if (bg.approximated && !warnedBackgrounds.has(source)) {
    warnedBackgrounds.add(source);
    const label = own.present ? name : masterName || baseName(masterPart);
    warn(`${label}: the background is a ${bg.approximated} fill; it is read as a ${bg.color ? 'solid colour' : 'plain background'}.`);
  }
  return {
    partName: part,
    name,
    ...(layout.attrs.type ? { layoutType: layout.attrs.type } : {}),
    masterPart,
    ...(masterName !== '' ? { masterName } : {}),
    ...(bg.color ? { background: bg.color } : {}),
    ...(hideMasterShapes ? { hideMasterShapes: true } : {}),
    placeholders,
    elements,
    // Shapes inside `mc:AlternateContent` or other wrappers are not counted: the count is advisory (the shape reader owns them).
    staticShapes: tree.shapes + (hideMasterShapes ? 0 : masterTree.shapes),
  };
}

