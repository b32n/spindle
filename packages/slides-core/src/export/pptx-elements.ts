// Element mappings between the PPTX interchange and the slides model, both ways: shapes and text boxes, pictures, tables,
// connectors, groups and opaque content. The interchange holds slide-space frames (a group's transform is already applied), so a
// group is only its members' `groupId` / `parentGroupId` here, and is rebuilt from them on the way out.

import type {
  PptxCellBorders,
  PptxConnection,
  PptxConnector,
  PptxCustomPath,
  PptxElement,
  PptxFrame,
  PptxGeometry,
  PptxGroup,
  PptxLine,
  PptxOpaque,
  PptxPicture,
  PptxShape,
  PptxTable,
  PptxTableCell,
} from '@b32nio/spindle-shared/export/pptx';
import { generateId } from '../utils/id';
import { indexesBetween } from '../scene/fractional-index';
import {
  createImageElement,
  createLineElement,
  createOpaqueElement,
  createShapeElement,
  createTableElement,
  createTextElement,
} from '../scene/elements';
import { connectionFrame, resolveConnectorFrame } from '../scene/connector';
import { frameAABB, unionAABB } from '../scene/geometry';
import { groupChain, groupParents } from '../scene/group-tree';
import { SHAPE_PRESET_INFO, isUnknownShape, shapeFromPrst, shapePrst } from '../scene/shape-presets';
import { evenFractions, resolveCellEdges } from '../scene/table';
import type {
  AnchorId,
  ArrowHead,
  BodyStyle,
  CellBorders,
  CellSide,
  ConnectorKind,
  CustomGeometry,
  ImageElement,
  LineElement,
  OpaqueElement,
  PathFillMode,
  PlaceholderMeta,
  PlaceholderType,
  ShapePreset,
  ShapeElement,
  ShapeKind,
  SlideElement,
  Stroke,
  TableCell,
  TableElement,
  TextElement,
} from '../scene/types';
import { emptyRichText, richTextToPlainText } from '../text/model';
import type { TextTally } from './pptx-text';
import { bodyStyleFromPptx, newTextTally, textFromPptx, textToPptx } from './pptx-text';
import { fillFromPptx, fillToPptx, finite, strokeFromPptx, strokeToPptx } from './pptx-values';

/** What an import changed or left out in the elements, counted for the importer's warnings. */
export interface ElementTally extends TextTally {
  hidden: number;
  pictureFills: number;
  unreadablePictures: number;
  customOutlines: number;
  emptyTables: number;
}

export const newElementTally = (): ElementTally => ({ ...newTextTally(), hidden: 0, pictureFills: 0, unreadablePictures: 0, customOutlines: 0, emptyTables: 0 });

// ── Placeholders ─────────────────────────────────────────────────────────────

const PLACEHOLDER_OF: Record<string, PlaceholderType> = { title: 'title', ctrTitle: 'centerTitle', subTitle: 'subtitle', body: 'body', obj: 'body', pic: 'image' };
const PH_TYPE_OF: Record<PlaceholderType, string> = { title: 'title', centerTitle: 'ctrTitle', subtitle: 'subTitle', body: 'body', image: 'pic' };

/** The placeholder type of the interchange (`p:ph/@type`) in the model, or `undefined` for one the model has no place for (date, footer, slide number...). */
export const placeholderTypeFromPptx = (type: string): PlaceholderType | undefined => PLACEHOLDER_OF[type];
export const placeholderTypeToPptx = (type: PlaceholderType): string => PH_TYPE_OF[type];

function placeholderFromPptx(ph: { type: string; idx?: number } | undefined): PlaceholderMeta | undefined {
  const type = ph ? placeholderTypeFromPptx(ph.type) : undefined;
  return ph && type ? { type, idx: ph.idx ?? 0 } : undefined;
}

// ── Frames ───────────────────────────────────────────────────────────────────

const frameOf = (f: PptxFrame | undefined) => ({ x: f?.x ?? 0, y: f?.y ?? 0, w: f?.w ?? 0, h: f?.h ?? 0, rotation: f?.rot ?? 0 });

function frameToPptx(e: { x: number; y: number; w: number; h: number; rotation: number }, flipH?: boolean, flipV?: boolean): PptxFrame {
  return { x: e.x, y: e.y, w: e.w, h: e.h, ...(e.rotation ? { rot: e.rotation } : {}), ...(flipH ? { flipH } : {}), ...(flipV ? { flipV } : {}) };
}

// ── Geometry ─────────────────────────────────────────────────────────────────

/** Presets whose first guides are angles (60000ths of a degree in the file; the model keeps a fraction of a turn). */
const ANGLE_GUIDES: Record<string, number> = { pie: 2, chord: 2, blockArc: 2 };
const TURN = 21600000;
const FRACTION = 100000;

/** The file's guide names for a preset with `count` adjustments: `adj`, or `adj1`... (`mathPlus` and `mathMinus` count from 1 for a single guide). */
function guideNames(prst: string, count: number): string[] {
  if (count === 1) return [prst === 'mathPlus' || prst === 'mathMinus' ? 'adj1' : 'adj'];
  return Array.from({ length: count }, (_, i) => `adj${i + 1}`);
}

const guideScale = (prst: string, i: number): number => (i < (ANGLE_GUIDES[prst] ?? 0) ? TURN : FRACTION);

/** The model's adjustments for a preset's guides: a built-in takes them by name (missing guides take the default), an unknown preset keeps the raw values in file order. */
function adjustmentsFromPptx(shape: ShapeKind, geometry: Extract<PptxGeometry, { kind: 'preset' }>): number[] | undefined {
  const given = geometry.adjustments;
  if (!given) return undefined;
  const values = Object.entries(given).filter(([, v]) => Number.isFinite(v));
  if (isUnknownShape(shape)) return values.length > 0 ? values.map(([, v]) => v) : undefined;
  const specs = SHAPE_PRESET_INFO[shape as ShapePreset].adj;
  if (!specs) return undefined;
  const prst = shapePrst(shape);
  const names = guideNames(prst, specs.length);
  if (!names.some((n) => Number.isFinite(given[n]))) return undefined;
  const resolved = specs.map((spec, i) => (Number.isFinite(given[names[i]]) ? given[names[i]] / guideScale(prst, i) : spec.def));
  // guides at their defaults say nothing: the preset draws the same without them
  return resolved.every((v, i) => Math.abs(v - specs[i].def) < 1e-9) ? undefined : resolved;
}

function adjustmentsToPptx(shape: ShapeKind, adjustments: readonly number[] | undefined): Record<string, number> | undefined {
  if (!adjustments || adjustments.length === 0) return undefined;
  const values = adjustments.map(finite);
  if (isUnknownShape(shape)) {
    const names = guideNames(shapePrst(shape), values.length);
    return Object.fromEntries(values.flatMap((v, i) => (v === undefined ? [] : [[names[i], v]])));
  }
  const specs = SHAPE_PRESET_INFO[shape as ShapePreset].adj;
  if (!specs) return undefined;
  const prst = shapePrst(shape);
  const names = guideNames(prst, specs.length);
  const out: Record<string, number> = {};
  specs.forEach((spec, i) => {
    const v = values[i];
    if (v !== undefined && v !== spec.def) out[names[i]] = Math.round(v * guideScale(prst, i));
  });
  return Object.keys(out).length > 0 ? out : undefined;
}

const PATH_FILLS = new Set<PathFillMode>(['norm', 'none', 'lighten', 'lightenLess', 'darken', 'darkenLess']);

function customFromPptx(paths: PptxCustomPath[]): CustomGeometry {
  return {
    paths: paths.map((p) => ({
      w: p.w,
      h: p.h,
      ...(p.fill !== undefined && PATH_FILLS.has(p.fill as PathFillMode) && p.fill !== 'norm' ? { fill: p.fill as PathFillMode } : {}),
      ...(p.stroke === false ? { stroke: false } : {}),
      ...(p.extrusionOk ? { extrusionOk: true } : {}),
      commands: p.commands.map((c) => ({ ...c })),
    })),
  };
}

function customToPptx(g: CustomGeometry): PptxCustomPath[] {
  return g.paths.map((p) => ({
    w: p.w,
    h: p.h,
    ...(p.fill && p.fill !== 'norm' ? { fill: p.fill } : {}),
    ...(p.stroke === false ? { stroke: false as const } : {}),
    ...(p.extrusionOk ? { extrusionOk: true as const } : {}),
    commands: p.commands.map((c) => ({ ...c })),
  }));
}

// ── Shapes and text boxes ────────────────────────────────────────────────────

const isRect = (g: PptxGeometry | undefined): boolean => !g || (g.kind === 'preset' && g.prst === 'rect');

function textBodyOf(shape: PptxShape, tally: ElementTally) {
  return shape.text ? textFromPptx(shape.text, tally) : undefined;
}

/** A `p:sp` as a text element (a text box, a placeholder, or a plain rectangle with text) or a shape element. */
function shapeFromPptx(s: PptxShape, containerId: string, tally: ElementTally): TextElement | ShapeElement {
  const placeholder = placeholderFromPptx(s.placeholder);
  const base = { containerId, index: '', ...frameOf(s.frame), ...(s.pptxId !== '' ? { pptxId: s.pptxId } : {}), ...(placeholder ? { placeholder } : {}) };
  const fill = fillFromPptx(s.fill);
  if (s.fill?.kind === 'picture') tally.pictureFills++;
  const stroke = strokeFromPptx(s.line);
  const plain = (!fill || fill.kind === 'none') && !stroke;
  if (isRect(s.geometry) && (s.textBox || placeholder || (s.text && plain))) {
    const body = textBodyOf(s, tally);
    return createTextElement({
      ...base,
      ...(body ?? { richText: emptyRichText(), bodyStyle: bodyStyleFromPptx({}, tally) }),
      ...(fill ? { fill } : {}),
      ...(stroke ? { stroke } : {}),
    });
  }
  let shape: ShapeKind = 'rect';
  let path: CustomGeometry | undefined;
  let adjustments: number[] | undefined;
  if (s.geometry?.kind === 'preset') {
    shape = shapeFromPrst(s.geometry.prst);
    adjustments = adjustmentsFromPptx(shape, s.geometry);
  } else if (s.geometry?.kind === 'custom') {
    if (s.geometry.paths && s.geometry.paths.length > 0) {
      shape = 'path';
      path = customFromPptx(s.geometry.paths);
    } else tally.customOutlines++;
  }
  const body = textBodyOf(s, tally);
  return createShapeElement({
    ...base,
    shape,
    // a shape that states no fill has none (the interchange states it when it is the shape's own style)
    fill: fill ?? { kind: 'none' },
    ...(path ? { path } : {}),
    ...(adjustments ? { adjustments } : {}),
    ...(s.frame?.flipH ? { flipH: true } : {}),
    ...(s.frame?.flipV ? { flipV: true } : {}),
    ...(stroke ? { stroke } : {}),
    ...(body ? { richText: body.richText, bodyStyle: body.bodyStyle } : {}),
  });
}

function textElementToPptx(el: TextElement, name: string): PptxShape {
  const fill = fillToPptx(el.fill);
  const line = strokeToPptx(el.stroke);
  const text = textToPptx(el.richText, el.bodyStyle);
  return {
    kind: 'shape',
    pptxId: el.pptxId as string,
    name,
    ...(el.placeholder ? { placeholder: { type: placeholderTypeToPptx(el.placeholder.type), ...(el.placeholder.idx !== 0 ? { idx: el.placeholder.idx } : {}) } } : { textBox: true }),
    frame: frameToPptx(el),
    geometry: { kind: 'preset', prst: 'rect', known: true },
    ...(fill ? { fill } : {}),
    ...(line ? { line } : {}),
    text,
  };
}

function shapeElementToPptx(el: ShapeElement, name: string): PptxShape {
  const fill = fillToPptx(el.fill);
  const line = strokeToPptx(el.stroke);
  const adjusted = el.shape === 'path' ? undefined : adjustmentsToPptx(el.shape, el.adjustments);
  const geometry: PptxGeometry =
    el.shape === 'path' && el.path
      ? { kind: 'custom', paths: customToPptx(el.path) }
      : { kind: 'preset', prst: el.shape === 'path' ? 'rect' : shapePrst(el.shape), known: !isUnknownShape(el.shape), ...(adjusted ? { adjustments: adjusted } : {}) };
  return {
    kind: 'shape',
    pptxId: el.pptxId as string,
    name,
    ...(el.placeholder ? { placeholder: { type: placeholderTypeToPptx(el.placeholder.type), ...(el.placeholder.idx !== 0 ? { idx: el.placeholder.idx } : {}) } } : {}),
    frame: frameToPptx(el, el.flipH, el.flipV),
    geometry,
    ...(fill ? { fill } : {}),
    ...(line ? { line } : {}),
    ...(el.richText && richTextToPlainText(el.richText) !== '' ? { text: textToPptx(el.richText, el.bodyStyle) } : {}),
  };
}

// ── Pictures ─────────────────────────────────────────────────────────────────

const DATA_IMAGE = /^data:image\//;

function pictureFromPptx(p: PptxPicture, containerId: string, tally: ElementTally): ImageElement | ShapeElement {
  const placeholder = placeholderFromPptx(p.placeholder);
  const base = { containerId, index: '', ...frameOf(p.frame), ...(p.pptxId !== '' ? { pptxId: p.pptxId } : {}), ...(placeholder ? { placeholder } : {}) };
  if (p.dataUrl === undefined || !DATA_IMAGE.test(p.dataUrl)) {
    tally.unreadablePictures++;
    const label = `Picture: ${p.name} (${p.unsupported ?? 'not shown'})`;
    return createShapeElement({
      ...base,
      shape: 'rect',
      fill: { kind: 'none' },
      stroke: { color: { kind: 'theme', slot: 'dk2' }, width: 1, dash: 'dash' },
      richText: { type: 'doc', content: [{ type: 'paragraph', attrs: { align: 'center' }, content: [{ type: 'text', text: label }] }] },
      bodyStyle: { vAlign: 'middle', insets: { l: 8, t: 8, r: 8, b: 8 } },
    });
  }
  const crop = p.crop;
  const el = createImageElement({
    ...base,
    src: p.dataUrl,
    naturalW: Math.max(1, base.w),
    naturalH: Math.max(1, base.h),
    ...(p.frame?.flipH ? { flipH: true } : {}),
    ...(p.frame?.flipV ? { flipV: true } : {}),
    fit: 'fill',
    ...(crop ? { crop } : {}),
  });
  if (p.alpha !== undefined && p.alpha < 1) el.opacity = p.alpha;
  return el;
}

function imageToPptx(el: ImageElement, name: string): PptxPicture {
  const c = el.crop;
  const cropped = c && (c.l !== 0 || c.t !== 0 || c.r !== 0 || c.b !== 0);
  return {
    kind: 'picture',
    pptxId: el.pptxId as string,
    name,
    ...(el.placeholder ? { placeholder: { type: placeholderTypeToPptx(el.placeholder.type), ...(el.placeholder.idx !== 0 ? { idx: el.placeholder.idx } : {}) } } : {}),
    frame: frameToPptx(el, el.flipH, el.flipV),
    dataUrl: el.src,
    ...(cropped ? { crop: { l: c.l, t: c.t, r: c.r, b: c.b } } : {}),
    ...(el.opacity !== undefined && el.opacity < 1 ? { alpha: el.opacity } : {}),
  };
}

// ── Tables ───────────────────────────────────────────────────────────────────

const DEFAULT_CELL_INSETS = { l: 9.6, t: 4.8, r: 9.6, b: 4.8 };
const SIDES = ['l', 't', 'r', 'b'] as const;

function cellFromPptx(c: PptxTableCell, tally: ElementTally): TableCell {
  const { richText, bodyStyle } = textFromPptx(c.text, tally, c.anchor);
  const margins = Object.fromEntries(Object.entries(c.margins ?? {}).filter(([, v]) => v !== undefined));
  const style: BodyStyle = { ...bodyStyle, vAlign: c.anchor ?? 'top', insets: { ...DEFAULT_CELL_INSETS, ...margins } };
  const fill = fillFromPptx(c.fill);
  if (c.fill?.kind === 'picture') tally.pictureFills++;
  // sides as the file states them: a line, `null` for an explicit no-line, absent for nothing said (settled over the grid, `settleBorders`)
  const borders: CellBorders = {};
  for (const side of SIDES) {
    const b = c.borders?.[side];
    if (b === null) borders[side] = null;
    else if (b) borders[side] = strokeFromPptx(b) ?? null;
  }
  return {
    richText,
    bodyStyle: style,
    ...(fill ? { fill } : {}),
    ...(c.gridSpan !== undefined && c.gridSpan > 1 ? { gridSpan: c.gridSpan } : {}),
    ...(c.rowSpan !== undefined && c.rowSpan > 1 ? { rowSpan: c.rowSpan } : {}),
    ...(c.hMerge ? { hMerge: true } : {}),
    ...(c.vMerge ? { vMerge: true } : {}),
    borders,
  };
}

const FACING: Record<CellSide, [number, number, CellSide]> = { l: [0, -1, 'r'], r: [0, 1, 'l'], t: [-1, 0, 'b'], b: [1, 0, 't'] };

/**
 * A side the file says nothing about has no line, but the model's absent side takes the table's grid line: so it is made an explicit
 * no-line unless the cell across the edge states a line there (the model reads a shared edge from either side, so that line shows).
 */
function settleBorders(cells: TableCell[][]): void {
  cells.forEach((row, r) =>
    row.forEach((cell, c) => {
      for (const side of SIDES) {
        if (cell.borders?.[side] !== undefined) continue;
        const [dr, dc, face] = FACING[side];
        const across = cells[r + dr]?.[c + dc]?.borders?.[face];
        if (across) continue;
        cell.borders = { ...cell.borders, [side]: null };
      }
    })
  );
}

function tableFromPptx(t: PptxTable, containerId: string, tally: ElementTally): TableElement | undefined {
  const cols = t.columns.length;
  const rows = t.rows.length;
  if (cols === 0 || rows === 0) {
    tally.emptyTables++;
    return undefined;
  }
  const width = t.columns.reduce((s, w) => s + w, 0);
  const height = t.rows.reduce((s, r) => s + r.height, 0);
  const fractions = (sizes: number[], total: number): number[] => (total > 0 ? sizes.map((s) => s / total) : evenFractions(sizes.length));
  const frame = t.frame ?? { x: 0, y: 0, w: width, h: height };
  const el = createTableElement({
    containerId,
    index: '',
    ...frameOf(frame),
    ...(t.pptxId !== '' ? { pptxId: t.pptxId } : {}),
    rows,
    cols,
    colFractions: fractions(t.columns, width),
    rowFractions: fractions(t.rows.map((r) => r.height), height),
    cells: t.rows.map((r) => r.cells.map((c) => cellFromPptx(c, tally))),
    ...(t.flags ? { styleFlags: { ...t.flags } } : {}),
  });
  settleBorders(el.cells);
  el.rowHeights = t.rows.map((r) => r.height);
  return el;
}

const DEFAULT_GRID: Stroke = { color: { kind: 'theme', slot: 'dk2' }, width: 1 };

function tableToPptx(t: TableElement, name: string): PptxTable {
  const frame = frameToPptx(t);
  const grid = t.border ?? DEFAULT_GRID;
  const columns = t.colFractions.map((f) => f * t.w);
  const rows = t.cells.map((cells, r) => ({
    height: t.rowHeights?.[r] || t.rowFractions[r] * t.h,
    cells: cells.map((cell, c): PptxTableCell => {
      const edges = resolveCellEdges(t, r, c, grid);
      const borders: PptxCellBorders = {};
      for (const side of SIDES) borders[side] = edges[side] ? (strokeToPptx(edges[side] as Stroke) as PptxLine) : null;
      const insets = cell.bodyStyle?.insets;
      return {
        text: textToPptx(cell.richText, cell.bodyStyle),
        ...(cell.gridSpan ? { gridSpan: cell.gridSpan } : {}),
        ...(cell.rowSpan ? { rowSpan: cell.rowSpan } : {}),
        ...(cell.hMerge ? { hMerge: true } : {}),
        ...(cell.vMerge ? { vMerge: true } : {}),
        ...(fillToPptx(cell.fill) ? { fill: fillToPptx(cell.fill) } : {}),
        borders,
        ...(insets ? { margins: { l: insets.l, t: insets.t, r: insets.r, b: insets.b } } : {}),
        ...(cell.bodyStyle?.vAlign ? { anchor: cell.bodyStyle.vAlign } : {}),
      };
    }),
  }));
  const flags = Object.fromEntries(Object.entries(t.styleFlags ?? {}).filter(([, v]) => v === true)) as PptxTable['flags'];
  return { kind: 'table', pptxId: t.pptxId as string, name, frame, columns, rows, ...(flags && Object.keys(flags).length > 0 ? { flags } : {}) };
}

// ── Connectors ───────────────────────────────────────────────────────────────

const ARROW_OF: Record<string, ArrowHead> = { triangle: 'triangle', stealth: 'triangle', diamond: 'circle', arrow: 'arrow', oval: 'circle' };
const END_OF: Record<ArrowHead, string | undefined> = { none: undefined, arrow: 'arrow', triangle: 'triangle', circle: 'oval' };

/** The connection sites of a shape in the file's order (`cxnLst`), as the model's anchors. */
const RECT_SITES: AnchorId[] = ['n', 'w', 's', 'e'];
const ELLIPSE_SITES: AnchorId[] = ['n', 'nw', 'w', 'sw', 's', 'se', 'e', 'ne'];
const TRIANGLE_SITES: AnchorId[] = ['n', 'nw', 'sw', 's', 'se', 'ne'];
const SITES: Record<string, AnchorId[]> = {
  ellipse: ELLIPSE_SITES,
  flowChartConnector: ELLIPSE_SITES,
  donut: ELLIPSE_SITES,
  noSmoking: ELLIPSE_SITES,
  smileyFace: ELLIPSE_SITES,
  wedgeEllipseCallout: ELLIPSE_SITES,
  triangle: TRIANGLE_SITES,
};

const sitesOf = (el: SlideElement | undefined): AnchorId[] => (el?.type === 'shape' ? (SITES[shapePrst(el.shape)] ?? RECT_SITES) : RECT_SITES);

function connectorFromPptx(c: PptxConnector, containerId: string): LineElement {
  const prst = c.geometry?.kind === 'preset' ? c.geometry.prst : 'straightConnector1';
  const connector: ConnectorKind = prst.startsWith('bent') ? 'elbow' : prst.startsWith('curved') ? 'curved' : 'straight';
  const adj = c.geometry?.kind === 'preset' ? c.geometry.adjustments?.adj1 : undefined;
  const stroke = strokeFromPptx(c.line) ?? (c.line?.fill?.kind === 'none' ? { color: { kind: 'theme' as const, slot: 'dk1' as const, alpha: 0 }, width: 1 } : { color: { kind: 'theme' as const, slot: 'dk1' as const }, width: 1 });
  const start = c.line?.head ? ARROW_OF[c.line.head.type] : undefined;
  const end = c.line?.tail ? ARROW_OF[c.line.tail.type] : undefined;
  return createLineElement({
    containerId,
    index: '',
    ...frameOf(c.frame),
    ...(c.pptxId !== '' ? { pptxId: c.pptxId } : {}),
    stroke,
    connector,
    ...(connector === 'elbow' && adj !== undefined ? { adjustments: [adj / FRACTION] } : {}),
    ...(c.frame?.flipH ? { flipH: true } : {}),
    ...(c.frame?.flipV ? { flipV: true } : {}),
    ...(start ? { startArrow: start } : {}),
    ...(end ? { endArrow: end } : {}),
  });
}

/** Glue the connectors of a slide to the shapes the file names (`stCxn`, `endCxn`), once every element exists. */
function bindConnectors(pairs: Array<{ el: LineElement; source: PptxConnector }>, byPptxId: Map<string, SlideElement>): void {
  for (const { el, source } of pairs) {
    const bind = (c: PptxConnection | undefined) => {
      const target = c ? byPptxId.get(c.id) : undefined;
      const anchor = target ? sitesOf(target)[c!.idx] : undefined;
      return target && anchor ? { elementId: target.id, anchor } : undefined;
    };
    const startBind = bind(source.start);
    const endBind = bind(source.end);
    if (startBind) el.startBind = startBind;
    if (endBind) el.endBind = endBind;
    // the end that is not glued keeps the corner of the box it had, since the glued one now moves with its shape
    const sx = el.flipH ? el.x + el.w : el.x;
    const sy = el.flipV ? el.y + el.h : el.y;
    const ex = el.flipH ? el.x : el.x + el.w;
    const ey = el.flipV ? el.y : el.y + el.h;
    if (startBind && !endBind) el.endPoint = { x: ex, y: ey };
    if (endBind && !startBind) el.startPoint = { x: sx, y: sy };
  }
}

function lineToPptx(el: LineElement, name: string, byId: Map<string, SlideElement>): PptxConnector {
  const box = resolveConnectorFrame(el, (id) => {
    const t = byId.get(id);
    return t ? connectionFrame(t) : undefined;
  });
  const line = strokeToPptx(el.stroke) as PptxLine;
  const head = el.startArrow && END_OF[el.startArrow] ? { type: END_OF[el.startArrow] as string } : undefined;
  const tail = el.endArrow && END_OF[el.endArrow] ? { type: END_OF[el.endArrow] as string } : undefined;
  const connection = (bind: LineElement['startBind']): PptxConnection | undefined => {
    const target = bind ? byId.get(bind.elementId) : undefined;
    const idx = target ? sitesOf(target).indexOf(bind!.anchor) : -1;
    return target?.pptxId !== undefined && idx >= 0 ? { id: target.pptxId, idx } : undefined;
  };
  const prst = el.connector === 'elbow' ? 'bentConnector3' : el.connector === 'curved' ? 'curvedConnector3' : 'straightConnector1';
  const adj = el.connector === 'elbow' ? finite(el.adjustments?.[0]) : undefined;
  const start = connection(el.startBind);
  const end = connection(el.endBind);
  return {
    kind: 'connector',
    pptxId: el.pptxId as string,
    name,
    frame: frameToPptx({ ...box, rotation: el.rotation }, box.flipH, box.flipV),
    geometry: { kind: 'preset', prst, known: true, ...(adj !== undefined && adj !== 0.5 ? { adjustments: { adj1: Math.round(adj * FRACTION) } } : {}) },
    line: { ...line, ...(head ? { head } : {}), ...(tail ? { tail } : {}) },
    ...(start ? { start } : {}),
    ...(end ? { end } : {}),
  };
}

// ── Opaque content ───────────────────────────────────────────────────────────

function opaqueFromPptx(o: PptxOpaque, containerId: string): OpaqueElement {
  return createOpaqueElement({
    containerId,
    index: '',
    ...frameOf(o.frame),
    ...(o.pptxId !== '' ? { pptxId: o.pptxId } : {}),
    kind: o.opaqueKind,
    partRefs: o.parts,
    ...(o.altText ? { alt: o.altText } : {}),
    ...(o.title ? { title: o.title } : {}),
    ...(o.poster ? { poster: o.poster } : {}),
    ...(o.props ? { props: o.props } : {}),
  });
}

function opaqueToPptx(el: OpaqueElement, name: string): PptxOpaque {
  return {
    kind: 'opaque',
    opaqueKind: el.kind,
    pptxId: el.pptxId as string,
    name,
    ...(el.alt ? { altText: el.alt } : {}),
    frame: frameToPptx(el),
    ...(el.title ? { title: el.title } : {}),
    parts: el.partRefs.map((r) => ({ ...r })),
    ...(el.poster ? { poster: el.poster } : {}),
    ...(el.props ? { props: { ...el.props } } : {}),
  };
}

// ── A slide's element list ───────────────────────────────────────────────────

/**
 * The model's elements for the interchange elements of one slide (or layout), in drawing order with fresh ids and z-order keys.
 * A hidden element is left out (counted), a group becomes the `groupId` of its members, a connector is glued to its shapes at the end.
 */
export function elementsFromPptx(list: PptxElement[], containerId: string, tally: ElementTally): SlideElement[] {
  const out: SlideElement[] = [];
  const byPptxId = new Map<string, SlideElement>();
  const connectors: Array<{ el: LineElement; source: PptxConnector }> = [];
  const walk = (items: PptxElement[], group?: { id: string; parent?: string }) => {
    for (const item of items) {
      if (item.hidden) {
        tally.hidden++;
        continue;
      }
      if (item.kind === 'group') {
        const id = generateId();
        walk(item.children, { id, ...(group ? { parent: group.id } : {}) });
        continue;
      }
      let el: SlideElement | undefined;
      switch (item.kind) {
        case 'shape': el = shapeFromPptx(item, containerId, tally); break;
        case 'picture': el = pictureFromPptx(item, containerId, tally); break;
        case 'table': el = tableFromPptx(item, containerId, tally); break;
        case 'opaque': el = opaqueFromPptx(item, containerId); break;
        case 'connector': {
          const line = connectorFromPptx(item, containerId);
          connectors.push({ el: line, source: item });
          el = line;
          break;
        }
      }
      if (!el) continue;
      if (group) {
        el.groupId = group.id;
        if (group.parent) el.parentGroupId = group.parent;
      }
      if (item.pptxId !== '' && !byPptxId.has(item.pptxId)) byPptxId.set(item.pptxId, el);
      out.push(el);
    }
  };
  walk(list);
  bindConnectors(connectors, byPptxId);
  const keys = indexesBetween(null, null, out.length);
  out.forEach((el, i) => {
    el.index = keys[i];
  });
  return out;
}

const NAMES: Record<SlideElement['type'], string> = { text: 'Text', shape: 'Shape', image: 'Picture', line: 'Connector', table: 'Table', opaque: 'Content' };

/** A source of ids for the elements that have no `pptxId` (and for groups), unique among the slide's. */
function idSource(elements: SlideElement[]): () => string {
  const used = new Set(elements.flatMap((e) => (e.pptxId !== undefined ? [e.pptxId] : [])));
  let next = 1;
  return () => {
    do next++;
    while (used.has(String(next)));
    used.add(String(next));
    return String(next);
  };
}

/**
 * The interchange elements of a slide, and the `pptxId` written for each model element: the model's elements in z-order, each with a `pptxId` (kept, else a free one), connectors resolved
 * against the shapes they are glued to, and groups rebuilt from `groupId` / `parentGroupId` with their box from their members.
 */
export function elementsToPptx(elements: SlideElement[]): { elements: PptxElement[]; pptxIds: Map<string, string> } {
  const sorted = [...elements].sort((a, b) => Number(a.index > b.index) - Number(a.index < b.index));
  const withIds = sorted.map((e) => (e.pptxId === undefined ? { ...e } : e)) as SlideElement[];
  const freshId = idSource(withIds);
  for (const e of withIds) e.pptxId ??= freshId();
  const byId = new Map(withIds.map((e) => [e.id, e]));
  const counts: Record<string, number> = {};
  const nameOf = (e: SlideElement): string => `${NAMES[e.type]} ${(counts[e.type] = (counts[e.type] ?? 0) + 1)}`;
  const convert = (e: SlideElement): PptxElement => {
    const name = nameOf(e);
    switch (e.type) {
      case 'text': return textElementToPptx(e, name);
      case 'shape': return shapeElementToPptx(e, name);
      case 'image': return imageToPptx(e, name);
      case 'table': return tableToPptx(e, name);
      case 'line': return lineToPptx(e, name, byId);
      case 'opaque': return opaqueToPptx(e, name);
    }
  };

  const top: PptxElement[] = [];
  const parents = groupParents(withIds);
  const nodes = new Map<string, PptxGroup>();
  const group = (id: string): PptxGroup => {
    let node = nodes.get(id);
    if (!node) {
      node = { kind: 'group', pptxId: freshId(), name: `Group ${nodes.size + 1}`, children: [] };
      nodes.set(id, node);
      const parent = parents.get(id);
      (parent && parent !== id ? group(parent).children : top).push(node);
    }
    return node;
  };
  const frames = new Map<string, SlideElement[]>();
  for (const e of withIds) {
    const target = e.groupId ? group(e.groupId).children : top;
    target.push(convert(e));
    for (const g of groupChain(e.groupId, parents)) frames.set(g, [...(frames.get(g) ?? []), e]);
  }
  for (const [id, node] of nodes) {
    const box = unionAABB((frames.get(id) as SlideElement[]).map((m) => ({ ...frameAABB(m), rotation: 0 })));
    if (box) node.frame = { x: box.x, y: box.y, w: box.w, h: box.h };
  }
  return { elements: top, pptxIds: new Map(withIds.map((e) => [e.id, e.pptxId as string])) };
}
