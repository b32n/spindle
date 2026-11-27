import type { PptxColorMap, PptxFill, PptxLayout, PptxLevelStyles, PptxLine, PptxShape } from '@b32nio/spindle-shared/export/pptx';
import { isTrue, kid, type XmlElement } from '../xlsx/xml';
import { colorElement, parseColor } from './color';
import { parseFill, parseGeometry, parseLine, type Rels } from './fill';
import { frameOf, matchLayoutPlaceholder } from './levels';
import { shapeLink } from './links';
import { readTextBody } from './text';
import { EMU_PER_PX, num } from './units';

// `p:sp` shapes (plan K9c): frame, geometry, fill and outline, placeholder inheritance and the text body.

export interface ShapeContext {
  map: PptxColorMap;
  rels: Rels;
  /** What a paragraph of a non-placeholder shape starts from: the presentation's default text style. */
  defaultLevels: PptxLevelStyles;
  /** The layout the shapes sit on; its placeholders give position, text style and body properties to the slide's placeholders. */
  layout?: PptxLayout;
  warn: (message: string) => void;
}

/** Outline widths of the Office theme's three line styles (`lnRef idx` 1 to 3: 6350, 12700 and 19050 EMU, that is 0.5, 1 and 1.5 pt), in px. */
const STYLE_LINE_PX = [0, 6350, 12700, 19050].map((emu) => emu / EMU_PER_PX);

const hasPlaceholder = (e: XmlElement): boolean => !!kid(kid(kid(e, 'nvSpPr'), 'nvPr'), 'ph');

/** `p:ph` as the model keeps it: its type (`obj` when absent) and `idx`. */
export function readPlaceholder(ph: XmlElement | undefined): { type: string; idx?: number } | undefined {
  return ph ? { type: ph.attrs.type || 'obj', ...(num(ph.attrs.idx) !== undefined ? { idx: num(ph.attrs.idx) as number } : {}) } : undefined;
}

/** The children of a shape tree, those inside `mc:AlternateContent` (its `Choice`, else its `Fallback`) included, in drawing order. */
export function treeChildren(tree: XmlElement | undefined): XmlElement[] {
  const out: XmlElement[] = [];
  for (const c of tree?.children ?? []) {
    if (c.name === 'AlternateContent' && c.ns === 'mc') {
      const branch = c.children.find((b) => b.name === 'Choice') ?? c.children.find((b) => b.name === 'Fallback');
      out.push(...treeChildren(branch));
    } else out.push(c);
  }
  return out;
}

/** The `p:sp` shapes of a shape tree (`p:spTree`). With `placeholders` false (a layout's or master's static content) the placeholders are left out. */
export function readShapes(tree: XmlElement | undefined, ctx: ShapeContext, placeholders: boolean): PptxShape[] {
  return treeChildren(tree)
    .filter((e) => e.name === 'sp')
    .filter((e) => placeholders || !hasPlaceholder(e))
    .map((e) => readShape(e, ctx));
}

/** A `p:style` reference (`fillRef`, `lnRef`): its index and its colour (`phClr` and unreadable colours give none). */
function styleRef(style: XmlElement | undefined, name: string, map: PptxColorMap): { idx: number; color?: NonNullable<ReturnType<typeof parseColor>> } | undefined {
  const ref = kid(style, name);
  const idx = num(ref?.attrs.idx);
  if (!ref || idx === undefined) return undefined;
  const color = parseColor(colorElement(ref), map);
  return { idx, ...(color ? { color } : {}) };
}

export function readShape(e: XmlElement, ctx: ShapeContext): PptxShape {
  const nv = kid(e, 'nvSpPr');
  const cNvPr = kid(nv, 'cNvPr');
  const ph = kid(kid(nv, 'nvPr'), 'ph');
  const spPr = kid(e, 'spPr');
  const style = kid(e, 'style');
  const placeholder = readPlaceholder(ph);
  const match = placeholder ? matchLayoutPlaceholder(ctx.layout, placeholder) : undefined;

  const frame = frameOf(kid(spPr, 'xfrm')) ?? match?.frame;
  const geometry = parseGeometry(spPr);
  // a `grpFill` reads as no fill here, so a style `fillRef` applies; PowerPoint takes the parent group's fill instead (groups are read in K10, which must resolve it)
  let fill = parseFill(spPr, ctx.map, ctx.rels);
  const fillRef = styleRef(style, 'fillRef', ctx.map);
  if (!fill && fillRef) fill = fillRef.idx === 0 || !fillRef.color ? undefined : styleFill(fillRef.color, fillRef.idx, ctx);
  let line = parseLine(kid(spPr, 'ln'), ctx.map, ctx.rels);
  const lnRef = styleRef(style, 'lnRef', ctx.map);
  if (lnRef && lnRef.idx > 0 && lnRef.color && line?.fill === undefined) {
    const idx = Math.min(3, lnRef.idx);
    line = { width: STYLE_LINE_PX[idx], ...line, fill: { kind: 'solid', color: lnRef.color } } satisfies PptxLine;
  }
  const body = kid(e, 'txBody');
  // `fontRef` names a font (`idx="minor"`), not a style index, so only its colour is read
  const fontColor = parseColor(colorElement(kid(style, 'fontRef')), ctx.map);
  const name = (cNvPr?.attrs.name ?? '').trim();
  const altText = (cNvPr?.attrs.descr ?? '').trim();
  const links = { rels: ctx.rels, warn: ctx.warn };
  const link = shapeLink(cNvPr, links);
  return {
    kind: 'shape',
    pptxId: cNvPr?.attrs.id ?? '',
    name,
    ...(altText !== '' ? { altText } : {}),
    ...(isTrue(cNvPr?.attrs.hidden) ? { hidden: true } : {}),
    ...(placeholder ? { placeholder } : {}),
    ...(isTrue(kid(nv, 'cNvSpPr')?.attrs.txBox) ? { textBox: true } : {}),
    ...(frame ? { frame } : {}),
    ...(geometry ? { geometry } : {}),
    ...(fill ? { fill } : {}),
    ...(line ? { line } : {}),
    ...(body ? { text: readTextBody(body, match?.levels ?? ctx.defaultLevels, match?.bodyProps, ctx.map, fontColor, links) } : {}),
    ...(link ? { link } : {}),
  };
}

/** A shape style's fill: the theme's fill styles are not read, so a solid style is its colour and a gradient style (index 2 and up in the Office theme) is approximated by that colour too, with a warning. */
function styleFill(color: NonNullable<ReturnType<typeof parseColor>>, idx: number, ctx: ShapeContext): PptxFill {
  if (idx >= 2) ctx.warn('A shape takes a gradient fill from the theme\'s style list; it is read as a solid colour.');
  return { kind: 'solid', color };
}
