import type { PptxElement, PptxFill, PptxGeometry, PptxGroup, PptxLine, PptxLineEnd, PptxShadow, PptxShape } from '@b32nio/spindle-shared/export/pptx';
import { hlinkXml, type WriteContext } from './writer-context';
import { txBodyXml } from './writer-body';
import { customGeometryXml } from './writer-custom-geometry';
import { angle, attrs, colorXml, pct, solidFillXml, xfrmXml } from './writer-xml';

// `p:sp` shapes and text boxes (plan K12b, K1, K5): geometry, fill, outline, placeholder and hyperlink. The inverse of `readShape`.

const MAX_ID = 4294967295;
const PRESET_NAME = /^[A-Za-z][A-Za-z0-9]*$/;
const SHADOW_ANCHORS = new Set(['tl', 't', 'tr', 'l', 'ctr', 'r', 'bl', 'b', 'br']);
const CAPS = { flat: 'flat', round: 'rnd', square: 'sq' } as const;
const GRADIENT_PATH = { radial: 'circle', rect: 'rect', shape: 'shape' } as const;
const FILL_TO_RECT = '<a:fillToRect l="50000" t="50000" r="50000" b="50000"/>';

/**
 * The `p:cNvPr` ids of a list of elements (groups' children included): an element keeps its `pptxId` when it is a usable number no
 * earlier element has (1 is the shape tree's own, and `reserved` ids are taken), else it gets the next free one.
 * Elements that are not written yet are numbered too, so a later step finds the ids settled.
 */
export function allocateElementIds(elements: PptxElement[], reserved = 1): Map<PptxElement, number> {
  const all: PptxElement[] = [];
  const walk = (list: PptxElement[]) => {
    for (const e of list) {
      all.push(e);
      if (e.kind === 'group') walk((e as PptxGroup).children);
    }
  };
  walk(elements);
  const taken = new Set<number>();
  for (let i = 1; i <= reserved; i++) taken.add(i);
  const out = new Map<PptxElement, number>();
  const pending: PptxElement[] = [];
  for (const e of all) {
    const n = /^\d+$/.test(e.pptxId) ? Number(e.pptxId) : NaN;
    if (Number.isSafeInteger(n) && n >= 1 && n <= MAX_ID && !taken.has(n)) {
      taken.add(n);
      out.set(e, n);
    } else pending.push(e);
  }
  let next = reserved;
  for (const e of pending) {
    do next++;
    while (taken.has(next));
    taken.add(next);
    out.set(e, next);
  }
  return out;
}

/** A fill as the element that holds it states it; pictures need a media part (K13), so they are left out. */
export function fillXml(fill: PptxFill | undefined, ctx: WriteContext): string {
  switch (fill?.kind) {
    case undefined:
      return '';
    case 'none':
      return '<a:noFill/>';
    case 'solid':
      return solidFillXml(fill.color);
    case 'gradient': {
      const stops = fill.stops.map((s) => `<a:gs pos="${pct(s.pos)}">${colorXml(s.color)}</a:gs>`).join('');
      const direction = fill.shape === 'linear' ? `<a:lin ang="${angle(fill.angle ?? 0)}" scaled="${fill.scaled ? 1 : 0}"/>` : `<a:path path="${GRADIENT_PATH[fill.shape]}">${FILL_TO_RECT}</a:path>`;
      return `<a:gradFill${attrs({ rotWithShape: fill.rotWithShape ? 1 : undefined })}><a:gsLst>${stops}</a:gsLst>${direction}</a:gradFill>`;
    }
    case 'pattern':
      return `<a:pattFill${attrs({ prst: fill.prst })}>${fill.fg ? `<a:fgClr>${colorXml(fill.fg)}</a:fgClr>` : ''}${fill.bg ? `<a:bgClr>${colorXml(fill.bg)}</a:bgClr>` : ''}</a:pattFill>`;
    case 'picture':
      ctx.lossy('pictureFill');
      return '';
  }
}

const lineEndXml = (tag: string, end: PptxLineEnd | undefined): string => (end ? `<a:${tag}${attrs({ type: end.type, w: end.w, len: end.len })}/>` : '');

/** `a:ln` (or the cell border `a:lnL` and its kin, `tag`): width, fill, dash, cap, join and the line ends, in the order of `CT_LineProperties`. */
export function lineXml(line: PptxLine | undefined, ctx: WriteContext, tag = 'a:ln'): string {
  if (!line) return '';
  const a = attrs({ w: line.width !== undefined ? Math.round(line.width * 9525) : undefined, cap: line.cap ? CAPS[line.cap] : undefined });
  if (line.dash === 'custom') ctx.lossy('customDash');
  const body =
    fillXml(line.fill, ctx) +
    (line.dash !== undefined && line.dash !== 'custom' ? `<a:prstDash${attrs({ val: line.dash })}/>` : '') +
    (line.join ? `<a:${line.join}/>` : '') +
    lineEndXml('headEnd', line.head) +
    lineEndXml('tailEnd', line.tail);
  return body === '' ? `<${tag}${a}/>` : `<${tag}${a}>${body}</${tag}>`;
}

/** `a:effectLst` with the outer shadow of a shape; nothing without one. Distances are EMU, the direction 60000ths of a degree. */
export function shadowXml(shadow: PptxShadow | undefined): string {
  if (!shadow) return '';
  const a = attrs({
    blurRad: shadow.blur !== undefined ? Math.round(shadow.blur * 9525) : undefined,
    dist: shadow.dist !== undefined ? Math.round(shadow.dist * 9525) : undefined,
    dir: shadow.dir !== undefined ? angle(shadow.dir) : undefined,
    algn: shadow.algn !== undefined && SHADOW_ANCHORS.has(shadow.algn) ? shadow.algn : undefined,
    rotWithShape: shadow.rotWithShape ? 1 : undefined,
  });
  return `<a:effectLst><a:outerShdw${a}>${colorXml(shadow.color)}</a:outerShdw></a:effectLst>`;
}

/** `a:prstGeom` with its adjustments. A name that is not an identifier (the editor's `unknown:` prefix is dropped) is written as a rectangle; custom geometry is written as its paths, or not at all when the reader could not keep them. */
export function geometryXml(geometry: PptxGeometry | undefined, ctx: WriteContext): string {
  if (!geometry) return '';
  if (geometry.kind === 'custom') {
    if (geometry.paths) return customGeometryXml(geometry.paths);
    ctx.lossy('customGeometry');
    return '';
  }
  const name = geometry.prst.replace(/^unknown:/, '');
  if (!PRESET_NAME.test(name)) ctx.lossy('geometryName');
  const gds = Object.entries(geometry.adjustments ?? {})
    .filter(([k, v]) => PRESET_NAME.test(k) && Number.isFinite(v))
    .map(([k, v]) => `<a:gd${attrs({ name: k, fmla: `val ${Math.round(v)}` })}/>`)
    .join('');
  return `<a:prstGeom${attrs({ prst: PRESET_NAME.test(name) ? name : 'rect' })}>${gds === '' ? '<a:avLst/>' : `<a:avLst>${gds}</a:avLst>`}</a:prstGeom>`;
}

/** One `p:sp`. */
export function shapeXml(shape: PptxShape, id: number, ctx: WriteContext): string {
  const ph = shape.placeholder;
  const cNvPr = attrs({ id, name: shape.name, descr: shape.altText, hidden: shape.hidden ? 1 : undefined });
  const link = hlinkXml(shape.link, ctx);
  const cNvSpPr = ph ? '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>' : `<p:cNvSpPr${attrs({ txBox: shape.textBox ? 1 : undefined })}/>`;
  const nvPr = ph ? `<p:nvPr><p:ph${attrs({ type: ph.type === 'obj' ? undefined : ph.type, idx: ph.idx })}/></p:nvPr>` : '<p:nvPr/>';
  const spPr = (shape.frame ? xfrmXml(shape.frame) : '') + geometryXml(shape.geometry, ctx) + fillXml(shape.fill, ctx) + lineXml(shape.line, ctx) + shadowXml(shape.shadow);
  return (
    `<p:sp><p:nvSpPr>${link === '' ? `<p:cNvPr${cNvPr}/>` : `<p:cNvPr${cNvPr}>${link}</p:cNvPr>`}${cNvSpPr}${nvPr}</p:nvSpPr>` +
    `<p:spPr>${spPr}</p:spPr>${shape.text ? `<p:txBody>${txBodyXml(shape.text, ctx)}</p:txBody>` : ''}</p:sp>`
  );
}
