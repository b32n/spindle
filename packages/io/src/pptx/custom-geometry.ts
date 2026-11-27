import type { PptxCustomPath, PptxPathCommand } from '@b32nio/spindle-shared/export/pptx';
import { isTrue, kid, kids, type XmlElement } from '../xlsx/xml';

// Custom geometry (`a:custGeom`, plan K10b): the paths of the shape, moveTo, lnTo, cubicBezTo, quadBezTo, arcTo and close, in the
// path's own coordinate space. Guide formulas (`a:gdLst`) are not evaluated: a coordinate may be a number or one of the
// built-in guides that the box fixes (`l t r b w h hc vc wd2 hd2 ...`); anything else makes the geometry unreadable and the
// shape keeps its frame, fill and text without its outline. This is the same reading as slides-core's `parseCustGeom`
// (the same cases are tested on both sides), kept here because the io package does not depend on slides-core.

/** Limits that keep a hostile file from building an unbounded path. */
export const MAX_GEOMETRY_PATHS = 64;
export const MAX_PATH_COMMANDS = 20000;

const FILL_MODES: ReadonlySet<string> = new Set(['norm', 'none', 'lighten', 'lightenLess', 'darken', 'darkenLess']);
const ANGLE_UNITS = 60000;

/** The built-in guides (ECMA-376 20.1.9.11) that depend only on the path box. */
function builtinGuide(name: string, w: number, h: number): number | undefined {
  switch (name) {
    case 'l': case 't': return 0;
    case 'r': case 'w': return w;
    case 'b': case 'h': return h;
    case 'hc': case 'wd2': return w / 2;
    case 'vc': case 'hd2': return h / 2;
    case 'wd4': return w / 4;
    case 'hd4': return h / 4;
    case 'wd8': return w / 8;
    case 'hd8': return h / 8;
    case 'wd3': return w / 3;
    case 'hd3': return h / 3;
    case 'wd6': return w / 6;
    case 'hd6': return h / 6;
    default: return undefined;
  }
}

/** A path `w`/`h` attribute: a positive finite integer, else undefined (the box size is used instead). */
function pathSize(v: string | undefined): number | undefined {
  if (v === undefined || !/^\d+$/.test(v)) return undefined;
  const n = Number(v);
  return n > 0 && Number.isFinite(n) ? n : undefined;
}

/** One OOXML integer or built-in guide name; undefined when it is anything else. */
function coord(v: string | undefined, w: number, h: number): number | undefined {
  if (v === undefined) return undefined;
  if (/^-?\d+$/.test(v)) {
    const n = Number(v);
    return Number.isFinite(n) ? n : undefined;
  }
  return builtinGuide(v, w, h);
}

function readCommand(n: XmlElement, w: number, h: number): PptxPathCommand | undefined {
  const pts = kids(n, 'pt').map((p) => [coord(p.attrs.x, w, h), coord(p.attrs.y, w, h)] as const);
  if (pts.some(([x, y]) => x === undefined || y === undefined)) return undefined;
  const p = pts as readonly (readonly [number, number])[];
  switch (n.name) {
    case 'moveTo': case 'lnTo':
      return p.length === 1 ? { op: n.name === 'moveTo' ? 'move' : 'line', x: p[0][0], y: p[0][1] } : undefined;
    case 'cubicBezTo':
      return p.length === 3 ? { op: 'cubic', x1: p[0][0], y1: p[0][1], x2: p[1][0], y2: p[1][1], x: p[2][0], y: p[2][1] } : undefined;
    case 'quadBezTo':
      return p.length === 2 ? { op: 'quad', x1: p[0][0], y1: p[0][1], x: p[1][0], y: p[1][1] } : undefined;
    case 'arcTo': {
      const wR = coord(n.attrs.wR, w, h);
      const hR = coord(n.attrs.hR, w, h);
      const st = coord(n.attrs.stAng, w, h);
      const sw = coord(n.attrs.swAng, w, h);
      return wR === undefined || hR === undefined || st === undefined || sw === undefined
        ? undefined
        : { op: 'arc', wR, hR, stAng: st / ANGLE_UNITS, swAng: sw / ANGLE_UNITS };
    }
    case 'close':
      return { op: 'close' };
    default:
      return undefined;
  }
}

const flag = (v: string | undefined, dflt: boolean): boolean => (v === undefined ? dflt : isTrue(v));

/**
 * The paths of an `a:custGeom`. `box` is the shape's size in EMU: the size of a path that gives no `w` and `h`, and the value
 * of the `w` and `h` guides. `undefined` when the geometry cannot be kept faithfully (a guide formula as a coordinate, a
 * malformed command, no path, too large).
 */
export function readCustomGeometry(geom: XmlElement, box: { w: number; h: number } | undefined): PptxCustomPath[] | undefined {
  const pathNodes = kids(kid(geom, 'pathLst'), 'path');
  if (pathNodes.length === 0 || pathNodes.length > MAX_GEOMETRY_PATHS) return undefined;
  const paths: PptxCustomPath[] = [];
  let budget = MAX_PATH_COMMANDS;
  for (const pn of pathNodes) {
    const w = pathSize(pn.attrs.w) ?? box?.w;
    const h = pathSize(pn.attrs.h) ?? box?.h;
    if (w === undefined || h === undefined || !(w > 0) || !(h > 0)) return undefined;
    budget -= pn.children.length;
    if (budget < 0) return undefined;
    const commands: PptxPathCommand[] = [];
    for (const c of pn.children) {
      const cmd = readCommand(c, w, h);
      if (!cmd) return undefined;
      commands.push(cmd);
    }
    const path: PptxCustomPath = { w, h, commands };
    const fill = pn.attrs.fill;
    if (fill !== undefined) {
      if (!FILL_MODES.has(fill)) return undefined;
      if (fill !== 'norm') path.fill = fill;
    }
    if (!flag(pn.attrs.stroke, true)) path.stroke = false;
    if (flag(pn.attrs.extrusionOk, false)) path.extrusionOk = true;
    paths.push(path);
  }
  return paths;
}
