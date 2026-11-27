// Custom geometry (`a:custGeom`, K6c): a path model that is kept, drawn and written back out. The model types live in
// `types.ts`; this file holds the validation of untrusted data, the DrawingML reading and writing (on a plain element
// tree, so slides-core needs no XML parser: the PPTX reader hands over its parsed nodes), and the arc maths shared by the
// renderer. Guide formulas (`a:gdLst`) are not evaluated: a coordinate may be a number or one of the built-in guides that
// are fixed by the box (`l t r b w h hc vc wd2 hd2 ...`); anything else makes the geometry unreadable and the caller keeps
// the shape as an `unknown:custGeom` passthrough instead.

import type { CustomGeometry, CustomPath, PathCommand, PathFillMode } from './types';

/** The part of a parsed XML element this reader needs (the io package's `XmlElement` satisfies it): local name, attributes by local name. */
export interface GeomNode {
  name: string;
  attrs: Readonly<Record<string, string>>;
  children: readonly GeomNode[];
}

/** Limits that keep a hostile file or room from building an unbounded path. */
export const MAX_GEOMETRY_PATHS = 64;
export const MAX_PATH_COMMANDS = 20000;

const FILL_MODES: ReadonlySet<string> = new Set(['norm', 'none', 'lighten', 'lightenLess', 'darken', 'darkenLess']);
const ANGLE_UNITS = 60000;

const isNum = (v: unknown): v is number => typeof v === 'number' && Number.isFinite(v);
const isObj = (v: unknown): v is Record<string, unknown> => typeof v === 'object' && v !== null && !Array.isArray(v);

/** True when every named field of `c` is a finite number. */
function allNums(c: Record<string, unknown>, keys: readonly string[]): boolean {
  return keys.every((k) => isNum(c[k]));
}

function normalizeCommand(c: unknown): PathCommand | undefined {
  if (!isObj(c)) return undefined;
  switch (c.op) {
    case 'move':
    case 'line':
      return allNums(c, ['x', 'y']) ? { op: c.op, x: c.x as number, y: c.y as number } : undefined;
    case 'cubic':
      return allNums(c, ['x1', 'y1', 'x2', 'y2', 'x', 'y'])
        ? { op: 'cubic', x1: c.x1 as number, y1: c.y1 as number, x2: c.x2 as number, y2: c.y2 as number, x: c.x as number, y: c.y as number }
        : undefined;
    case 'quad':
      return allNums(c, ['x1', 'y1', 'x', 'y']) ? { op: 'quad', x1: c.x1 as number, y1: c.y1 as number, x: c.x as number, y: c.y as number } : undefined;
    case 'arc':
      return allNums(c, ['wR', 'hR', 'stAng', 'swAng'])
        ? { op: 'arc', wR: c.wR as number, hR: c.hR as number, stAng: c.stAng as number, swAng: c.swAng as number }
        : undefined;
    case 'close':
      return { op: 'close' };
    default:
      return undefined;
  }
}

/**
 * Validate custom geometry that came from outside the type system (a collab room, hand-edited data): a path with a
 * non-positive or non-finite size is dropped, so is a command that is malformed, and the totals are capped. Returns
 * undefined when no path survives.
 */
export function normalizeCustomGeometry(value: unknown): CustomGeometry | undefined {
  if (!isObj(value) || !Array.isArray(value.paths)) return undefined;
  const paths: CustomPath[] = [];
  let budget = MAX_PATH_COMMANDS;
  for (const raw of value.paths.slice(0, MAX_GEOMETRY_PATHS)) {
    if (!isObj(raw) || !isNum(raw.w) || !isNum(raw.h) || raw.w <= 0 || raw.h <= 0 || !Array.isArray(raw.commands)) continue;
    const commands: PathCommand[] = [];
    for (const c of raw.commands.slice(0, budget)) {
      const cmd = normalizeCommand(c);
      if (cmd) commands.push(cmd);
    }
    budget -= commands.length;
    const path: CustomPath = { w: raw.w, h: raw.h, commands };
    if (typeof raw.fill === 'string' && FILL_MODES.has(raw.fill) && raw.fill !== 'norm') path.fill = raw.fill as PathFillMode;
    if (raw.stroke === false) path.stroke = false;
    if (raw.extrusionOk === true) path.extrusionOk = true;
    paths.push(path);
  }
  return paths.length > 0 ? { paths } : undefined;
}

// ── DrawingML reading ───────────────────────────────────────────────────────

const named = (n: GeomNode, name: string): GeomNode[] => n.children.filter((c) => c.name === name);

/**
 * The built-in guides (ECMA-376 20.1.9.11) that depend only on the path box. They are evaluated against the path's own
 * `w` x `h` (ECMA-376 gives shape-box values): the two agree when the path has no size of its own, the normal case; in a
 * path that has one, `r`, `b`, `hc` and the rest mean that path's edges and centre.
 */
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
    case 'ss': return Math.min(w, h);
    case 'ls': return Math.max(w, h);
    case 'cd2': return 10800000;
    case 'cd4': return 5400000;
    case 'cd8': return 2700000;
    case '3cd4': return 16200000;
    case '3cd8': return 8100000;
    case '5cd8': return 13500000;
    case '7cd8': return 18900000;
    default: return undefined;
  }
}

/** A path `w`/`h` attribute: a positive finite integer, else undefined (the box size is used instead). */
function pathSize(v: string | undefined): number | undefined {
  if (v === undefined || !/^\d+$/.test(v)) return undefined;
  const n = Number(v);
  return n > 0 && Number.isFinite(n) ? n : undefined;
}

/** Reads one OOXML integer or built-in guide name; undefined when it is anything else. */
function coord(v: string | undefined, w: number, h: number): number | undefined {
  if (v === undefined) return undefined;
  if (/^-?\d+$/.test(v)) {
    const n = Number(v);
    return Number.isFinite(n) ? n : undefined;
  }
  return builtinGuide(v, w, h);
}

function readPoint(n: GeomNode, w: number, h: number): [number, number] | undefined {
  const x = coord(n.attrs.x, w, h);
  const y = coord(n.attrs.y, w, h);
  return x === undefined || y === undefined ? undefined : [x, y];
}

function readCommand(n: GeomNode, w: number, h: number): PathCommand | undefined {
  const pts = named(n, 'pt').map((p) => readPoint(p, w, h));
  if (pts.some((p) => p === undefined)) return undefined;
  const p = pts as [number, number][];
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

const flag = (v: string | undefined, dflt: boolean): boolean => (v === undefined ? dflt : v === '1' || v === 'true');

/**
 * Read an `a:custGeom` element. `box` is the shape's size in the coordinate unit of the file (EMU): the default size of
 * a path that gives no `w` and `h`, and the value of the `w` and `h` guides. Returns undefined when the geometry cannot
 * be kept faithfully (a guide formula as a coordinate, a malformed command, no path, too large): the caller then keeps it
 * as an `unknown:custGeom` passthrough and reports it.
 */
export function parseCustGeom(node: GeomNode, box?: { w: number; h: number }): CustomGeometry | undefined {
  const list = named(node, 'pathLst')[0];
  if (!list) return undefined;
  const pathNodes = named(list, 'path');
  if (pathNodes.length === 0 || pathNodes.length > MAX_GEOMETRY_PATHS) return undefined;
  const paths: CustomPath[] = [];
  let budget = MAX_PATH_COMMANDS;
  for (const pn of pathNodes) {
    const w = pathSize(pn.attrs.w) ?? box?.w;
    const h = pathSize(pn.attrs.h) ?? box?.h;
    if (w === undefined || h === undefined || !(w > 0) || !(h > 0) || !Number.isFinite(w) || !Number.isFinite(h)) return undefined;
    budget -= pn.children.length;
    if (budget < 0) return undefined;
    const commands: PathCommand[] = [];
    for (const c of pn.children) {
      const cmd = readCommand(c, w, h);
      if (!cmd) return undefined;
      commands.push(cmd);
    }
    const path: CustomPath = { w, h, commands };
    const fill = pn.attrs.fill;
    if (fill !== undefined) {
      if (!FILL_MODES.has(fill)) return undefined;
      if (fill !== 'norm') path.fill = fill as PathFillMode;
    }
    if (!flag(pn.attrs.stroke, true)) path.stroke = false;
    if (flag(pn.attrs.extrusionOk, false)) path.extrusionOk = true;
    paths.push(path);
  }
  return { paths };
}

// ── DrawingML writing ───────────────────────────────────────────────────────

const int = (n: number): string => String(Math.round(n));
const pt = (x: number, y: number): string => `<a:pt x="${int(x)}" y="${int(y)}"/>`;

function writeCommand(c: PathCommand): string {
  switch (c.op) {
    case 'move': return `<a:moveTo>${pt(c.x, c.y)}</a:moveTo>`;
    case 'line': return `<a:lnTo>${pt(c.x, c.y)}</a:lnTo>`;
    case 'cubic': return `<a:cubicBezTo>${pt(c.x1, c.y1)}${pt(c.x2, c.y2)}${pt(c.x, c.y)}</a:cubicBezTo>`;
    case 'quad': return `<a:quadBezTo>${pt(c.x1, c.y1)}${pt(c.x, c.y)}</a:quadBezTo>`;
    case 'arc':
      return `<a:arcTo wR="${int(c.wR)}" hR="${int(c.hR)}" stAng="${int(c.stAng * ANGLE_UNITS)}" swAng="${int(c.swAng * ANGLE_UNITS)}"/>`;
    case 'close': return '<a:close/>';
  }
}

/**
 * Write geometry as an `a:custGeom` element (namespace prefix `a:`, no declaration). Coordinates are rounded to whole
 * units as the schema requires. The guide, handle and connection-site lists are written empty and the text rectangle is
 * the whole box, since the model keeps none of them.
 */
export function serializeCustGeom(geom: CustomGeometry): string {
  const paths = geom.paths.map((p) => {
    const attrs =
      `w="${int(p.w)}" h="${int(p.h)}"` +
      (p.fill && p.fill !== 'norm' ? ` fill="${p.fill}"` : '') +
      (p.stroke === false ? ' stroke="0"' : '') +
      (p.extrusionOk ? ' extrusionOk="1"' : '');
    return `<a:path ${attrs}>${p.commands.map(writeCommand).join('')}</a:path>`;
  });
  return (
    '<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/>' +
    `<a:pathLst>${paths.join('')}</a:pathLst></a:custGeom>`
  );
}

// ── Drawing helpers ─────────────────────────────────────────────────────────

/**
 * The end point of an `arcTo` that starts at `(x, y)`: the ellipse of radii `wR` x `hR` is placed so the start point
 * sits at angle `stAng` on it (angles in degrees, clockwise from +x, measured on the ellipse as OOXML does), and the
 * arc sweeps `swAng` degrees from there.
 */
export function arcEndPoint(x: number, y: number, wR: number, hR: number, stAng: number, swAng: number): { x: number; y: number } {
  const rad = (d: number): number => (d * Math.PI) / 180;
  const st = rad(stAng);
  const end = rad(stAng + swAng);
  // OOXML angles are visual: the parametric angle of the ellipse point at visual angle a is atan2(wR sin a, hR cos a)
  const param = (a: number): number => Math.atan2(wR * Math.sin(a), hR * Math.cos(a));
  const t0 = param(st);
  const t1 = param(end);
  const cx = x - wR * Math.cos(t0);
  const cy = y - hR * Math.sin(t0);
  return { x: cx + wR * Math.cos(t1), y: cy + hR * Math.sin(t1) };
}
