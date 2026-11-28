import type { PptxCustomPath, PptxPathCommand } from '@b32nio/spindle-shared/export/pptx';
import { attrs } from './writer-xml';

// Custom geometry (`a:custGeom`, plan K13): the paths the reader kept, written back command by command. The inverse of `readCustomGeometry`.

/** 60000ths of a degree, as `arcTo` states its angles. */
const ANGLE_UNITS = 60000;

/** A coordinate as a whole number; a value that is not finite (a model built by hand) is 0. */
const n = (v: number): number => (Number.isFinite(v) ? Math.round(v) : 0);

const pt = (x: number, y: number): string => `<a:pt x="${n(x)}" y="${n(y)}"/>`;

function commandXml(c: PptxPathCommand): string {
  switch (c.op) {
    case 'move':
      return `<a:moveTo>${pt(c.x, c.y)}</a:moveTo>`;
    case 'line':
      return `<a:lnTo>${pt(c.x, c.y)}</a:lnTo>`;
    case 'cubic':
      return `<a:cubicBezTo>${pt(c.x1, c.y1)}${pt(c.x2, c.y2)}${pt(c.x, c.y)}</a:cubicBezTo>`;
    case 'quad':
      return `<a:quadBezTo>${pt(c.x1, c.y1)}${pt(c.x, c.y)}</a:quadBezTo>`;
    case 'arc':
      return `<a:arcTo${attrs({ wR: n(c.wR), hR: n(c.hR), stAng: Math.round(c.stAng * ANGLE_UNITS), swAng: Math.round(c.swAng * ANGLE_UNITS) })}/>`;
    case 'close':
      return '<a:close/>';
  }
}

function pathXml(p: PptxCustomPath): string {
  const a = attrs({ w: Math.max(1, n(p.w)), h: Math.max(1, n(p.h)), fill: p.fill, stroke: p.stroke === false ? 0 : undefined, extrusionOk: p.extrusionOk ? 1 : undefined });
  return `<a:path${a}>${p.commands.map(commandXml).join('')}</a:path>`;
}

/** `a:custGeom` with its paths (the guide, handle and connection-site lists are empty: the reader does not evaluate them), the text rectangle being the whole box. */
export const customGeometryXml = (paths: PptxCustomPath[]): string =>
  `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst>${paths.map(pathXml).join('')}</a:pathLst></a:custGeom>`;
