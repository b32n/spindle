// SVG for custom geometry (`a:custGeom` paths, K6c). Each path is scaled from its own coordinate space onto the box.

import { arcEndPoint, type CustomPath, type PathFillMode } from '@b32nio/spindle-slides-core';

const round = (n: number): number => Math.round(n * 100) / 100;

/** The SVG path data of one custom path scaled to a `w` x `h` box. Arcs become SVG arcs; a full turn is drawn as two halves. */
export function customPathD(path: CustomPath, w: number, h: number): string {
  const sx = w / path.w;
  const sy = h / path.h;
  const X = (v: number): number => round(v * sx);
  const Y = (v: number): number => round(v * sy);
  const out: string[] = [];
  let cur = { x: 0, y: 0 };
  let start = { x: 0, y: 0 };
  for (const c of path.commands) {
    switch (c.op) {
      case 'move':
        cur = start = { x: c.x, y: c.y };
        out.push(`M${X(c.x)} ${Y(c.y)}`);
        break;
      case 'line':
        cur = { x: c.x, y: c.y };
        out.push(`L${X(c.x)} ${Y(c.y)}`);
        break;
      case 'cubic':
        cur = { x: c.x, y: c.y };
        out.push(`C${X(c.x1)} ${Y(c.y1)} ${X(c.x2)} ${Y(c.y2)} ${X(c.x)} ${Y(c.y)}`);
        break;
      case 'quad':
        cur = { x: c.x, y: c.y };
        out.push(`Q${X(c.x1)} ${Y(c.y1)} ${X(c.x)} ${Y(c.y)}`);
        break;
      case 'arc': {
        const rx = Math.abs(c.wR);
        const ry = Math.abs(c.hR);
        if (rx === 0 || ry === 0 || c.swAng === 0) break;
        const sweep = c.swAng > 0 ? 1 : 0;
        // more than a turn only retraces the circle: clamp to one turn. A sweep of a full turn starts and ends at one
        // point, which SVG cannot draw as one arc: go via the half way
        const total = Math.sign(c.swAng) * Math.min(Math.abs(c.swAng), 360);
        const turns = Math.abs(total) >= 360 ? [total / 2, total / 2] : [total];
        let from = c.stAng;
        for (const sw of turns) {
          const end = arcEndPoint(cur.x, cur.y, rx, ry, from, sw);
          out.push(`A${round(rx * sx)} ${round(ry * sy)} 0 ${Math.abs(sw) > 180 ? 1 : 0} ${sweep} ${X(end.x)} ${Y(end.y)}`);
          cur = end;
          from += sw;
        }
        break;
      }
      case 'close':
        cur = start;
        out.push('Z');
        break;
    }
  }
  return out.join(' ');
}

/** The tint laid over a shaded fill mode: the colour and its opacity (`lighten` and `darken` are twice as strong as the `Less` ones). */
export function pathShade(mode: PathFillMode | undefined): { color: string; opacity: number } | undefined {
  switch (mode) {
    case 'lighten': return { color: '#ffffff', opacity: 0.4 };
    case 'lightenLess': return { color: '#ffffff', opacity: 0.2 };
    case 'darken': return { color: '#000000', opacity: 0.4 };
    case 'darkenLess': return { color: '#000000', opacity: 0.2 };
    default: return undefined;
  }
}
