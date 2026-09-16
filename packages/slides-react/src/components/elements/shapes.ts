// Pure SVG geometry for the built-in shape presets. Each generator maps a
// preset + box size (+ optional adjustment handles) to an SVG description the
// ShapeView renders inside a `viewBox="0 0 w h"`. Kept framework-free so it can
// be unit-tested and reused by the PDF/print path.

import type { ShapePreset } from '@b32nio/spindle-slides-core';

export type ShapeGeom =
  | { type: 'path'; d: string }
  | { type: 'ellipse'; cx: number; cy: number; rx: number; ry: number };

type Point = [number, number];

function polyPath(points: Point[]): string {
  return points.map(([x, y], i) => `${i === 0 ? 'M' : 'L'}${round(x)} ${round(y)}`).join(' ') + ' Z';
}

function round(n: number): number {
  return Math.round(n * 100) / 100;
}

/** Regular polygon inscribed in the box, first vertex pointing up. */
function regularPolygon(sides: number, w: number, h: number): Point[] {
  const cx = w / 2;
  const cy = h / 2;
  const pts: Point[] = [];
  for (let i = 0; i < sides; i++) {
    const angle = -Math.PI / 2 + (i * 2 * Math.PI) / sides;
    pts.push([cx + (w / 2) * Math.cos(angle), cy + (h / 2) * Math.sin(angle)]);
  }
  return pts;
}

function star(points: number, w: number, h: number, innerRatio: number): Point[] {
  const cx = w / 2;
  const cy = h / 2;
  const pts: Point[] = [];
  for (let i = 0; i < points * 2; i++) {
    const angle = -Math.PI / 2 + (i * Math.PI) / points;
    const rx = (w / 2) * (i % 2 === 0 ? 1 : innerRatio);
    const ry = (h / 2) * (i % 2 === 0 ? 1 : innerRatio);
    pts.push([cx + rx * Math.cos(angle), cy + ry * Math.sin(angle)]);
  }
  return pts;
}

export function shapeGeom(preset: ShapePreset, w: number, h: number, adjustments?: number[]): ShapeGeom {
  const adj = adjustments?.[0];
  switch (preset) {
    case 'rect':
      return { type: 'path', d: polyPath([[0, 0], [w, 0], [w, h], [0, h]]) };

    case 'roundRect': {
      const r = Math.min(w, h) * (adj ?? 0.15);
      const d =
        `M${round(r)} 0 H${round(w - r)} A${round(r)} ${round(r)} 0 0 1 ${round(w)} ${round(r)} ` +
        `V${round(h - r)} A${round(r)} ${round(r)} 0 0 1 ${round(w - r)} ${round(h)} ` +
        `H${round(r)} A${round(r)} ${round(r)} 0 0 1 0 ${round(h - r)} ` +
        `V${round(r)} A${round(r)} ${round(r)} 0 0 1 ${round(r)} 0 Z`;
      return { type: 'path', d };
    }

    case 'ellipse':
      return { type: 'ellipse', cx: w / 2, cy: h / 2, rx: w / 2, ry: h / 2 };

    case 'triangle':
      return { type: 'path', d: polyPath([[w / 2, 0], [w, h], [0, h]]) };

    case 'rightTriangle':
      return { type: 'path', d: polyPath([[0, 0], [0, h], [w, h]]) };

    case 'diamond':
      return { type: 'path', d: polyPath([[w / 2, 0], [w, h / 2], [w / 2, h], [0, h / 2]]) };

    case 'pentagon':
      return { type: 'path', d: polyPath(regularPolygon(5, w, h)) };

    case 'hexagon':
      return {
        type: 'path',
        d: polyPath([[0.25 * w, 0], [0.75 * w, 0], [w, 0.5 * h], [0.75 * w, h], [0.25 * w, h], [0, 0.5 * h]]),
      };

    case 'octagon': {
      const f = 0.29;
      return {
        type: 'path',
        d: polyPath([
          [f * w, 0], [(1 - f) * w, 0], [w, f * h], [w, (1 - f) * h],
          [(1 - f) * w, h], [f * w, h], [0, (1 - f) * h], [0, f * h],
        ]),
      };
    }

    case 'star5':
      return { type: 'path', d: polyPath(star(5, w, h, 0.4)) };

    case 'arrowRight': {
      const bodyTop = 0.3 * h;
      const bodyBot = 0.7 * h;
      const neck = 0.6 * w;
      return {
        type: 'path',
        d: polyPath([
          [0, bodyTop], [neck, bodyTop], [neck, 0], [w, h / 2],
          [neck, h], [neck, bodyBot], [0, bodyBot],
        ]),
      };
    }

    case 'arrowLeft': {
      const bodyTop = 0.3 * h;
      const bodyBot = 0.7 * h;
      const neck = 0.4 * w;
      return {
        type: 'path',
        d: polyPath([
          [w, bodyTop], [neck, bodyTop], [neck, 0], [0, h / 2],
          [neck, h], [neck, bodyBot], [w, bodyBot],
        ]),
      };
    }

    case 'chevron':
      return {
        type: 'path',
        d: polyPath([[0, 0], [0.7 * w, 0], [w, h / 2], [0.7 * w, h], [0, h], [0.3 * w, h / 2]]),
      };

    case 'parallelogram':
      return { type: 'path', d: polyPath([[0.25 * w, 0], [w, 0], [0.75 * w, h], [0, h]]) };

    case 'trapezoid':
      return { type: 'path', d: polyPath([[0.25 * w, 0], [0.75 * w, 0], [w, h], [0, h]]) };

    case 'plus':
      return {
        type: 'path',
        d: polyPath([
          [0.33 * w, 0], [0.67 * w, 0], [0.67 * w, 0.33 * h], [w, 0.33 * h],
          [w, 0.67 * h], [0.67 * w, 0.67 * h], [0.67 * w, h], [0.33 * w, h],
          [0.33 * w, 0.67 * h], [0, 0.67 * h], [0, 0.33 * h], [0.33 * w, 0.33 * h],
        ]),
      };

    case 'heart': {
      const p = (x: number, y: number) => `${round(x * w)} ${round(y * h)}`;
      const d =
        `M${p(0.5, 0.28)} C ${p(0.5, 0.1)}, ${p(0.2, 0.05)}, ${p(0.1, 0.25)} ` +
        `C ${p(0, 0.42)}, ${p(0.15, 0.6)}, ${p(0.5, 0.9)} ` +
        `C ${p(0.85, 0.6)}, ${p(1, 0.42)}, ${p(0.9, 0.25)} ` +
        `C ${p(0.8, 0.05)}, ${p(0.5, 0.1)}, ${p(0.5, 0.28)} Z`;
      return { type: 'path', d };
    }

    case 'cloud': {
      const p = (x: number, y: number) => `${round(x * w)} ${round(y * h)}`;
      const d =
        `M${p(0.25, 0.78)} A ${round(0.18 * w)} ${round(0.18 * h)} 0 0 1 ${p(0.2, 0.45)} ` +
        `A ${round(0.2 * w)} ${round(0.2 * h)} 0 0 1 ${p(0.5, 0.32)} ` +
        `A ${round(0.2 * w)} ${round(0.2 * h)} 0 0 1 ${p(0.8, 0.45)} ` +
        `A ${round(0.18 * w)} ${round(0.18 * h)} 0 0 1 ${p(0.75, 0.78)} Z`;
      return { type: 'path', d };
    }

    default:
      return { type: 'path', d: polyPath([[0, 0], [w, 0], [w, h], [0, h]]) };
  }
}
