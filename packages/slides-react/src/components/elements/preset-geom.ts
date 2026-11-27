// Hand-built SVG geometry for the K5 presets (everything except the original 18, which stay in shapes.ts). Guide
// semantics follow the OOXML preset definitions with adjustments normalised to fractions (see SHAPE_PRESET_INFO);
// each preset is a function of the box and its resolved guides. Shapes whose OOXML definition has an open outline
// (brackets, braces) return `strokeD` for the stroked line and `d` for the filled area.

import { resolveAdjustments, type ShapePreset } from '@b32nio/spindle-slides-core';

export interface PresetPath {
  d: string;
  /** When set, `d` is only filled and this open path is only stroked. */
  strokeD?: string;
}

type Pt = [number, number];

const r2 = (n: number): number => Math.round(n * 100) / 100;
const pt = ([x, y]: Pt): string => `${r2(x)} ${r2(y)}`;

const poly = (pts: Pt[]): string => 'M' + pts.map(pt).join(' L') + ' Z';
const seg = (a: Pt, b: Pt): string => `M${pt(a)} L${pt(b)}`;
/** An elliptical arc to (x, y); `cw` is the SVG sweep flag (clockwise on screen). */
const arc = (rx: number, ry: number, cw: boolean, x: number, y: number, large = false): string =>
  ` A${r2(rx)} ${r2(ry)} 0 ${large ? 1 : 0} ${cw ? 1 : 0} ${r2(x)} ${r2(y)}`;
const ellipsePath = (cx: number, cy: number, rx: number, ry: number, cw = true): string =>
  `M${r2(cx - rx)} ${r2(cy)}` + arc(rx, ry, cw, cx + rx, cy) + arc(rx, ry, cw, cx - rx, cy) + ' Z';

function ngon(n: number, w: number, h: number, start: number): Pt[] {
  return Array.from({ length: n }, (_, i) => {
    const t = start + (i * 2 * Math.PI) / n;
    return [w / 2 + (w / 2) * Math.cos(t), h / 2 + (h / 2) * Math.sin(t)] as Pt;
  });
}

/** A star with `n` points, first point up; `inner` is the inner radius as a fraction of the outer. */
function star(n: number, w: number, h: number, inner: number, xScale = 1): Pt[] {
  return Array.from({ length: n * 2 }, (_, i) => {
    const t = -Math.PI / 2 + (i * Math.PI) / n;
    const k = i % 2 === 0 ? 1 : inner;
    return [w / 2 + (w / 2) * k * xScale * Math.cos(t), h / 2 + (h / 2) * k * Math.sin(t)] as Pt;
  });
}

/** A rectangle with a per-corner treatment, clockwise from the top-left: 0 square, > 0 round by that radius, < 0 snip by that size. */
function cornered(w: number, h: number, [tl, tr, br, bl]: [number, number, number, number]): string {
  const at = (x: number, y: number): string => ` L${r2(x)} ${r2(y)}`;
  // each corner is entered along one edge and left along the next: square = one point, snip = two, round = point + arc
  const corner = (c: number, enter: Pt, leave: Pt, corner: Pt): string =>
    c === 0 ? at(...corner) : c < 0 ? at(...enter) + at(...leave) : at(...enter) + arc(c, c, true, ...leave);
  const [a, b, c, d] = [Math.abs(tl), Math.abs(tr), Math.abs(br), Math.abs(bl)];
  return (
    `M${r2(a)} 0` +
    corner(tr, [w - b, 0], [w, b], [w, 0]) +
    corner(br, [w, h - c], [w - c, h], [w, h]) +
    corner(bl, [d, h], [0, h - d], [0, h]) +
    corner(tl, [0, a], [a, 0], [0, 0]) +
    ' Z'
  );
}

/**
 * A (rounded) rectangle whose tail runs from one edge to the tip. The tail leaves the edge nearest the tip, over the
 * OOXML span: the 2/12 to 5/12 or 7/12 to 10/12 of the edge, on the tip's side of the centre.
 */
function calloutPath(w: number, h: number, tipX: number, tipY: number, radius: number): string {
  const dx = tipX - w / 2;
  const dy = tipY - h / 2;
  const edge = Math.abs(dx) * h > Math.abs(dy) * w ? (dx < 0 ? 'left' : 'right') : dy < 0 ? 'top' : 'bottom';
  const span = (len: number, pos: number): [number, number] =>
    pos < len / 2 ? [(len * 2) / 12, (len * 5) / 12] : [(len * 7) / 12, (len * 10) / 12];
  const [x0, x1] = span(w, tipX);
  const [y0, y1] = span(h, tipY);
  const r = Math.min(radius, w / 2, h / 2);
  const to = (x: number, y: number): string => ` L${r2(x)} ${r2(y)}`;
  const tip = to(tipX, tipY);
  let d = `M${r2(r)} 0`;
  if (edge === 'top') d += to(x0, 0) + tip + to(x1, 0);
  d += to(w - r, 0) + (r ? arc(r, r, true, w, r) : '');
  if (edge === 'right') d += to(w, y0) + tip + to(w, y1);
  d += to(w, h - r) + (r ? arc(r, r, true, w - r, h) : '');
  if (edge === 'bottom') d += to(x1, h) + tip + to(x0, h);
  d += to(r, h) + (r ? arc(r, r, true, 0, h - r) : '');
  if (edge === 'left') d += to(0, y1) + tip + to(0, y0);
  d += to(0, r) + (r ? arc(r, r, true, r, 0) : '');
  return d + ' Z';
}

/** A point on the box-inscribed ellipse at `turns` of a full circle, clockwise from three o'clock. */
function onEllipse(w: number, h: number, turns: number, shrink = 0): Pt {
  const t = turns * 2 * Math.PI;
  return [w / 2 + (w / 2 - shrink) * Math.cos(t), h / 2 + (h / 2 - shrink) * Math.sin(t)];
}

/** The clockwise sweep from `a` to `b`, in turns, in [0, 1). */
const sweepTurns = (a: number, b: number): number => (((b - a) % 1) + 1) % 1;

/** A curly brace (open): spine at mid-width, curve radius `ry`, the point at height `ym`. `mirror` makes it a right brace. */
function brace(w: number, h: number, ry: number, ym: number, mirror: boolean): string {
  const X = (x: number): number => (mirror ? w - x : x);
  const cw = (b: boolean): boolean => (mirror ? !b : b);
  const rx = w / 2;
  return (
    `M${r2(X(w))} 0` + arc(rx, ry, cw(false), X(w / 2), ry) +
    ` L${r2(X(w / 2))} ${r2(ym - ry)}` + arc(rx, ry, cw(true), X(0), ym) +
    arc(rx, ry, cw(true), X(w / 2), ym + ry) +
    ` L${r2(X(w / 2))} ${r2(h - ry)}` + arc(rx, ry, cw(false), X(w), h)
  );
}

const BOLT: Pt[] = [
  [8472, 0], [12860, 6080], [11050, 6797], [16577, 12007], [14767, 12877], [21600, 21600],
  [10012, 14915], [12222, 13987], [5022, 9705], [7602, 9096], [0, 3163],
];

/** The original 18 presets, drawn by `shapeGeom` itself. */
type OriginalPreset =
  | 'rect' | 'roundRect' | 'ellipse' | 'triangle' | 'rightTriangle' | 'diamond' | 'pentagon' | 'hexagon' | 'octagon'
  | 'star5' | 'arrowRight' | 'arrowLeft' | 'chevron' | 'parallelogram' | 'trapezoid' | 'plus' | 'heart' | 'cloud';

/** The path of a K5 preset. `adjustments` are the element's raw values: they are resolved (defaulted and clamped) here. */
export function presetPath(preset: Exclude<ShapePreset, OriginalPreset>, w: number, h: number, adjustments?: readonly number[]): PresetPath {
  const [a1 = 0, a2 = 0, a3 = 0] = resolveAdjustments(preset, adjustments);
  const ss = Math.min(w, h);
  const hc = w / 2;
  const vc = h / 2;
  const d = (path: string): PresetPath => ({ d: path });

  switch (preset) {
    // ---- rectangles ----
    case 'snip1Rect': return d(cornered(w, h, [0, -a1 * ss, 0, 0]));
    case 'snip2SameRect': return d(cornered(w, h, [-a1 * ss, -a1 * ss, -a2 * ss, -a2 * ss]));
    case 'snip2DiagRect': return d(cornered(w, h, [-a1 * ss, -a2 * ss, -a1 * ss, -a2 * ss]));
    case 'snipRoundRect': return d(cornered(w, h, [a1 * ss, -a2 * ss, 0, 0]));
    case 'round1Rect': return d(cornered(w, h, [0, a1 * ss, 0, 0]));
    case 'round2SameRect': return d(cornered(w, h, [a1 * ss, a1 * ss, a2 * ss, a2 * ss]));
    case 'round2DiagRect': return d(cornered(w, h, [a1 * ss, a2 * ss, a1 * ss, a2 * ss]));
    case 'foldedCorner': {
      const f = a1 * ss;
      return d(poly([[0, 0], [w, 0], [w, h - f], [w - f, h], [0, h]]) + ` M${pt([w - f, h])} L${pt([w - f, h - f])} L${pt([w, h - f])}`);
    }
    case 'frame': {
      const t = Math.min(a1 * ss, ss / 2);
      return d(poly([[0, 0], [w, 0], [w, h], [0, h]]) + ' ' + poly([[t, t], [t, h - t], [w - t, h - t], [w - t, t]]));
    }
    case 'plaque': {
      const c = a1 * ss;
      return d(
        `M${r2(c)} 0 L${r2(w - c)} 0` + arc(c, c, false, w, c) + ` L${r2(w)} ${r2(h - c)}` + arc(c, c, false, w - c, h) +
          ` L${r2(c)} ${r2(h)}` + arc(c, c, false, 0, h - c) + ` L0 ${r2(c)}` + arc(c, c, false, c, 0) + ' Z',
      );
    }
    case 'bevel': {
      const t = a1 * ss;
      return d(
        poly([[0, 0], [w, 0], [w, h], [0, h]]) + ' ' + poly([[t, t], [w - t, t], [w - t, h - t], [t, h - t]]) + ' ' +
          [seg([0, 0], [t, t]), seg([w, 0], [w - t, t]), seg([w, h], [w - t, h - t]), seg([0, h], [t, h - t])].join(' '),
      );
    }
    case 'can': {
      const y = Math.min((ss * a1) / 2, h / 2);
      return d(
        `M0 ${r2(y)}` + arc(hc, y, true, w, y) + ` L${r2(w)} ${r2(h - y)}` + arc(hc, y, true, 0, h - y) + ' Z' +
          ` M${r2(w)} ${r2(y)}` + arc(hc, y, true, 0, y),
      );
    }
    case 'cube': {
      const y = a1 * ss;
      return d(
        poly([[0, y], [y, 0], [w, 0], [w, h - y], [w - y, h], [0, h]]) + ' ' +
          [seg([0, y], [w - y, y]), seg([w - y, y], [w, 0]), seg([w - y, y], [w - y, h])].join(' '),
      );
    }

    // ---- basic ----
    case 'donut': {
      const t = a1 * ss;
      return d(ellipsePath(hc, vc, hc, vc) + ' ' + ellipsePath(hc, vc, Math.max(hc - t, 0.01), Math.max(vc - t, 0.01), false));
    }
    case 'noSmoking': {
      const t = a1 * ss;
      // the bar's centre line ends inside the ring (3/4 of the way out), so its square ends stay inside the outer edge
      const mx = hc - t * 0.75;
      const my = vc - t * 0.75;
      const len = Math.hypot(mx, my);
      const ux = (mx / len) * (t / 2);
      const uy = (my / len) * (t / 2);
      const k = Math.SQRT1_2;
      // the bar runs clockwise on screen, so it adds to the ring's nonzero fill instead of cutting it
      const bar: Pt[] = [
        [hc - mx * k + uy, vc - my * k - ux], [hc + mx * k + uy, vc + my * k - ux],
        [hc + mx * k - uy, vc + my * k + ux], [hc - mx * k - uy, vc - my * k + ux],
      ];
      return d(ellipsePath(hc, vc, hc, vc) + ' ' + ellipsePath(hc, vc, Math.max(hc - t, 0.01), Math.max(vc - t, 0.01), false) + ' ' + poly(bar));
    }
    case 'teardrop': {
      const tx = hc + hc * a1;
      const ty = vc - vc * a1;
      return d(
        `M0 ${r2(vc)}` + arc(hc, vc, true, hc, 0) + ` Q${pt([(hc + tx) / 2, 0])} ${pt([tx, ty])} Q${pt([w, (vc + ty) / 2])} ${pt([w, vc])}` +
          arc(hc, vc, true, hc, h) + arc(hc, vc, true, 0, vc) + ' Z',
      );
    }
    case 'pie': {
      const turns = sweepTurns(a1, a2);
      return d(`M${pt([hc, vc])} L${pt(onEllipse(w, h, a1))}` + arc(hc, vc, true, ...onEllipse(w, h, a2), turns > 0.5) + ' Z');
    }
    case 'chord': {
      const turns = sweepTurns(a1, a2);
      return d(`M${pt(onEllipse(w, h, a1))}` + arc(hc, vc, true, ...onEllipse(w, h, a2), turns > 0.5) + ' Z');
    }
    case 'blockArc': {
      const t = Math.min(a3 * ss, ss / 2);
      const large = sweepTurns(a1, a2) > 0.5;
      return d(
        `M${pt(onEllipse(w, h, a1))}` + arc(hc, vc, true, ...onEllipse(w, h, a2), large) + ` L${pt(onEllipse(w, h, a2, t))}` +
          arc(hc - t, vc - t, false, ...onEllipse(w, h, a1, t), large) + ' Z',
      );
    }
    case 'moon':
      return d(`M${r2(w)} 0` + arc(w, vc, false, w, h) + arc(w * (1 - a1), vc, true, w, 0) + ' Z');
    case 'smileyFace': {
      const eye = (x: number): string => ellipsePath(x * w, 0.36 * h, 0.065 * w, 0.085 * h, false);
      const mouth = `M${pt([0.22 * w, 0.62 * h])} Q${pt([hc, 0.92 * h])} ${pt([0.78 * w, 0.62 * h])} Q${pt([hc, 0.8 * h])} ${pt([0.22 * w, 0.62 * h])} Z`;
      return d(ellipsePath(hc, vc, hc, vc) + ' ' + eye(0.3) + ' ' + eye(0.7) + ' ' + mouth);
    }
    case 'lightningBolt':
      return d(poly(BOLT.map(([x, y]) => [(x / 21600) * w, (y / 21600) * h] as Pt)));
    case 'heptagon': return d(poly(ngon(7, w, h, -Math.PI / 2)));
    case 'decagon': return d(poly(ngon(10, w, h, 0)));
    case 'dodecagon': return d(poly(ngon(12, w, h, -Math.PI / 2)));

    // ---- block arrows ----
    case 'rightArrow': {
      const hl = Math.min(ss * a2, w);
      const dy = (h * a1) / 2;
      return d(poly([[0, vc - dy], [w - hl, vc - dy], [w - hl, 0], [w, vc], [w - hl, h], [w - hl, vc + dy], [0, vc + dy]]));
    }
    case 'leftArrow': {
      const hl = Math.min(ss * a2, w);
      const dy = (h * a1) / 2;
      return d(poly([[w, vc - dy], [hl, vc - dy], [hl, 0], [0, vc], [hl, h], [hl, vc + dy], [w, vc + dy]]));
    }
    case 'upArrow': {
      const hl = Math.min(ss * a2, h);
      const dx = (w * a1) / 2;
      return d(poly([[hc - dx, h], [hc - dx, hl], [0, hl], [hc, 0], [w, hl], [hc + dx, hl], [hc + dx, h]]));
    }
    case 'downArrow': {
      const hl = Math.min(ss * a2, h);
      const dx = (w * a1) / 2;
      return d(poly([[hc - dx, 0], [hc - dx, h - hl], [0, h - hl], [hc, h], [w, h - hl], [hc + dx, h - hl], [hc + dx, 0]]));
    }
    case 'leftRightArrow': {
      const hl = Math.min(ss * a2, w / 2);
      const dy = (h * a1) / 2;
      return d(poly([[0, vc], [hl, 0], [hl, vc - dy], [w - hl, vc - dy], [w - hl, 0], [w, vc], [w - hl, h], [w - hl, vc + dy], [hl, vc + dy], [hl, h]]));
    }
    case 'upDownArrow': {
      const hl = Math.min(ss * a2, h / 2);
      const dx = (w * a1) / 2;
      return d(poly([[hc, 0], [w, hl], [hc + dx, hl], [hc + dx, h - hl], [w, h - hl], [hc, h], [0, h - hl], [hc - dx, h - hl], [hc - dx, hl], [0, hl]]));
    }
    case 'homePlate': {
      const hl = Math.min(ss * a1, w);
      return d(poly([[0, 0], [w - hl, 0], [w, vc], [w - hl, h], [0, h]]));
    }

    // ---- stars ----
    case 'star4': return d(poly(star(4, w, h, a1 * 2)));
    case 'star6': return d(poly(star(6, w, h, a1 * 2, 2 / Math.sqrt(3))));
    case 'star8': return d(poly(star(8, w, h, a1 * 2)));
    case 'star12': return d(poly(star(12, w, h, a1 * 2)));

    // ---- callouts ----
    case 'wedgeRectCallout': return d(calloutPath(w, h, hc + w * a1, vc + h * a2, 0));
    case 'wedgeRoundRectCallout': return d(calloutPath(w, h, hc + w * a1, vc + h * a2, a3 * ss));
    case 'wedgeEllipseCallout': {
      const tx = hc + w * a1;
      const ty = vc + h * a2;
      const angle = Math.atan2((ty - vc) / vc, (tx - hc) / hc) / (2 * Math.PI);
      const half = 0.045;
      return d(`M${pt(onEllipse(w, h, angle + half))}` + arc(hc, vc, true, ...onEllipse(w, h, angle - half), true) + ` L${pt([tx, ty])} Z`);
    }

    // ---- flowchart ----
    case 'flowChartProcess': return d(poly([[0, 0], [w, 0], [w, h], [0, h]]));
    case 'flowChartAlternateProcess': return d(cornered(w, h, [ss / 6, ss / 6, ss / 6, ss / 6]));
    case 'flowChartDecision': return d(poly([[hc, 0], [w, vc], [hc, h], [0, vc]]));
    case 'flowChartInputOutput': return d(poly([[w / 5, 0], [w, 0], [(4 * w) / 5, h], [0, h]]));
    case 'flowChartTerminator': {
      const cap = 0.1609 * w;
      return d(`M${r2(cap)} 0 L${r2(w - cap)} 0` + arc(cap, vc, true, w - cap, h) + ` L${r2(cap)} ${r2(h)}` + arc(cap, vc, true, cap, 0) + ' Z');
    }
    case 'flowChartConnector': return d(ellipsePath(hc, vc, hc, vc));
    case 'flowChartPredefinedProcess':
      return d(poly([[0, 0], [w, 0], [w, h], [0, h]]) + ' ' + seg([w / 8, 0], [w / 8, h]) + ' ' + seg([(7 * w) / 8, 0], [(7 * w) / 8, h]));
    case 'flowChartDelay': return d(`M0 0 L${r2(hc)} 0` + arc(hc, vc, true, hc, h) + ` L0 ${r2(h)} Z`);
    case 'flowChartPreparation': return d(poly([[0, vc], [w / 5, 0], [(4 * w) / 5, 0], [w, vc], [(4 * w) / 5, h], [w / 5, h]]));
    case 'flowChartOffpageConnector': return d(poly([[0, 0], [w, 0], [w, 0.8 * h], [hc, h], [0, 0.8 * h]]));
    case 'flowChartMerge': return d(poly([[0, 0], [w, 0], [hc, h]]));

    // ---- brackets and braces ----
    case 'leftBracket': {
      const y = Math.min(a1 * ss, h / 2);
      const open = `M${r2(w)} 0` + arc(w, y, false, 0, y) + ` L0 ${r2(h - y)}` + arc(w, y, false, w, h);
      return { d: open + ' Z', strokeD: open };
    }
    case 'rightBracket': {
      const y = Math.min(a1 * ss, h / 2);
      const open = 'M0 0' + arc(w, y, true, w, y) + ` L${r2(w)} ${r2(h - y)}` + arc(w, y, true, 0, h);
      return { d: open + ' Z', strokeD: open };
    }
    case 'leftBrace':
    case 'rightBrace': {
      const ym = a2 * h;
      const open = brace(w, h, Math.min(a1 * ss, ym / 2, (h - ym) / 2), ym, preset === 'rightBrace');
      return { d: open + ' Z', strokeD: open };
    }
    case 'bracketPair': {
      const c = Math.min(a1 * ss, ss / 2);
      const left = `M${r2(c)} 0` + arc(c, c, false, 0, c) + ` L0 ${r2(h - c)}` + arc(c, c, false, c, h);
      const right = `M${r2(w - c)} 0` + arc(c, c, true, w, c) + ` L${r2(w)} ${r2(h - c)}` + arc(c, c, true, w - c, h);
      return { d: cornered(w, h, [c, c, c, c]), strokeD: left + ' ' + right };
    }

    // ---- math and lines ----
    case 'mathPlus': {
      const dx = (w * 0.36745);
      const dy = (h * 0.36745);
      const t = Math.min((ss * a1) / 2, dx, dy);
      return d(poly([
        [hc - dx, vc - t], [hc - t, vc - t], [hc - t, vc - dy], [hc + t, vc - dy], [hc + t, vc - t], [hc + dx, vc - t],
        [hc + dx, vc + t], [hc + t, vc + t], [hc + t, vc + dy], [hc - t, vc + dy], [hc - t, vc + t], [hc - dx, vc + t],
      ]));
    }
    case 'mathMinus': {
      const dx = w * 0.36745;
      const dy = Math.min((h * a1) / 2, vc);
      return d(poly([[hc - dx, vc - dy], [hc + dx, vc - dy], [hc + dx, vc + dy], [hc - dx, vc + dy]]));
    }
    case 'lineInv': return d(seg([w, 0], [0, h]));
  }
}
