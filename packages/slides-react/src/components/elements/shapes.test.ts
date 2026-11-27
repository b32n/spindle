import { SHAPE_PRESETS, SHAPE_PRESET_INFO, isShapePreset, unknownShape, type ShapePreset } from '@b32nio/spindle-slides-core';
import { shapeGeom, type ShapeGeom } from './shapes';

const ARGS: Record<string, number> = { M: 2, L: 2, H: 1, V: 1, C: 6, Q: 4, A: 7, Z: 0 };

/** The end points a path visits (arc and curve control points are not on the outline, except Q controls kept). */
function points(d: string): Array<[number, number]> {
  const out: Array<[number, number]> = [];
  let cur: [number, number] = [0, 0];
  const tokens = d.match(/[MLHVQCAZ]|-?\d+(?:\.\d+)?(?:e-?\d+)?/g) ?? [];
  for (let i = 0; i < tokens.length; ) {
    const cmd = tokens[i++];
    const n = ARGS[cmd];
    if (n === undefined) throw new Error(`bad token ${cmd} in ${d}`);
    const nums = tokens.slice(i, i + n).map(Number);
    i += n;
    if (cmd === 'H') {
      cur = [nums[0], cur[1]];
      out.push(cur);
      continue;
    }
    if (cmd === 'V') {
      cur = [cur[0], nums[0]];
      out.push(cur);
      continue;
    }
    if (cmd === 'A') out.push([nums[5], nums[6]]);
    else if (cmd === 'Q') out.push([nums[2], nums[3]]);
    else if (cmd === 'C') out.push([nums[4], nums[5]]);
    else if (n === 2) out.push([nums[0], nums[1]]);
    if (out.length) cur = out[out.length - 1];
  }
  return out;
}

/**
 * Points along the whole outline, arcs and curves included (`points` only has the end points). Arcs follow the SVG
 * endpoint-to-centre conversion, radii scaled up when too small; curves are sampled at 16 steps.
 */
function outline(d: string): Array<[number, number]> {
  const out: Array<[number, number]> = [];
  let cur: [number, number] = [0, 0];
  let start: [number, number] = [0, 0];
  const tokens = d.match(/[MLHVQCAZ]|-?\d+(?:\.\d+)?(?:e-?\d+)?/g) ?? [];
  for (let i = 0; i < tokens.length; ) {
    const cmd = tokens[i++];
    const nums = tokens.slice(i, i + ARGS[cmd]).map(Number);
    i += ARGS[cmd];
    const from = cur;
    if (cmd === 'Z') cur = start;
    else if (cmd === 'H') cur = [nums[0], from[1]];
    else if (cmd === 'V') cur = [from[0], nums[0]];
    else if (cmd === 'A') cur = [nums[5], nums[6]];
    else cur = [nums[nums.length - 2], nums[nums.length - 1]];
    if (cmd === 'M') start = cur;
    if (cmd === 'A') {
      const [rx0, ry0, , large, sweep] = nums;
      const [x1, y1] = [(from[0] - cur[0]) / 2, (from[1] - cur[1]) / 2];
      const lambda = (x1 / rx0) ** 2 + (y1 / ry0) ** 2;
      const [rx, ry] = lambda > 1 ? [rx0 * Math.sqrt(lambda), ry0 * Math.sqrt(lambda)] : [rx0, ry0];
      const num = rx * rx * ry * ry - rx * rx * y1 * y1 - ry * ry * x1 * x1;
      const k = (large === sweep ? -1 : 1) * Math.sqrt(Math.max(0, num / (rx * rx * y1 * y1 + ry * ry * x1 * x1)));
      const [cx, cy] = [k * ((rx * y1) / ry) + (from[0] + cur[0]) / 2, k * ((-ry * x1) / rx) + (from[1] + cur[1]) / 2];
      const a0 = Math.atan2((from[1] - cy) / ry, (from[0] - cx) / rx);
      let da = Math.atan2((cur[1] - cy) / ry, (cur[0] - cx) / rx) - a0;
      if (sweep && da < 0) da += 2 * Math.PI;
      if (!sweep && da > 0) da -= 2 * Math.PI;
      for (let t = 1; t <= 32; t++) out.push([cx + rx * Math.cos(a0 + (da * t) / 32), cy + ry * Math.sin(a0 + (da * t) / 32)]);
    } else if (cmd === 'C' || cmd === 'Q') {
      const ctrl: Array<[number, number]> = [from];
      for (let j = 0; j < nums.length; j += 2) ctrl.push([nums[j], nums[j + 1]]);
      for (let t = 1; t <= 16; t++) {
        let layer = ctrl;
        while (layer.length > 1) layer = layer.slice(1).map((q, j) => [layer[j][0] + (q[0] - layer[j][0]) * (t / 16), layer[j][1] + (q[1] - layer[j][1]) * (t / 16)] as [number, number]);
        out.push(layer[0]);
      }
    } else out.push(cur);
  }
  return out;
}

const path = (g: ShapeGeom): Extract<ShapeGeom, { type: 'path' }> => {
  if (g.type !== 'path') throw new Error('expected a path');
  return g;
};

const NEW = SHAPE_PRESETS.filter((p) => !['rect', 'roundRect', 'ellipse', 'triangle', 'rightTriangle', 'diamond', 'pentagon', 'hexagon', 'octagon', 'star5', 'arrowRight', 'arrowLeft', 'chevron', 'parallelogram', 'trapezoid', 'plus', 'heart', 'cloud'].includes(p));
const CALLOUTS = ['wedgeRectCallout', 'wedgeRoundRectCallout', 'wedgeEllipseCallout'];

describe('shapeGeom: the preset library', () => {
  it('draws every preset as a non-empty, finite path inside its box (arcs and curves sampled, tails excepted)', () => {
    for (const [w, h] of [[200, 100], [100, 200], [120, 120], [1, 1]] as const) {
      for (const p of SHAPE_PRESETS) {
        const g = shapeGeom(p, w, h);
        const d = g.type === 'path' ? g.d : 'M0 0';
        expect(d.length).toBeGreaterThan(3);
        expect(d).not.toMatch(/NaN|Infinity/);
        if (g.type === 'ellipse') {
          expect([g.cx, g.cy, g.rx, g.ry].every(Number.isFinite)).toBe(true);
          continue;
        }
        for (const [x, y] of [...outline(d), ...(g.strokeD ? outline(g.strokeD) : [])]) {
          const slack = CALLOUTS.includes(p) ? Math.max(w, h) : 0.51;
          expect({ p, w, h, ok: x >= -slack && x <= w + slack && y >= -slack && y <= h + slack }).toEqual({ p, w, h, ok: true });
        }
      }
    }
  });

  it('adds 58 presets to the original 18', () => {
    expect(NEW.length).toBe(58);
    for (const p of NEW) expect(isShapePreset(p)).toBe(true);
  });

  it('lets a wide arrow keep a head longer than the box height', () => {
    // head length is ss * adj2 = 100 * 3 on a 400 x 100 box, so the neck sits at x = 100
    const xs = points(path(shapeGeom('rightArrow', 400, 100, [0.5, 3])).d).map((p) => p[0]);
    expect(xs).toContain(100);
  });

  it('keeps the original 18 geometry byte for byte', () => {
    expect(path(shapeGeom('rect', 100, 50)).d).toBe('M0 0 L100 0 L100 50 L0 50 Z');
    expect(path(shapeGeom('roundRect', 100, 50)).d).toBe(path(shapeGeom('roundRect', 100, 50, [0.15])).d);
    expect(shapeGeom('ellipse', 100, 50)).toEqual({ type: 'ellipse', cx: 50, cy: 25, rx: 50, ry: 25 });
    expect(path(shapeGeom('arrowRight', 100, 100)).d).toBe('M0 30 L60 30 L60 0 L100 50 L60 100 L60 70 L0 70 Z');
  });

  it('draws distinct outlines for every preset at the default size', () => {
    const seen = new Map<string, string>();
    for (const p of SHAPE_PRESETS) {
      const g = shapeGeom(p, 160, 120);
      const key = JSON.stringify(g);
      // flowchart process and decision are deliberately the rectangle and the diamond
      if (seen.has(key)) expect([p, seen.get(key)].sort().join()).toMatch(/^(flowChartProcess,rect|diamond,flowChartDecision)$/);
      seen.set(key, p);
    }
  });
});

describe('shapeGeom: oracle values (OOXML guide semantics at default adjustments)', () => {
  it('snip1Rect snips the top-right by a sixth of the short side', () => {
    expect(path(shapeGeom('snip1Rect', 120, 60)).d).toBe('M0 0 L110 0 L120 10 L120 60 L0 60 L0 0 Z');
  });

  it('round1Rect rounds only the top-right corner', () => {
    expect(path(shapeGeom('round1Rect', 120, 60)).d).toBe('M0 0 L110 0 A10 10 0 0 1 120 10 L120 60 L0 60 L0 0 Z');
  });

  it('rightArrow: shaft half the height, head half the short side', () => {
    expect(path(shapeGeom('rightArrow', 100, 100)).d).toBe('M0 25 L50 25 L50 0 L100 50 L50 100 L50 75 L0 75 Z');
  });

  it('upArrow and downArrow are the transposes of the horizontal ones', () => {
    expect(path(shapeGeom('upArrow', 100, 100)).d).toBe('M25 100 L25 50 L0 50 L50 0 L100 50 L75 50 L75 100 Z');
    expect(path(shapeGeom('downArrow', 100, 100)).d).toBe('M25 0 L25 50 L0 50 L50 100 L100 50 L75 50 L75 0 Z');
  });

  it('frame has a counter-wound hole one eighth wide', () => {
    expect(path(shapeGeom('frame', 80, 80)).d).toBe('M0 0 L80 0 L80 80 L0 80 Z M10 10 L10 70 L70 70 L70 10 Z');
  });

  it('pie sweeps three quarters by default, and a half-turn pie takes the small arc flag', () => {
    expect(path(shapeGeom('pie', 100, 100)).d).toBe('M50 50 L100 50 A50 50 0 1 1 50 0 Z');
    expect(path(shapeGeom('pie', 100, 100, [0, 0.25])).d).toBe('M50 50 L100 50 A50 50 0 0 1 50 100 Z');
    // end before start wraps around
    expect(path(shapeGeom('pie', 100, 100, [0.75, 0.25])).d).toContain('A50 50 0 0 1');
  });

  it('chord and blockArc close their arcs', () => {
    expect(path(shapeGeom('chord', 100, 100)).d).toMatch(/^M85\.36 85\.36 A50 50 0 1 1 50 0 Z$/);
    expect(path(shapeGeom('blockArc', 100, 100)).d).toBe('M0 50 A50 50 0 0 1 100 50 L75 50 A25 25 0 0 0 25 50 Z');
  });

  it('moon is a crescent half as thick as it is wide', () => {
    expect(path(shapeGeom('moon', 100, 100)).d).toBe('M100 0 A100 50 0 0 0 100 100 A50 50 0 0 1 100 0 Z');
  });

  it('star4 default points at the edge midpoints with a quarter-radius waist', () => {
    const pts = points(path(shapeGeom('star4', 100, 100)).d);
    expect(pts[0]).toEqual([50, 0]);
    expect(pts[1]).toEqual([58.84, 41.16]);
  });

  it('star6 spans the full width', () => {
    const xs = points(path(shapeGeom('star6', 100, 100)).d).map((p) => p[0]);
    expect(Math.min(...xs)).toBeCloseTo(0, 1);
    expect(Math.max(...xs)).toBeCloseTo(100, 1);
  });

  it('a wedge callout puts its tail on the edge nearest the tip', () => {
    const tail = (a: number[]) => path(shapeGeom('wedgeRectCallout', 120, 60, a)).d;
    expect(tail([-0.2, 0.6])).toContain('L50 60 L36 66 L20 60'); // below, left half
    expect(tail([0.2, 0.6])).toContain('L100 60 L84 66 L70 60'); // below, right half
    expect(tail([0.2, -0.6])).toContain('L70 0 L84 -6 L100 0'); // above
    expect(tail([-0.9, 0.1])).toContain('L0 50 L-48 36 L0 35'); // left, lower half
    expect(tail([0.9, 0.1])).toContain('L120 35 L168 36 L120 50'); // right, lower half
    expect(tail([0.9, -0.5])).toContain('L120 10 L168 0 L120 25'); // right, upper half
  });

  it('rounded callouts use arcs, plain ones do not', () => {
    expect(path(shapeGeom('wedgeRoundRectCallout', 120, 60)).d).toContain('A');
    expect(path(shapeGeom('wedgeRectCallout', 120, 60)).d).not.toContain('A');
    expect(path(shapeGeom('wedgeRoundRectCallout', 120, 60, [-0.2, 0.6, 0])).d).not.toContain('A');
  });

  it('an oval callout draws its tail to the tip', () => {
    expect(path(shapeGeom('wedgeEllipseCallout', 100, 100, [0, 0.7])).d).toContain('L50 120');
  });

  it('brackets split the filled area from the stroked line', () => {
    for (const p of ['leftBracket', 'rightBracket', 'leftBrace', 'rightBrace', 'bracketPair'] as ShapePreset[]) {
      const g = path(shapeGeom(p, 40, 120));
      expect(g.strokeD).toBeDefined();
      expect(g.strokeD).not.toMatch(/Z/);
    }
    expect(path(shapeGeom('rect', 40, 40)).strokeD).toBeUndefined();
  });

  it('braces keep their point at the guide height and mirror each other', () => {
    const left = path(shapeGeom('leftBrace', 40, 120, [1 / 12, 0.25])).strokeD!;
    expect(left).toContain('0 30'); // the point at a quarter of the height
    const right = path(shapeGeom('rightBrace', 40, 120, [1 / 12, 0.25])).strokeD!;
    expect(right).toContain('40 30');
  });

  it('noSmoking adds a clockwise bar across the ring (still filled where it crosses)', () => {
    const d = path(shapeGeom('noSmoking', 100, 100)).d;
    expect(d.split('M').length - 1).toBe(3);
    // the bar (third subpath) runs clockwise on screen: positive shoelace area in y-down coordinates
    const bar = points(d.slice(d.lastIndexOf('M')));
    const area = bar.reduce((sum, [x, y], i) => sum + x * bar[(i + 1) % bar.length][1] - bar[(i + 1) % bar.length][0] * y, 0);
    expect(area).toBeGreaterThan(0);
    for (const [w, h] of [[200, 100], [100, 200], [120, 120], [300, 60]] as const) {
      const b = points('M' + path(shapeGeom('noSmoking', w, h)).d.split('M')[3]);
      const ar = b.reduce((sum, [x, y], i) => sum + x * b[(i + 1) % b.length][1] - b[(i + 1) % b.length][0] * y, 0);
      expect(ar).toBeGreaterThan(0);
      // the bar's corners stay inside the outer ellipse
      for (const [x, y] of b) expect(((x - w / 2) / (w / 2)) ** 2 + ((y - h / 2) / (h / 2)) ** 2).toBeLessThanOrEqual(1);
    }
  });

  it('smileyFace has eyes and a mouth cut from the face', () => {
    expect(path(shapeGeom('smileyFace', 100, 100)).d.split('M').length - 1).toBe(4);
  });

  it('lightningBolt scales its 21600 grid to the box', () => {
    expect(points(path(shapeGeom('lightningBolt', 216, 216)).d)[0]).toEqual([84.72, 0]);
  });

  it('math plus and minus are centred crosses and bars', () => {
    const plus = points(path(shapeGeom('mathPlus', 100, 100)).d);
    expect(plus).toHaveLength(12);
    const minus = points(path(shapeGeom('mathMinus', 100, 100)).d);
    expect(minus[0][0]).toBeCloseTo(13.26, 1);
    expect(minus[1][0]).toBeCloseTo(86.74, 1);
  });

  it('lineInv runs from the top-right to the bottom-left', () => {
    expect(path(shapeGeom('lineInv', 80, 40)).d).toBe('M80 0 L0 40');
  });

  it('flowchart shapes match their OOXML outlines', () => {
    expect(path(shapeGeom('flowChartInputOutput', 100, 50)).d).toBe('M20 0 L100 0 L80 50 L0 50 Z');
    expect(path(shapeGeom('flowChartMerge', 100, 50)).d).toBe('M0 0 L100 0 L50 50 Z');
    expect(path(shapeGeom('flowChartTerminator', 100, 40)).d).toContain('A16.09 20');
  });
});

describe('shapeGeom: adjustments', () => {
  it('clamps out-of-range guides', () => {
    const wide = path(shapeGeom('frame', 80, 80, [9])).d;
    expect(wide).toBe(path(shapeGeom('frame', 80, 80, [0.5])).d);
    const neg = path(shapeGeom('snip1Rect', 100, 100, [-3])).d;
    expect(neg).toBe(path(shapeGeom('snip1Rect', 100, 100, [0])).d);
    expect(path(shapeGeom('roundRect', 100, 50, [5])).d).toBe(path(shapeGeom('roundRect', 100, 50, [0.5])).d);
  });

  it('limits arrow heads to the box and mirrors left and right', () => {
    expect(path(shapeGeom('rightArrow', 50, 100, [0.5, 1])).d).toBe('M0 25 L0 25 L0 0 L50 50 L0 100 L0 75 L0 75 Z');
    expect(path(shapeGeom('leftArrow', 100, 100)).d).toBe('M100 25 L50 25 L50 0 L0 50 L50 100 L50 75 L100 75 Z');
  });

  it('applies every guide of a multi-guide preset', () => {
    expect(path(shapeGeom('snip2SameRect', 100, 100, [0.1, 0.3])).d).toBe('M10 0 L90 0 L100 10 L100 70 L70 100 L30 100 L0 70 L0 10 L10 0 Z');
  });

  it('keeps each default in its table', () => {
    for (const p of SHAPE_PRESETS) expect(SHAPE_PRESET_INFO[p].label.length).toBeGreaterThan(0);
  });
});

describe('shapeGeom: unknown presets and bad names', () => {
  it('draws an unknown preset as a default rounded rectangle, ignoring its raw adjustments', () => {
    const approx = shapeGeom(unknownShape('flowChartDocument'), 100, 50, [123456, -9]);
    expect(approx).toEqual(shapeGeom('roundRect', 100, 50));
  });

  it('draws a name that is neither as a rectangle', () => {
    expect(shapeGeom('bogus' as never, 100, 50)).toEqual(shapeGeom('rect', 100, 50));
  });
});
