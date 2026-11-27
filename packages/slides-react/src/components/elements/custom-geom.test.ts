import type { CustomPath } from '@b32nio/spindle-slides-core';
import { customPathD, pathShade } from './custom-geom';

const path = (commands: CustomPath['commands'], w = 100, h = 100): CustomPath => ({ w, h, commands });

describe('customPathD', () => {
  it('scales path space onto the box and writes every command kind', () => {
    const d = customPathD(
      path(
        [
          { op: 'move', x: 0, y: 0 },
          { op: 'line', x: 100, y: 0 },
          { op: 'cubic', x1: 100, y1: 50, x2: 50, y2: 100, x: 0, y: 100 },
          { op: 'quad', x1: 0, y1: 50, x: 0, y: 0 },
          { op: 'close' },
        ],
        100,
        100,
      ),
      200,
      50,
    );
    expect(d).toBe('M0 0 L200 0 C200 25 100 50 0 50 Q0 25 0 0 Z');
  });

  it('draws an arc from the current point and keeps the large-arc and sweep flags', () => {
    // from (0, 10) at angle 180 on a circle of radius 10, a half turn clockwise ends at (20, 10)
    const d = customPathD(path([{ op: 'move', x: 0, y: 10 }, { op: 'arc', wR: 10, hR: 10, stAng: 180, swAng: 90 }], 20, 20), 20, 20);
    expect(d).toBe('M0 10 A10 10 0 0 1 10 0');
    expect(customPathD(path([{ op: 'move', x: 0, y: 10 }, { op: 'arc', wR: 10, hR: 10, stAng: 180, swAng: 270 }], 20, 20), 20, 20)).toMatch(/A10 10 0 1 1 /);
    expect(customPathD(path([{ op: 'move', x: 10, y: 0 }, { op: 'arc', wR: 10, hR: 10, stAng: 270, swAng: -90 }], 20, 20), 20, 20)).toBe('M10 0 A10 10 0 0 0 0 10');
  });

  it('scales arc radii independently per axis', () => {
    const d = customPathD(path([{ op: 'move', x: 0, y: 10 }, { op: 'arc', wR: 10, hR: 10, stAng: 180, swAng: 90 }], 20, 20), 40, 20);
    expect(d).toBe('M0 10 A20 10 0 0 1 20 0');
  });

  it('splits a full turn into two half arcs that end where they began', () => {
    const d = customPathD(path([{ op: 'move', x: 0, y: 10 }, { op: 'arc', wR: 10, hR: 10, stAng: 180, swAng: 360 }], 20, 20), 20, 20);
    expect(d).toBe('M0 10 A10 10 0 0 1 20 10 A10 10 0 0 1 0 10');
  });

  it('skips a degenerate arc', () => {
    const base = [{ op: 'move', x: 1, y: 1 }] as CustomPath['commands'];
    for (const arc of [{ wR: 0, hR: 5, stAng: 0, swAng: 90 }, { wR: 5, hR: 0, stAng: 0, swAng: 90 }, { wR: 5, hR: 5, stAng: 0, swAng: 0 }]) {
      expect(customPathD(path([...base, { op: 'arc', ...arc }]), 100, 100)).toBe('M1 1');
    }
  });

  it('continues from the subpath start after a close', () => {
    const d = customPathD(
      path([{ op: 'move', x: 0, y: 10 }, { op: 'line', x: 10, y: 10 }, { op: 'close' }, { op: 'arc', wR: 10, hR: 10, stAng: 180, swAng: 90 }], 20, 20),
      20,
      20,
    );
    expect(d.endsWith('A10 10 0 0 1 10 0')).toBe(true);
  });

  it('is empty without commands', () => {
    expect(customPathD(path([]), 10, 10)).toBe('');
  });
});

describe('pathShade', () => {
  it('tints for the lighten and darken modes only', () => {
    expect(pathShade('lighten')).toEqual({ color: '#ffffff', opacity: 0.4 });
    expect(pathShade('lightenLess')).toEqual({ color: '#ffffff', opacity: 0.2 });
    expect(pathShade('darken')).toEqual({ color: '#000000', opacity: 0.4 });
    expect(pathShade('darkenLess')).toEqual({ color: '#000000', opacity: 0.2 });
    expect(pathShade('norm')).toBeUndefined();
    expect(pathShade('none')).toBeUndefined();
    expect(pathShade(undefined)).toBeUndefined();
  });
});
