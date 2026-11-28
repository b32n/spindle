import type { PptxColor } from '@b32nio/spindle-shared/export/pptx';
import { colorFromPptx, colorToPptx, fillFromPptx, fillToPptx, finite, ptToPx, pxToPt, strokeFromPptx, strokeToPptx } from './pptx-values';

const red: PptxColor = { kind: 'rgb', hex: '#FF0000' };
const accent: PptxColor = { kind: 'theme', slot: 'accent1', lumMod: 0.75, alpha: 0.5 };

describe('units', () => {
  it('converts points to px and back (a point is 4/3 px) and keeps 11 pt stable', () => {
    expect(ptToPx(18)).toBe(24);
    expect(pxToPt(24)).toBe(18);
    expect(ptToPx(11)).toBe(14.6667);
    expect(Math.round(pxToPt(ptToPx(11)) * 100)).toBe(1100);
    expect(ptToPx(0)).toBe(0);
  });
  it('finite() keeps finite numbers only', () => {
    expect(finite(3)).toBe(3);
    for (const v of [NaN, Infinity, '3', null, undefined, {}]) expect(finite(v)).toBeUndefined();
  });
});

describe('colours', () => {
  it('copies colours both ways without sharing the object', () => {
    expect(colorFromPptx(accent)).toEqual(accent);
    expect(colorFromPptx(accent)).not.toBe(accent);
    expect(colorToPptx(colorFromPptx(red))).toEqual(red);
  });
});

describe('fills', () => {
  it('maps the fills the model has, both ways', () => {
    expect(fillFromPptx(undefined)).toBeUndefined();
    expect(fillToPptx(undefined)).toBeUndefined();
    expect(fillFromPptx({ kind: 'none' })).toEqual({ kind: 'none' });
    expect(fillToPptx({ kind: 'none' })).toEqual({ kind: 'none' });
    expect(fillFromPptx({ kind: 'solid', color: red })).toEqual({ kind: 'solid', color: red });
    expect(fillToPptx({ kind: 'solid', color: red })).toEqual({ kind: 'solid', color: red });
  });
  it('maps linear, radial and path gradients', () => {
    const stops = [{ pos: 0, color: red }, { pos: 1, color: accent }];
    expect(fillFromPptx({ kind: 'gradient', shape: 'linear', angle: 45, stops })).toEqual({ kind: 'gradient', type: 'linear', angle: 45, stops });
    expect(fillFromPptx({ kind: 'gradient', shape: 'linear', stops })).toMatchObject({ angle: 0 });
    expect(fillFromPptx({ kind: 'gradient', shape: 'radial', stops })).toMatchObject({ type: 'radial' });
    expect(fillFromPptx({ kind: 'gradient', shape: 'rect', stops })).toMatchObject({ type: 'path' });
    expect(fillFromPptx({ kind: 'gradient', shape: 'shape', stops })).toMatchObject({ type: 'path' });
    expect(fillToPptx({ kind: 'gradient', type: 'linear', angle: 10, stops })).toEqual({ kind: 'gradient', shape: 'linear', angle: 10, stops });
    expect(fillToPptx({ kind: 'gradient', type: 'linear', stops })).toMatchObject({ angle: 90 });
    expect(fillToPptx({ kind: 'gradient', type: 'radial', stops })).toMatchObject({ shape: 'radial' });
    expect(fillToPptx({ kind: 'gradient', type: 'path', stops })).toMatchObject({ shape: 'rect' });
  });
  it('maps patterns with defaults for a missing colour', () => {
    expect(fillFromPptx({ kind: 'pattern', prst: 'pct50', fg: red })).toEqual({ kind: 'pattern', preset: 'pct50', fg: red, bg: { kind: 'rgb', hex: '#FFFFFF' } });
    expect(fillFromPptx({ kind: 'pattern', prst: 'pct50', bg: red })).toEqual({ kind: 'pattern', preset: 'pct50', fg: { kind: 'rgb', hex: '#000000' }, bg: red });
    expect(fillToPptx({ kind: 'pattern', preset: 'x', fg: red, bg: accent })).toEqual({ kind: 'pattern', prst: 'x', fg: red, bg: accent });
  });
  it('has no model for the picture of a fill coming in, and passes the model one out (the writer reports it)', () => {
    expect(fillFromPptx({ kind: 'picture', part: 'ppt/media/a.png' })).toBeUndefined();
    expect(fillToPptx({ kind: 'picture', src: 'data:image/png;base64,AA==', mode: 'stretch' })).toEqual({ kind: 'picture' });
    expect(fillToPptx({ kind: 'picture', src: 'data:image/png;base64,AA==', mode: 'tile', alpha: 0.5 })).toEqual({ kind: 'picture', tile: true, alpha: 0.5 });
  });
});

describe('outlines', () => {
  it('takes the colour of a solid, gradient or pattern line, and none from anything else', () => {
    expect(strokeFromPptx(undefined)).toBeUndefined();
    expect(strokeFromPptx({})).toBeUndefined();
    expect(strokeFromPptx({ fill: { kind: 'none' }, width: 2 })).toBeUndefined();
    expect(strokeFromPptx({ fill: { kind: 'solid', color: red }, width: 2 })).toEqual({ color: red, width: 2 });
    expect(strokeFromPptx({ fill: { kind: 'solid', color: red } })).toEqual({ color: red, width: 1 });
    expect(strokeFromPptx({ fill: { kind: 'gradient', shape: 'linear', stops: [{ pos: 0, color: accent }] } })).toMatchObject({ color: accent });
    expect(strokeFromPptx({ fill: { kind: 'gradient', shape: 'linear', stops: [] } })).toBeUndefined();
    expect(strokeFromPptx({ fill: { kind: 'pattern', prst: 'x', fg: red } })).toMatchObject({ color: red });
    expect(strokeFromPptx({ fill: { kind: 'pattern', prst: 'x' } })).toBeUndefined();
  });
  it('maps dashes to the model\'s two', () => {
    const line = (dash: string) => strokeFromPptx({ fill: { kind: 'solid', color: red }, dash });
    expect(line('sysDot')).toMatchObject({ dash: 'dot' });
    expect(line('dash')).toMatchObject({ dash: 'dash' });
    expect(line('custom')).toMatchObject({ dash: 'dash' });
    expect(line('solid')).not.toHaveProperty('dash');
  });
  it('writes a stroke back', () => {
    expect(strokeToPptx(undefined)).toBeUndefined();
    expect(strokeToPptx({ color: red, width: 3 })).toEqual({ fill: { kind: 'solid', color: red }, width: 3 });
    expect(strokeToPptx({ color: red, width: 3, dash: 'dash' })).toMatchObject({ dash: 'dash' });
    expect(strokeToPptx({ color: red, width: 3, dash: 'dot' })).toMatchObject({ dash: 'sysDot' });
    expect(strokeToPptx({ color: red, width: 3, dash: 'solid' })).not.toHaveProperty('dash');
  });
});
