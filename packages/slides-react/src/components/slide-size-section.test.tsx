import React from 'react';
import { SlideSizeSection } from './DeckControls';

const setSlideSize = jest.fn();
let size = { w: 1280, h: 720 };
jest.mock('../hooks', () => ({
  useDeck: () => ({ getSlideSize: () => size, setSlideSize }),
  useActiveSlideId: () => 's',
  useTheme: () => ({ name: 'Clean' }),
  useSlide: () => undefined,
}));

type P = Record<string, any>; // eslint-disable-line @typescript-eslint/no-explicit-any
const find = (node: React.ReactNode, pred: (e: React.ReactElement<P>) => boolean, out: React.ReactElement<P>[] = []) => {
  React.Children.forEach(node, (c) => {
    if (!React.isValidElement<P>(c)) return;
    if (pred(c)) out.push(c);
    find(c.props.children, pred, out);
  });
  return out;
};
const byLabel = (l: string) => (e: React.ReactElement<P>) => e.props['aria-label'] === l;

let slots: unknown[] = [];
let n = 0;
const render = () => {
  n = 0;
  return SlideSizeSection();
};

beforeEach(() => {
  slots = [];
  size = { w: 1280, h: 720 };
  setSlideSize.mockClear();
  jest.spyOn(React, 'useState').mockImplementation(((init: unknown) => {
    const i = n++;
    if (!(i in slots)) slots[i] = init;
    return [slots[i], (v: unknown) => { slots[i] = v; }];
  }) as never);
});
afterEach(() => jest.restoreAllMocks());

describe('SlideSizeSection', () => {
  it('lists the presets, selects the matching one and labels every control', () => {
    const out = render();
    const [select] = find(out, byLabel('Slide size'));
    expect(select.props.value).toBe('16:9');
    const options = find(select, (e) => e.type === 'option');
    expect(options.map((o) => o.props.value)).toEqual(['16:9', '16:10', '4:3', 'a4', 'a4-portrait', 'portrait']);
    expect(find(out, byLabel('Slide width'))[0].props.value).toBe('1280');
    expect(find(out, byLabel('Slide height'))[0].props.value).toBe('720');
    expect(find(out, byLabel('Apply slide size'))).toHaveLength(1);
  });

  it('picking a preset resizes with content scaling; an unknown value is ignored', () => {
    find(render(), byLabel('Slide size'))[0].props.onChange({ target: { value: '4:3' } });
    expect(setSlideSize).toHaveBeenCalledWith({ w: 960, h: 720 }, { scaleContent: true });
    setSlideSize.mockClear();
    find(render(), byLabel('Slide size'))[0].props.onChange({ target: { value: 'custom' } });
    expect(setSlideSize).not.toHaveBeenCalled();
  });

  it('a custom size shows a Custom option, applies clamped and rounded, and falls back for junk', () => {
    size = { w: 1000, h: 700 };
    let out = render();
    const [select] = find(out, byLabel('Slide size'));
    expect(select.props.value).toBe('custom');
    expect(find(select, (e) => e.type === 'option').map((o) => o.props.value)).toContain('custom');
    find(out, byLabel('Slide width'))[0].props.onChange({ target: { value: '50' } });
    find(render(), byLabel('Slide height'))[0].props.onChange({ target: { value: '801.4' } });
    out = render();
    expect(find(out, byLabel('Slide width'))[0].props.value).toBe('50');
    find(out, byLabel('Apply slide size'))[0].props.onClick();
    expect(setSlideSize).toHaveBeenLastCalledWith({ w: 100, h: 801 }, { scaleContent: true });
    find(render(), byLabel('Slide width'))[0].props.onChange({ target: { value: '' } });
    find(render(), byLabel('Apply slide size'))[0].props.onClick();
    expect(setSlideSize).toHaveBeenLastCalledWith({ w: 1000, h: 801 }, { scaleContent: true });
    find(render(), byLabel('Slide height'))[0].props.onChange({ target: { value: '  ' } });
    find(render(), byLabel('Apply slide size'))[0].props.onClick();
    expect(setSlideSize).toHaveBeenLastCalledWith({ w: 1000, h: 700 }, { scaleContent: true });
  });

  it('the inputs follow the deck when its size changes', () => {
    render();
    size = { w: 960, h: 720 };
    render(); // the first render after the change schedules the sync, the second shows it
    const again = render();
    expect(find(again, byLabel('Slide width'))[0].props.value).toBe('960');
    expect(find(again, byLabel('Slide size'))[0].props.value).toBe('4:3');
  });
});
