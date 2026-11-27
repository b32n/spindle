import React from 'react';
import { renderToStaticMarkup } from 'react-dom/server';
import { DeckImpl, GRADIENT_PRESETS, getBuiltinTheme, type Fill } from '@b32nio/spindle-slides-core';
import { DeckProvider } from '../context/DeckContext';
import { SlideView } from './SlideView';
import { htmlFill, paintId, svgFill } from './elements/fill';
import { FillExtras, readImageAsFill } from './FillExtras';

const theme = getBuiltinTheme();
const PNG = 'data:image/png;base64,iVBORw0KGgo=';
const red = { kind: 'rgb', hex: '#ff0000' } as const;
const blue = { kind: 'rgb', hex: '#0000ff' } as const;
const gradient: Fill = { kind: 'gradient', type: 'linear', angle: 0, stops: [{ pos: 0, color: red }, { pos: 1, color: blue }] };
const radial: Fill = { kind: 'gradient', type: 'radial', stops: [{ pos: 0, color: red }, { pos: 1, color: blue }] };
const picture = (extra: object = {}): Fill => ({ kind: 'picture', src: PNG, mode: 'stretch', ...extra }) as Fill;

describe('htmlFill', () => {
  it('no fill, none and unusable fills paint nothing', () => {
    for (const f of [undefined, { kind: 'none' } as Fill, { kind: 'picture', src: 'https://x/a.png', mode: 'stretch' } as Fill]) {
      expect(htmlFill(f, theme)).toEqual({ style: {}, layer: null });
    }
  });
  it('colour and gradient go straight into the box style', () => {
    expect(htmlFill({ kind: 'solid', color: red }, theme).style).toEqual({ background: 'rgb(255, 0, 0)' });
    const g = htmlFill(gradient, theme);
    expect(g.style).toEqual({ background: 'linear-gradient(90deg, rgb(255, 0, 0) 0%, rgb(0, 0, 255) 100%)' });
    expect(g.layer).toBeNull();
  });
  it('an opaque picture is a plain background; a translucent one gets its own layer', () => {
    const opaque = htmlFill(picture({ alpha: 1 }), theme);
    expect(opaque.layer).toBeNull();
    expect(opaque.style).toMatchObject({ backgroundSize: '100% 100%', backgroundRepeat: 'no-repeat' });
    const t = htmlFill(picture({ alpha: 0.4, mode: 'tile', tile: { w: 8, h: 8 } }), theme);
    expect(t.style).toEqual({ position: 'relative', isolation: 'isolate' });
    const html = renderToStaticMarkup(t.layer!);
    expect(html).toContain('data-fill-layer');
    expect(html).toContain('opacity:0.4');
    expect(html).toContain('background-size:8px 8px');
  });
});

const svgHtml = (n: React.ReactNode): string => renderToStaticMarkup(<svg>{n}</svg>);

describe('svgFill', () => {
  it('ids are selector safe', () => expect(paintId('a b:c.1')).toBe('fill-a_b_c_1'));
  it('none, colour', () => {
    expect(svgFill({ kind: 'none' }, theme, 'p', 10, 10)).toEqual({ defs: null, fill: 'none' });
    expect(svgFill({ kind: 'solid', color: red }, theme, 'p', 10, 10)).toEqual({ defs: null, fill: 'rgb(255, 0, 0)' });
  });
  it('linear gradient in user space', () => {
    const r = svgFill(gradient, theme, 'p', 200, 100);
    expect(r.fill).toBe('url(#p)');
    const html = svgHtml(r.defs!);
    expect(html).toContain('<linearGradient id="p" gradientUnits="userSpaceOnUse" x1="0" y1="50" x2="200" y2="50">');
    expect(html).toContain('<stop offset="0" stop-color="rgb(255, 0, 0)">');
  });
  it('radial and path gradients', () => {
    for (const type of ['radial', 'path'] as const) {
      const html = svgHtml(svgFill({ ...(radial as object), type } as Fill, theme, 'p', 50, 50).defs!);
      expect(html).toContain('<radialGradient id="p" cx="0.5" cy="0.5" r="0.7071">');
    }
  });
  it('picture stretch and tile patterns', () => {
    const s = svgHtml(svgFill(picture({ alpha: 0.5 }), theme, 'p', 30, 20).defs!);
    expect(s).toContain('<pattern id="p" patternUnits="userSpaceOnUse" width="30" height="20">');
    expect(s).toContain('opacity="0.5"');
    const t = svgHtml(svgFill(picture({ mode: 'tile', tile: { w: 6, h: 7 } }), theme, 'p', 30, 20).defs!);
    expect(t).toContain('width="6" height="7"');
  });
});

function render(setup: (d: DeckImpl, slideId: string) => void): string {
  const deck = new DeckImpl('d', 'Deck');
  const sid = deck.getActiveSlideId();
  setup(deck, sid);
  return renderToStaticMarkup(
    <DeckProvider deck={deck}>
      <SlideView slideId={sid} />
    </DeckProvider>,
  );
}

describe('rendered slide', () => {
  // Server rendering has no snapshot source for external stores; read the client snapshot directly.
  beforeAll(() => {
    jest.spyOn(React, 'useSyncExternalStore').mockImplementation((_sub, getSnapshot) => getSnapshot());
  });
  afterAll(() => jest.restoreAllMocks());
  it('uses the theme background when none is set, and for an unusable one', () => {
    expect(render(() => {})).toContain('background:rgb(');
    expect(render((d, s) => d.setSlideBackground(s, picture({ src: 'https://x/a.png' })))).toContain('background:rgb(');
  });
  it('paints a gradient background', () => {
    expect(render((d, s) => d.setSlideBackground(s, gradient))).toContain('linear-gradient(90deg');
  });
  it('paints a translucent picture background on its own layer', () => {
    const html = render((d, s) => d.setSlideBackground(s, picture({ alpha: 0.3 })));
    expect(html).toContain('isolation:isolate');
    expect(html).toContain('data-fill-layer');
  });
  it('paints shape fills through SVG paint servers', () => {
    const html = render((d, s) => { d.addElement(s, { type: 'shape', shape: 'rect', x: 0, y: 0, w: 100, h: 50, fill: gradient }); });
    expect(html).toContain('<linearGradient');
    expect(html).toContain('fill="url(#fill-');
  });
  it('paints text box fills', () => {
    const html = render((d, s) => { d.addElement(s, { type: 'text', x: 0, y: 0, w: 100, h: 50, fill: radial }); });
    expect(html).toContain('radial-gradient(ellipse farthest-corner');
  });
  it('paints table cell fills, including a translucent picture layer', () => {
    const html = render((d, s) => {
      const t = d.addElement(s, { type: 'table', rows: 1, cols: 2, x: 0, y: 0, w: 100, h: 20 });
      d.updateTableCells(t.id, [[0, 0]], { fill: gradient });
      d.updateTableCells(t.id, [[0, 1]], { fill: picture({ alpha: 0.5 }) });
    });
    expect(html).toContain('linear-gradient(90deg');
    expect(html).toContain('data-fill-layer');
  });
  it('pattern fills render as their solid blend', () => {
    const html = render((d, s) => {
      d.addElement(s, { type: 'shape', shape: 'rect', x: 0, y: 0, w: 10, h: 10, fill: { kind: 'pattern', preset: 'pct50', fg: red, bg: blue } });
    });
    expect(html).toContain('fill="rgb(128, 0, 128)"');
  });
});

describe('readImageAsFill', () => {
  class FakeReader {
    result: unknown = null;
    onload: (() => void) | null = null;
    onerror: (() => void) | null = null;
    static outcome: 'ok' | 'error' | 'bad' | 'nonstring' = 'ok';
    readAsDataURL(): void {
      const o = FakeReader.outcome;
      if (o === 'error') this.onerror?.();
      else {
        this.result = o === 'ok' ? PNG : o === 'bad' ? 'data:text/html;base64,AAAA' : new ArrayBuffer(1);
        this.onload?.();
      }
    }
  }
  const g = globalThis as unknown as { FileReader?: unknown };
  beforeAll(() => { g.FileReader = FakeReader; });
  afterAll(() => { delete g.FileReader; });
  const file = {} as File;
  it('turns an image into a stretch picture fill', async () => {
    FakeReader.outcome = 'ok';
    expect(await readImageAsFill(file)).toEqual({ kind: 'picture', src: PNG, mode: 'stretch' });
  });
  it('rejects read errors, non-image data and non-string results', async () => {
    for (const o of ['error', 'bad', 'nonstring'] as const) {
      FakeReader.outcome = o;
      expect(await readImageAsFill(file)).toBeNull();
    }
  });
});

describe('FillExtras controls', () => {
  type P = Record<string, any>; // eslint-disable-line @typescript-eslint/no-explicit-any
  const click = jest.fn();
  const find = (node: React.ReactNode, type: string, out: P[] = []): P[] => {
    React.Children.forEach(node, (c) => {
      if (!React.isValidElement(c)) return;
      const p = c.props as P;
      if (c.type === type) out.push(p);
      find(p.children, type, out);
    });
    return out;
  };
  const build = (props: Partial<React.ComponentProps<typeof FillExtras>> = {}) => {
    const onPick = jest.fn();
    jest.spyOn(React, 'useRef').mockReturnValue({ current: { click } });
    const tree = FillExtras({ onPick, ...props });
    return { onPick, select: find(tree, 'select')[0], button: find(tree, 'button')[0], input: find(tree, 'input')[0] };
  };
  afterEach(() => jest.restoreAllMocks());

  it('a gradient preset is picked; an empty or unknown choice is ignored', () => {
    const { onPick, select } = build();
    select.onChange({ target: { value: 'radial' } });
    expect(onPick).toHaveBeenCalledWith(GRADIENT_PRESETS.find((p) => p.id === 'radial')!.fill);
    select.onChange({ target: { value: '' } });
    select.onChange({ target: { value: 'nope' } });
    expect(onPick).toHaveBeenCalledTimes(1);
    select.onMouseDown({ stopPropagation: jest.fn() });
  });

  it('the image button opens the hidden file input', () => {
    const { button } = build();
    button.onMouseDown({ preventDefault: jest.fn() });
    button.onClick();
    expect(click).toHaveBeenCalledTimes(1);
  });

  it('picking a file applies an image fill; no file or a bad file applies nothing', async () => {
    const g = globalThis as unknown as { FileReader?: unknown };
    g.FileReader = class {
      result: unknown = PNG;
      onload: (() => void) | null = null;
      readAsDataURL(): void { this.onload?.(); }
    };
    const { onPick, input } = build({ disabled: true, testId: 'x', style: { height: 1 } });
    const target = { files: [{}] as unknown as FileList, value: 'c:\\a.png' };
    await input.onChange({ target });
    expect(target.value).toBe('');
    expect(onPick).toHaveBeenCalledWith({ kind: 'picture', src: PNG, mode: 'stretch' });
    await input.onChange({ target: { files: [], value: '' } });
    expect(onPick).toHaveBeenCalledTimes(1);
    g.FileReader = class {
      result: unknown = 'data:text/html,x';
      onload: (() => void) | null = null;
      readAsDataURL(): void { this.onload?.(); }
    };
    await input.onChange({ target: { files: [{}], value: '' } });
    expect(onPick).toHaveBeenCalledTimes(1);
    delete g.FileReader;
  });
});
