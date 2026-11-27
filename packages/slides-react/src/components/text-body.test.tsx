import React from 'react';
import { renderToStaticMarkup } from 'react-dom/server';
import { DeckImpl, getBuiltinTheme, type BodyStyle, type RichTextDoc } from '@b32nio/spindle-slides-core';
import { DeckProvider } from '../context/DeckContext';
import { SlideView } from './SlideView';
import { StaticRichText } from './elements/StaticRichText';
import { bodyBoxStyle, bodyOverflow, centresBlock, columnVars, columnsStyle, scaleVars } from './elements/body';
import { bodyDefaults } from './elements/TextView';
import { TextBodyFields, DraftNumber, AUTOFIT_OPTIONS } from './TextBodyFields';
import {
  createFitController,
  endSession,
  fitPlugin,
  measureContentHeight,
  measureShrink,
  overflows,
  persistFit,
  refitElement,
  registerTextBody,
  textBodyRef,
  withShrink,
} from '../text-fit';

const theme = getBuiltinTheme();
type P = Record<string, any>; // eslint-disable-line @typescript-eslint/no-explicit-any
const find = (node: React.ReactNode, pred: (e: React.ReactElement<P>) => boolean, out: React.ReactElement<P>[] = []) => {
  React.Children.forEach(node, (c) => {
    if (!React.isValidElement<P>(c)) return;
    if (pred(c)) out.push(c);
    find(c.props.children, pred, out);
  });
  return out;
};
const byTestId = (id: string) => (e: React.ReactElement<P>) => e.props['data-testid'] === id;
const doc: RichTextDoc = {
  type: 'doc',
  content: [{ type: 'paragraph', attrs: { lineHeight: 1.5 }, content: [{ type: 'text', text: 'Hi', marks: [{ type: 'fontSize', attrs: { size: 24 } }] }] }],
};

describe('body layout helpers', () => {
  it('centresBlock: a centred block shrinks to its content, except with columns', () => {
    expect(centresBlock(undefined)).toBe(false);
    expect(centresBlock({ anchorCtr: true })).toBe(true);
    expect(centresBlock({ anchorCtr: true, columns: { count: 2, gap: 4 } })).toBe(false);
  });
  it('bodyBoxStyle: insets, anchor, centred block and writing mode', () => {
    expect(bodyBoxStyle(undefined)).toMatchObject({ padding: '8px 8px 8px 8px', justifyContent: 'flex-start', display: 'flex', flexDirection: 'column' });
    expect(bodyBoxStyle({ padding: 2 }, 'middle')).toMatchObject({ padding: '2px 2px 2px 2px', justifyContent: 'center' });
    const s = bodyBoxStyle({ insets: { l: 1, t: 2, r: 3, b: 4 }, vAlign: 'bottom', anchorCtr: true, vert: 'eaVert' });
    expect(s).toMatchObject({ padding: '2px 3px 4px 1px', justifyContent: 'flex-end', alignItems: 'center', writingMode: 'vertical-rl', textOrientation: 'mixed' });
    expect(bodyBoxStyle({ vAlign: 'weird' as never })).toMatchObject({ justifyContent: 'flex-start' });
    expect(bodyBoxStyle(undefined)).not.toHaveProperty('alignItems');
  });
  it('scale and column variables', () => {
    expect(scaleVars({ autofit: { mode: 'shrink', fontScale: 0.625, lnSpcReduction: 0.2 } })).toEqual({ '--spindle-fs': '0.625', '--spindle-ls': '0.8' });
    expect(scaleVars(undefined)).toEqual({ '--spindle-fs': '1', '--spindle-ls': '1' });
    expect(columnVars({ columns: { count: 3, gap: 10 } })).toEqual({ '--spindle-cols': '3', '--spindle-colgap': '10px' });
    expect(columnVars(undefined)).toEqual({ '--spindle-cols': '1', '--spindle-colgap': '0px' });
  });
  it('overflow is visible only for a body that declares autofit; columns only from two', () => {
    expect(bodyOverflow(undefined)).toBe('hidden');
    expect(bodyOverflow({ autofit: { mode: 'none' } })).toBe('visible');
    expect(columnsStyle(undefined)).toBeUndefined();
    expect(columnsStyle({ columns: { count: 1 } })).toBeUndefined();
    expect(columnsStyle({ columns: { count: 2, gap: 6 } })).toMatchObject({ columnCount: 2, columnGap: 6 });
  });
  it('bodyDefaults routes the size through the scale variable only under shrink', () => {
    expect(bodyDefaults({ fontSize: 20 }, theme).fontSize).toBe(20);
    expect(bodyDefaults(undefined, theme).fontSize).toBe(18);
    expect(bodyDefaults({ fontSize: 20, autofit: { mode: 'shrink' } }, theme).fontSize).toBe('calc(20px * var(--spindle-fs, 1))');
    expect(bodyDefaults({ autofit: { mode: 'shrink' } }, theme).fontSize).toBe('calc(18px * var(--spindle-fs, 1))');
  });
});

describe('StaticRichText scaling', () => {
  it('is unchanged without scaling', () => {
    const h = renderToStaticMarkup(<StaticRichText doc={doc} theme={theme} />);
    expect(h).toContain('font-size:24px');
    expect(h).toContain('line-height:1.5');
  });
  it('routes sizes and line heights through the variables when scaled', () => {
    const h = renderToStaticMarkup(<StaticRichText doc={doc} theme={theme} scaled />);
    expect(h).toContain('font-size:calc(24px * var(--spindle-fs, 1))');
    expect(h).toContain('line-height:calc(1.5 * var(--spindle-ls, 1))');
    const plain = renderToStaticMarkup(<StaticRichText doc={{ type: 'doc', content: [{ type: 'paragraph' }] }} theme={theme} scaled />);
    expect(plain).toContain('line-height:calc(1.2 * var(--spindle-ls, 1))');
  });
});

function renderDeck(setup: (d: DeckImpl, slideId: string) => void, interactive = false): string {
  const deck = new DeckImpl('d', 'Deck');
  const sid = deck.getActiveSlideId();
  setup(deck, sid);
  return renderToStaticMarkup(
    <DeckProvider deck={deck}>
      <SlideView slideId={sid} interactive={interactive} />
    </DeckProvider>,
  );
}

describe('rendered text bodies', () => {
  beforeAll(() => {
    jest.spyOn(React, 'useSyncExternalStore').mockImplementation((_sub, getSnapshot) => getSnapshot());
  });
  afterAll(() => jest.restoreAllMocks());
  const text = (bodyStyle?: BodyStyle) => (d: DeckImpl, s: string) => {
    d.addElement(s, { type: 'text', x: 0, y: 0, w: 200, h: 80, richText: doc, ...(bodyStyle ? { bodyStyle } : {}) });
  };
  it('default text box renders as before: clipped, 8px padding on every side, no scale variables', () => {
    const h = renderDeck(text());
    expect(h).toContain('overflow:hidden');
    expect(h).toContain('padding:8px 8px 8px 8px');
    expect(h).not.toContain('--spindle-fs');
    expect(h).not.toContain('calc(');
  });
  it('insets, vertical text, centred block and unclipped overflow', () => {
    const h = renderDeck(text({ insets: { l: 1, t: 2, r: 3, b: 4 }, vert: 'vert', anchorCtr: true, autofit: { mode: 'resize' } }));
    expect(h).toContain('padding:2px 3px 4px 1px');
    expect(h).toContain('writing-mode:vertical-rl');
    expect(h).toContain('text-orientation:sideways');
    expect(h).toContain('align-items:center');
    expect(h).toContain('overflow:visible');
    expect(h).toContain('data-text-body');
  });
  it('shrink applies the stored scales through variables', () => {
    const h = renderDeck(text({ fontSize: 20, autofit: { mode: 'shrink', fontScale: 0.5, lnSpcReduction: 0.1 } }));
    expect(h).toContain('--spindle-fs:0.5');
    expect(h).toContain('--spindle-ls:0.9');
    expect(h).toContain('font-size:calc(24px * var(--spindle-fs, 1))');
    expect(h).toContain('font-size:calc(20px * var(--spindle-fs, 1))');
  });
  it('columns wrap the text in a multi-column box', () => {
    const h = renderDeck(text({ columns: { count: 2, gap: 12 } }));
    expect(h).toContain('column-count:2');
    expect(h).toContain('column-gap:12px');
  });
  it('an empty placeholder with columns still renders its prompt and shapes use the same body layout', () => {
    const h = renderDeck((d, s) => {
      d.addElement(s, {
        type: 'shape', shape: 'rect', x: 0, y: 0, w: 100, h: 50, richText: doc,
        bodyStyle: { insets: { l: 5, t: 5, r: 5, b: 5 }, autofit: { mode: 'shrink', fontScale: 0.8 }, columns: { count: 2 } },
      });
      d.addElement(s, { type: 'shape', shape: 'rect', x: 0, y: 0, w: 100, h: 50, richText: doc });
    });
    expect(h).toContain('padding:5px 5px 5px 5px');
    expect(h).toContain('--spindle-fs:0.8');
    expect(h).toContain('column-count:2');
    // the plain shape keeps centred, clipped text
    expect(h).toContain('justify-content:center');
    expect(h).toContain('overflow:hidden');
  });
});

describe('DraftNumber', () => {
  let slots: unknown[] = [];
  let n = 0;
  beforeEach(() => {
    slots = [];
    jest.spyOn(React, 'useState').mockImplementation(((init: unknown) => {
      const i = n++;
      if (!(i in slots)) slots[i] = init;
      return [slots[i], (v: unknown) => { slots[i] = v; }];
    }) as never);
  });
  afterEach(() => jest.restoreAllMocks());
  const render = (value: number, onValue: (n: number) => void) => {
    n = 0;
    return DraftNumber({ value, onValue, id: 'x', min: 0 }) as React.ReactElement<P>;
  };
  it('shows the stored value, keeps a cleared field empty, reports only finite numbers, resets on blur', () => {
    const onValue = jest.fn();
    expect(render(5, onValue).props).toMatchObject({ value: '5', type: 'number', id: 'x', min: 0 });
    render(5, onValue).props.onChange({ target: { value: '', valueAsNumber: NaN } });
    expect(onValue).not.toHaveBeenCalled();
    expect(render(5, onValue).props.value).toBe(''); // the user can retype
    render(5, onValue).props.onChange({ target: { value: '7', valueAsNumber: 7 } });
    expect(onValue).toHaveBeenCalledWith(7);
    expect(render(5, onValue).props.value).toBe('7');
    render(5, onValue).props.onBlur();
    expect(render(5, onValue).props.value).toBe('5');
  });
});

describe('TextBodyFields', () => {
  it('shows the current values with labels', () => {
    const tree = TextBodyFields({ body: { insets: { l: 1, t: 2, r: 3, b: 4 }, autofit: { mode: 'shrink' }, vert: 'vert', columns: { count: 3, gap: 7 }, anchorCtr: true }, onChange: jest.fn() });
    expect(find(tree, byTestId('text-autofit'))[0].props.value).toBe('shrink');
    expect(find(tree, byTestId('text-vert'))[0].props.value).toBe('vert');
    expect(find(tree, byTestId('text-inset-t'))[0].props.value).toBe(2);
    expect(find(tree, byTestId('text-columns'))[0].props.value).toBe(3);
    expect(find(tree, byTestId('text-column-gap'))[0].props.value).toBe(7);
    expect(find(tree, byTestId('text-column-gap'))[0].props.disabled).toBe(false);
    expect(find(tree, byTestId('text-anchor-ctr'))[0].props.checked).toBe(true);
    expect(find(tree, (e) => e.type === 'label').length).toBe(9);
    expect(AUTOFIT_OPTIONS).toHaveLength(3);
    const defaults = TextBodyFields({ body: undefined, onChange: jest.fn() });
    expect(find(defaults, byTestId('text-autofit'))[0].props.value).toBe('none');
    expect(find(defaults, byTestId('text-column-gap'))[0].props.disabled).toBe(true);
    expect(find(defaults, byTestId('text-anchor-ctr'))[0].props.checked).toBe(false);
  });
  it('offers resizing only to text boxes, but keeps an imported choice on a shape', () => {
    const values = (tree: React.ReactNode) => find(tree, byTestId('text-autofit'))[0].props.children.map((o: React.ReactElement<P>) => o.props.value);
    expect(values(TextBodyFields({ body: undefined, onChange: jest.fn() }))).toEqual(['none', 'shrink', 'resize']);
    expect(values(TextBodyFields({ body: undefined, onChange: jest.fn(), allowResize: false }))).toEqual(['none', 'shrink']);
    expect(values(TextBodyFields({ body: { autofit: { mode: 'resize' } }, onChange: jest.fn(), allowResize: false }))).toEqual(['none', 'shrink', 'resize']);
  });
  it('reports each change as a patch', () => {
    const onChange = jest.fn();
    const body: BodyStyle = { padding: 4, columns: { count: 2, gap: 5 } };
    const tree = TextBodyFields({ body, onChange });
    const ev = (extra: object) => ({ target: extra }) as never;
    find(tree, byTestId('text-autofit'))[0].props.onChange(ev({ value: 'resize' }));
    expect(onChange).toHaveBeenLastCalledWith({ autofit: { mode: 'resize' } });
    find(tree, byTestId('text-vert'))[0].props.onChange(ev({ value: 'eaVert' }));
    expect(onChange).toHaveBeenLastCalledWith({ vert: 'eaVert' });
    find(tree, byTestId('text-vert'))[0].props.onChange(ev({ value: 'horz' }));
    expect(onChange).toHaveBeenLastCalledWith({ vert: undefined });
    find(tree, byTestId('text-inset-l'))[0].props.onValue(12);
    expect(onChange).toHaveBeenLastCalledWith({ insets: { l: 12, t: 4, r: 4, b: 4 } });
    find(tree, byTestId('text-inset-r'))[0].props.onValue(-3);
    expect(onChange).toHaveBeenLastCalledWith({ insets: { l: 4, t: 4, r: 0, b: 4 } });
    find(tree, byTestId('text-columns'))[0].props.onValue(3);
    expect(onChange).toHaveBeenLastCalledWith({ columns: { count: 3, gap: 5 } });
    find(tree, byTestId('text-columns'))[0].props.onValue(1);
    expect(onChange).toHaveBeenLastCalledWith({ columns: undefined });
    find(tree, byTestId('text-column-gap'))[0].props.onValue(9);
    expect(onChange).toHaveBeenLastCalledWith({ columns: { count: 2, gap: 9 } });
    find(tree, byTestId('text-column-gap'))[0].props.onValue(-2);
    expect(onChange).toHaveBeenLastCalledWith({ columns: { count: 2, gap: 0 } });
    find(tree, byTestId('text-anchor-ctr'))[0].props.onChange(ev({ checked: true }));
    expect(onChange).toHaveBeenLastCalledWith({ anchorCtr: true });
    find(tree, byTestId('text-anchor-ctr'))[0].props.onChange(ev({ checked: false }));
    expect(onChange).toHaveBeenLastCalledWith({ anchorCtr: undefined });
  });
});

// ── DOM measurement, against a fake box whose content height scales with the font-scale variable ───────────────────

interface FakeBox extends HTMLElement {
  vars: Record<string, string>;
}
/** A box `limit` px tall (or wide) whose content needs `natural * fontScale` px. */
function fakeBox(natural: number, limit: number, opts: { vertical?: boolean; connected?: boolean } = {}): FakeBox {
  const vars: Record<string, string> = {};
  const style = {
    justifyContent: 'center',
    height: '100%',
    setProperty: (k: string, v: string) => { vars[k] = v; },
  } as unknown as CSSStyleDeclaration;
  const need = () => natural * Number(vars['--spindle-fs'] ?? 1);
  const box = {
    vars,
    style,
    isConnected: opts.connected ?? true,
    get scrollHeight() { return opts.vertical ? limit : need(); },
    get clientHeight() { return limit; },
    get scrollWidth() { return opts.vertical ? need() : limit; },
    get clientWidth() { return limit; },
    get offsetHeight() { return style.height === 'auto' ? need() + 16.4 : limit; },
  };
  return box as unknown as FakeBox;
}

describe('text-fit measurement', () => {
  it('overflows compares along the flow axis with 1px of slack', () => {
    expect(overflows(fakeBox(100, 100), false)).toBe(false);
    expect(overflows(fakeBox(101.5, 100), false)).toBe(true);
    expect(overflows(fakeBox(150, 100, { vertical: true }), true)).toBe(true);
    expect(overflows(fakeBox(50, 100, { vertical: true }), true)).toBe(false);
  });
  it('measureShrink finds the fitting scale, leaves it applied, and restores the anchor', () => {
    const box = fakeBox(200, 100);
    const r = measureShrink(box);
    expect(r.fontScale).toBeCloseTo(0.5, 6);
    expect(r.lnSpcReduction).toBe(0.2);
    expect(box.vars['--spindle-fs']).toBe('0.5');
    expect(box.vars['--spindle-ls']).toBe('0.8');
    expect(box.style.justifyContent).toBe('center');
    // text that already fits keeps scale 1; vertical text measures its width
    expect(measureShrink(fakeBox(80, 100)).fontScale).toBe(1);
    expect(measureShrink(fakeBox(300, 100, { vertical: true }), true).fontScale).toBeCloseTo(0.325, 6);
  });
  it('measureContentHeight releases the height and restores it', () => {
    const box = fakeBox(100, 40);
    expect(measureContentHeight(box)).toBe(117);
    expect(box.style.height).toBe('100%');
  });
  it('withShrink keeps other body attributes', () => {
    expect(withShrink({ padding: 3, autofit: { mode: 'shrink' } }, { fontScale: 0.9, lnSpcReduction: 0.1 })).toEqual({
      padding: 3,
      autofit: { mode: 'shrink', fontScale: 0.9, lnSpcReduction: 0.1 },
    });
    expect(withShrink(undefined, { fontScale: 1, lnSpcReduction: 0 })).toEqual({ autofit: { mode: 'shrink', fontScale: 1, lnSpcReduction: 0 } });
  });
});

function fakeDeck(el: Record<string, unknown> | undefined) {
  const updates: Array<Record<string, unknown>> = [];
  return {
    updates,
    getElement: jest.fn(() => el),
    updateElement: jest.fn((_id: string, patch: Record<string, unknown>) => { updates.push(patch); }),
  };
}

describe('persistFit', () => {
  it('ignores a missing element and non text boxes', () => {
    const box = fakeBox(200, 100);
    for (const el of [undefined, { type: 'image' }]) {
      const d = fakeDeck(el);
      persistFit(d as never, 'e', box);
      expect(d.updateElement).not.toHaveBeenCalled();
    }
  });
  it('shrink stores the new scale, and skips the write when nothing changed', () => {
    const d = fakeDeck({ type: 'text', h: 100, bodyStyle: { autofit: { mode: 'shrink' } } });
    persistFit(d as never, 'e', fakeBox(200, 100));
    expect(d.updates).toEqual([{ bodyStyle: { autofit: { mode: 'shrink', fontScale: 0.5, lnSpcReduction: 0.2 } } }]);
    const same = fakeDeck({ type: 'shape', bodyStyle: { autofit: { mode: 'shrink', fontScale: 0.5, lnSpcReduction: 0.2 } }, h: 10 });
    persistFit(same as never, 'e', fakeBox(200, 100));
    expect(same.updateElement).not.toHaveBeenCalled();
    const vert = fakeDeck({ type: 'text', bodyStyle: { vert: 'vert', autofit: { mode: 'shrink' } } });
    persistFit(vert as never, 'e', fakeBox(300, 100, { vertical: true }));
    expect(vert.updates[0]).toEqual({ bodyStyle: { vert: 'vert', autofit: { mode: 'shrink', fontScale: 0.325, lnSpcReduction: 0.2 } } });
  });
  it('resize sets the box height to the content, never to nothing', () => {
    const grow = fakeDeck({ type: 'text', h: 40, bodyStyle: { autofit: { mode: 'resize' } } });
    persistFit(grow as never, 'e', fakeBox(100, 40));
    expect(grow.updates).toEqual([{ h: 117 }]);
    const sameH = fakeDeck({ type: 'text', h: 117, bodyStyle: { autofit: { mode: 'resize' } } });
    persistFit(sameH as never, 'e', fakeBox(100, 40));
    expect(sameH.updateElement).not.toHaveBeenCalled();
    const empty = fakeDeck({ type: 'text', h: 40, bodyStyle: { autofit: { mode: 'resize' } } });
    persistFit(empty as never, 'e', fakeBox(-20, 40)); // content shorter than the insets alone
    expect(empty.updateElement).not.toHaveBeenCalled();
  });
  it('resize does not apply to shapes or vertical text, and none does nothing', () => {
    for (const el of [
      { type: 'shape', h: 40, bodyStyle: { autofit: { mode: 'resize' } } },
      { type: 'text', h: 40, bodyStyle: { vert: 'vert', autofit: { mode: 'resize' } } },
      { type: 'text', h: 40, bodyStyle: { autofit: { mode: 'none' } } },
      { type: 'text', h: 40 },
    ]) {
      const d = fakeDeck(el);
      persistFit(d as never, 'e', fakeBox(100, 40));
      expect(d.updateElement).not.toHaveBeenCalled();
    }
  });
});

describe('stage registry', () => {
  it('refits a registered, attached body; ignores unknown or detached ones', () => {
    const d = fakeDeck({ type: 'text', h: 100, bodyStyle: { autofit: { mode: 'shrink' } } });
    refitElement(d as never, 'nobody');
    expect(d.updateElement).not.toHaveBeenCalled();
    const off = registerTextBody('a', fakeBox(200, 100));
    refitElement(d as never, 'a');
    expect(d.updateElement).toHaveBeenCalledTimes(1);
    off();
    refitElement(d as never, 'a');
    expect(d.updateElement).toHaveBeenCalledTimes(1);
    const detached = registerTextBody('b', fakeBox(200, 100, { connected: false }));
    refitElement(d as never, 'b');
    expect(d.updateElement).toHaveBeenCalledTimes(1);
    detached();
  });
  it('unregistering a replaced node leaves the newer one registered', () => {
    const first = registerTextBody('c', fakeBox(200, 100));
    const second = registerTextBody('c', fakeBox(200, 100));
    first();
    const d = fakeDeck({ type: 'text', h: 100, bodyStyle: { autofit: { mode: 'shrink' } } });
    refitElement(d as never, 'c');
    expect(d.updateElement).toHaveBeenCalledTimes(1);
    second();
  });
  it('textBodyRef registers only interactive bodies', () => {
    const node = fakeBox(200, 100) as unknown as HTMLDivElement;
    expect(textBodyRef('r', false)(node)).toBeUndefined();
    expect(textBodyRef('r', true)(null)).toBeUndefined();
    const off = textBodyRef('r', true)(node)!;
    const d = fakeDeck({ type: 'text', h: 100, bodyStyle: { autofit: { mode: 'shrink' } } });
    refitElement(d as never, 'r');
    expect(d.updateElement).toHaveBeenCalledTimes(1);
    off();
  });
});

describe('fitPlugin and endSession', () => {
  it('reports only real document changes', () => {
    const live = jest.fn();
    const view = (fitPlugin({ live }).spec.view as never as () => { update: (v: unknown, prev: unknown) => void })();
    const same = { doc: { eq: () => true } };
    view.update({ state: same }, same);
    expect(live).not.toHaveBeenCalled();
    view.update({ state: same }, { doc: { eq: () => false } });
    expect(live).toHaveBeenCalledTimes(1);
  });
  it('writes the document first when asked, then flushes the fit', () => {
    const order: string[] = [];
    endSession({ flush: () => order.push('flush') }, () => order.push('write'));
    endSession({ flush: () => order.push('flush') });
    expect(order).toEqual(['write', 'flush', 'flush']);
  });
});

describe('createFitController', () => {
  it('settle leaves a stored scale alone until the text was edited in the session', () => {
    const d = fakeDeck({ type: 'text', h: 100, bodyStyle: { autofit: { mode: 'shrink', fontScale: 0.625, lnSpcReduction: 0.2 } } });
    const c = createFitController(d as never, 'e', fakeBox(200, 100), false);
    c.settle(); // opened and left without typing
    expect(d.updateElement).not.toHaveBeenCalled();
    c.live();
    c.settle();
    expect(d.updateElement).toHaveBeenCalledTimes(1);
    d.updateElement.mockClear();
    c.settle(); // already settled
    expect(d.updateElement).not.toHaveBeenCalled();
  });
  it('live rescales through variables only; flush stores the last live result without measuring', () => {
    const d = fakeDeck({ type: 'text', h: 100, bodyStyle: { autofit: { mode: 'shrink' } } });
    const box = fakeBox(200, 100);
    const c = createFitController(d as never, 'e', box, false);
    c.flush(); // nothing pending
    expect(d.updateElement).not.toHaveBeenCalled();
    c.live();
    expect(box.vars['--spindle-fs']).toBe('0.5');
    expect(d.updateElement).not.toHaveBeenCalled();
    c.flush();
    expect(d.updates).toEqual([{ bodyStyle: { autofit: { mode: 'shrink', fontScale: 0.5, lnSpcReduction: 0.2 } } }]);
    c.flush(); // cleared
    expect(d.updateElement).toHaveBeenCalledTimes(1);
  });
  it('settle measures and stores, and clears the pending result', () => {
    const d = fakeDeck({ type: 'text', h: 100, bodyStyle: { autofit: { mode: 'shrink' } } });
    const c = createFitController(d as never, 'e', fakeBox(200, 100), false);
    c.live();
    c.settle();
    expect(d.updateElement).toHaveBeenCalledTimes(1);
    c.flush();
    expect(d.updateElement).toHaveBeenCalledTimes(1);
  });
  it('does nothing for table cells or when the mode is not shrink', () => {
    const cell = fakeDeck({ type: 'table' });
    const c1 = createFitController(cell as never, 'e', fakeBox(200, 100), true);
    c1.live(); c1.settle(); c1.flush();
    expect(cell.updateElement).not.toHaveBeenCalled();
    const none = fakeDeck({ type: 'text', h: 100, bodyStyle: { autofit: { mode: 'none' } } });
    const box = fakeBox(200, 100);
    const c2 = createFitController(none as never, 'e', box, false);
    c2.live(); c2.flush();
    expect(box.vars['--spindle-fs']).toBeUndefined();
    expect(none.updateElement).not.toHaveBeenCalled();
    // the mode changed away from shrink between a live fit and the flush: nothing is written
    const flip = { type: 'text', h: 100, bodyStyle: { autofit: { mode: 'shrink' } } } as Record<string, unknown>;
    const d = fakeDeck(flip);
    const c3 = createFitController(d as never, 'e', fakeBox(200, 100), false);
    c3.live();
    flip.bodyStyle = {};
    c3.flush();
    expect(d.updateElement).not.toHaveBeenCalled();
    // a vanished element is tolerated
    const gone = createFitController(fakeDeck(undefined) as never, 'e', fakeBox(200, 100), false);
    gone.live(); gone.flush();
  });
});
