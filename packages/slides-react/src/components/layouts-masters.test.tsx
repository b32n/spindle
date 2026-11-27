import React from 'react';
import { renderToStaticMarkup } from 'react-dom/server';
import { DeckImpl, type LayoutData, type ShapeElement } from '@b32nio/spindle-slides-core';
import { DeckProvider } from '../context/DeckContext';
import { SlideView } from './SlideView';
import { Filmstrip } from './Filmstrip';
import { Menu } from './Menu';
import { SlideContextMenu } from './SlideContextMenu';
import { presentableSlideIds, presentStartIndex } from './present-navigation';

type P = Record<string, any>; // eslint-disable-line @typescript-eslint/no-explicit-any

const logo = {
  id: 'logo', containerId: '', type: 'shape', index: 'a0', x: 1100, y: 640, w: 120, h: 40, rotation: 0, shape: 'rect',
  fill: { kind: 'solid', color: { kind: 'rgb', hex: '#ff0000' } },
} as unknown as ShapeElement;
const layout: LayoutData = { id: 'brand', name: 'Brand', placeholders: [], elements: [logo] };

function deckWithBrand(): { deck: DeckImpl; slideId: string } {
  const deck = new DeckImpl('d');
  deck.setData({ ...deck.getData(), layouts: [...deck.getLayouts(), layout] });
  const slideId = deck.addSlide({ layoutId: 'brand' }).id;
  return { deck, slideId };
}

// Server rendering has no snapshot source for external stores; read the client snapshot directly.
beforeAll(() => {
  jest.spyOn(React, 'useSyncExternalStore').mockImplementation((_sub, getSnapshot) => getSnapshot());
});
afterAll(() => jest.restoreAllMocks());

const html = (deck: DeckImpl, node: React.ReactNode) => renderToStaticMarkup(<DeckProvider deck={deck}>{node}</DeckProvider>);

describe('layout static elements', () => {
  it('SlideView draws the layout elements under the slide elements, not selectable', () => {
    const { deck, slideId } = deckWithBrand();
    const el = deck.addElement(slideId, { type: 'text', x: 5, y: 5, w: 50, h: 50 });
    const out = html(deck, <SlideView slideId={slideId} interactive />);
    expect(out).toContain('data-layout-element-id="logo"');
    expect(out).toContain('translate(1100px, 640px)');
    expect(out).toContain('pointer-events:none');
    expect(out.indexOf('data-layout-element-id')).toBeLessThan(out.indexOf(`data-element-id="${el.id}"`));
  });

  it('a slide without a layout, or whose layout has no elements, draws none', () => {
    const { deck, slideId } = deckWithBrand();
    expect(html(deck, <SlideView slideId={deck.getSlideIds()[0]} />)).not.toContain('data-layout-element-id');
    deck.setSlideLayout(slideId, 'blank');
    expect(html(deck, <SlideView slideId={slideId} />)).not.toContain('data-layout-element-id');
    deck.setSlideLayout(slideId, 'gone');
    expect(html(deck, <SlideView slideId={slideId} />)).not.toContain('data-layout-element-id');
  });

  it('a slide size change moves the layout elements with the content', () => {
    const { deck, slideId } = deckWithBrand();
    deck.setSlideSize({ w: 960, h: 720 }, { scaleContent: true });
    expect(html(deck, <SlideView slideId={slideId} />)).toContain('translate(825px, 640px)');
  });
});

describe('hidden slides in the filmstrip', () => {
  it('marks only hidden thumbnails and says so in the accessible name', () => {
    const { deck, slideId } = deckWithBrand();
    expect(html(deck, <Filmstrip />)).not.toContain('slide-hidden-badge');
    deck.setSlideHidden(slideId, true);
    const out = html(deck, <Filmstrip />);
    expect(out.match(/slide-hidden-badge/g)).toHaveLength(1);
    expect(out).toContain('aria-label="Go to slide 2 (hidden)"');
    expect(out).toContain('aria-label="Go to slide 1"');
  });
});

describe('SlideContextMenu', () => {
  const items = (deck: DeckImpl, slideId: string): Array<{ label: string; run: () => void }> => {
    let menu: React.ReactElement<P> | undefined;
    const Probe = () => {
      menu = SlideContextMenu({ slideId, x: 0, y: 0, onClose: () => {} }) as React.ReactElement<P>;
      return null;
    };
    html(deck, <Probe />);
    expect(menu!.type).toBe(Menu);
    return (menu!.props.items as Array<{ label: string; run: () => void } | 'sep'>).filter((i) => i !== 'sep') as never;
  };

  it('offers Hide slide, then Show slide once hidden, and toggles the slide', () => {
    const { deck, slideId } = deckWithBrand();
    const hide = items(deck, slideId).find((i) => i.label === 'Hide slide')!;
    hide.run();
    expect(deck.getSlide(slideId)!.hidden).toBe(true);
    const show = items(deck, slideId).find((i) => i.label === 'Show slide')!;
    expect(items(deck, slideId).some((i) => i.label === 'Hide slide')).toBe(false);
    show.run();
    expect(deck.getSlide(slideId)!.hidden).toBeUndefined();
  });
});

describe('presentable slides', () => {
  const ids = ['a', 'b', 'c', 'd'];
  const hiddenSet = (...h: string[]) => (id: string) => h.includes(id);

  it('skips hidden slides unless asked, and falls back to all when every slide is hidden', () => {
    expect(presentableSlideIds(ids, hiddenSet('b', 'c'))).toEqual(['a', 'd']);
    expect(presentableSlideIds(ids, hiddenSet('b', 'c'), true)).toEqual(ids);
    expect(presentableSlideIds(ids, hiddenSet(...ids))).toEqual(ids);
    expect(presentableSlideIds([], hiddenSet())).toEqual([]);
  });

  it('starts at the active slide, else the next played slide, else the last', () => {
    const played = ['a', 'd'];
    expect(presentStartIndex(ids, played, 'd')).toBe(1);
    expect(presentStartIndex(ids, played, 'b')).toBe(1); // hidden: next played slide is d
    expect(presentStartIndex(ids, ['a', 'b'], 'd')).toBe(1); // nothing after: last
    expect(presentStartIndex(ids, played, 'zzz')).toBe(1); // unknown active slide
    expect(presentStartIndex([], [], 'a')).toBe(0);
  });
});
