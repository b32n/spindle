import * as Y from 'yjs';
import { InMemoryProvider, __resetInMemoryRooms, type CollabIdentity } from '@b32nio/spindle-shared';
import { DeckImpl } from './deck';
import {
  SLIDE_SIZE_PRESETS,
  MIN_SLIDE_SIDE,
  MAX_SLIDE_SIDE,
  detectSlideSizePreset,
  clampSlideSize,
  scaleElementPatch,
  scaleLayout,
} from './slide-size';
import { resolveLevelStyle, buildPlaceholderElement } from './theme/materialize';
import type { LayoutData } from './theme/types';
import type { LineElement, ShapeElement, TableElement, TextElement } from './scene/types';
import {
  COLLAB_SCHEMA_VERSION,
  CollabSchemaVersionError,
  assertSchemaVersion,
  createSlideYMap,
  hydrateYDocFromData,
  serializeYDocToData,
  yMapToSlide,
} from './collab/y-schema';
import { getCollabBlockers } from './collab/blockers';

const identity = (name: string): CollabIdentity => ({ userId: name, displayName: name, color: '#123456' });

const logo: ShapeElement = {
  id: 'logo', containerId: '', type: 'shape', index: 'a0', x: 1100, y: 640, w: 120, h: 40, rotation: 0, shape: 'rect',
  fill: { kind: 'solid', color: { kind: 'rgb', hex: '#ff0000' } },
} as unknown as ShapeElement;

const layout: LayoutData = {
  id: 'brand',
  name: 'Branded title',
  masterName: 'Corporate master',
  background: { kind: 'solid', color: { kind: 'rgb', hex: '#eeeeee' } },
  placeholders: [
    { type: 'title', idx: 0, frame: { x: 100, y: 100, w: 800, h: 120, rotation: 0 }, style: { fontSize: 44 } },
    { type: 'body', idx: 1, frame: { x: 100, y: 260, w: 800, h: 300, rotation: 0 } },
  ],
  elements: [logo],
  textStyles: {
    title: [{ fontSize: 40, bold: true, align: 'center' }],
    body: [
      { fontSize: 28, bullet: { kind: 'char', char: '•' }, marginLeft: 27, firstLineIndent: -27, lineHeight: 1.2, spaceBefore: 8 },
      { fontSize: 24, marginLeft: 54 },
    ],
  },
};

function deckWithLayout(): DeckImpl {
  const d = new DeckImpl('d', 'D');
  d.setData({ ...d.getData(), layouts: [...d.getLayouts(), layout] });
  return d;
}

describe('slide size presets', () => {
  it('has the six presets with unique ids and in-range sides', () => {
    expect(SLIDE_SIZE_PRESETS.map((p) => p.id)).toEqual(['16:9', '16:10', '4:3', 'a4', 'a4-portrait', 'portrait']);
    for (const p of SLIDE_SIZE_PRESETS) {
      expect(p.w).toBeGreaterThanOrEqual(MIN_SLIDE_SIDE);
      expect(p.h).toBeLessThanOrEqual(MAX_SLIDE_SIDE);
    }
  });

  it('detects a preset within a pixel, otherwise custom', () => {
    for (const p of SLIDE_SIZE_PRESETS) expect(detectSlideSizePreset(p)).toBe(p.id);
    expect(detectSlideSizePreset({ w: 1280.4, h: 719.6 })).toBe('16:9');
    expect(detectSlideSizePreset({ w: 1000, h: 700 })).toBe('custom');
  });

  it('clamps, rounds and falls back for non-finite sides', () => {
    const fb = { w: 1280, h: 720 };
    expect(clampSlideSize({ w: 10, h: 99999 }, fb)).toEqual({ w: MIN_SLIDE_SIDE, h: MAX_SLIDE_SIDE });
    expect(clampSlideSize({ w: 800.6, h: 600.4 }, fb)).toEqual({ w: 801, h: 600 });
    expect(clampSlideSize({ w: NaN, h: Infinity }, fb)).toEqual({ w: 1280, h: 720 });
  });
});

describe('scaling', () => {
  it('scales a frame per axis and the line ends and table row minimums', () => {
    const base = { id: 'x', containerId: 's', index: 'a0', x: 100, y: 200, w: 300, h: 400, rotation: 0 };
    expect(scaleElementPatch({ ...base, type: 'text' } as TextElement, 0.5, 2)).toEqual({ x: 50, y: 400, w: 150, h: 800 });
    const line = { ...base, type: 'line', startPoint: { x: 10, y: 20 }, endPoint: { x: 30, y: 40 } } as unknown as LineElement;
    expect(scaleElementPatch(line, 2, 3)).toMatchObject({ startPoint: { x: 20, y: 60 }, endPoint: { x: 60, y: 120 } });
    const bare = { ...base, type: 'line' } as unknown as LineElement;
    expect(scaleElementPatch(bare, 2, 3)).not.toHaveProperty('startPoint');
    const table = { ...base, type: 'table', rowHeights: [10, 20] } as unknown as TableElement;
    expect(scaleElementPatch(table, 2, 3)).toMatchObject({ rowHeights: [30, 60] });
    const noRows = { ...base, type: 'table' } as unknown as TableElement;
    expect(scaleElementPatch(noRows, 2, 3)).not.toHaveProperty('rowHeights');
  });

  it('scales a layout, with and without static elements', () => {
    const scaled = scaleLayout(layout, 0.75, 1);
    expect(scaled.placeholders[0].frame).toEqual({ x: 75, y: 100, w: 600, h: 120, rotation: 0 });
    expect(scaled.elements![0]).toMatchObject({ x: 825, w: 90, y: 640 });
    expect(layout.placeholders[0].frame.x).toBe(100); // input untouched
    const { elements: _e, ...bare } = layout;
    expect(scaleLayout(bare, 2, 2)).not.toHaveProperty('elements');
  });
});

describe('level text styles', () => {
  it('inherits from lower levels and clamps the level', () => {
    const t = layout.textStyles;
    expect(resolveLevelStyle(t, 'body', 0)).toMatchObject({ fontSize: 28, marginLeft: 27 });
    expect(resolveLevelStyle(t, 'body', 1)).toMatchObject({ fontSize: 24, marginLeft: 54, lineHeight: 1.2 });
    expect(resolveLevelStyle(t, 'body', 7)).toMatchObject({ fontSize: 24 }); // past the defined levels
    expect(resolveLevelStyle(t, 'body', -3)).toMatchObject({ fontSize: 28 });
    expect(resolveLevelStyle(t, 'body', NaN)).toMatchObject({ fontSize: 28 });
    expect(resolveLevelStyle(t, 'subtitle', 0)).toEqual({});
    expect(resolveLevelStyle(undefined, 'body', 0)).toEqual({});
  });

  it('bakes level 0 into a materialized placeholder, the placeholder own style winning', () => {
    const title = buildPlaceholderElement(layout.placeholders[0], 's', 'a0', layout.textStyles);
    expect(title.bodyStyle).toMatchObject({ fontSize: 44, bold: true }); // own 44 over level 0's 40
    expect(title.richText!.content[0].attrs).toMatchObject({ align: 'center' });
    const body = buildPlaceholderElement(layout.placeholders[1], 's', 'a1', layout.textStyles);
    expect(body.richText!.content[0].attrs).toMatchObject({
      bullet: { kind: 'char', char: '•' }, marginLeft: 27, firstLineIndent: -27, lineHeight: 1.2, spaceBefore: 8,
    });
    const plain = buildPlaceholderElement(layout.placeholders[1], 's', 'a1');
    expect(plain.richText!.content[0].attrs ?? {}).not.toHaveProperty('bullet');
  });

  it('addSlide with a layout applies its text styles and background', () => {
    const d = deckWithLayout();
    const s = d.addSlide({ layoutId: 'brand' });
    expect(s.layoutRef).toBe('brand');
    expect(d.getSlide(s.id)!.background).toEqual(layout.background);
    const els = d.getElementsForSlide(s.id) as TextElement[];
    expect(els).toHaveLength(2);
    expect(els[1].richText!.content[0].attrs).toMatchObject({ marginLeft: 27 });
    expect(d.getLayout('brand')!.elements).toHaveLength(1); // static content stays on the layout, not on the slide
  });
});

describe('setSlideSize', () => {
  it('only changes the size without scaleContent', () => {
    const d = new DeckImpl('d');
    const id = d.addElement(d.getActiveSlideId(), { type: 'text', x: 100, y: 100, w: 200, h: 100 }).id;
    d.setSlideSize({ w: 960, h: 720 });
    expect(d.getSlideSize()).toEqual({ w: 960, h: 720 });
    expect(d.getElement(id)).toMatchObject({ x: 100, w: 200 });
  });

  it('scales elements and layouts as one undo step when asked', () => {
    const d = deckWithLayout();
    const sid = d.addSlide({ layoutId: 'brand' }).id;
    const t = d.addElement(sid, { type: 'text', x: 640, y: 360, w: 320, h: 180 }).id;
    const seen: string[] = [];
    d.on('elementChange', (e) => seen.push((e.payload as { elementId: string }).elementId));
    d.setSlideSize({ w: 960, h: 720 }, { scaleContent: true });
    expect(d.getSlideSize()).toEqual({ w: 960, h: 720 });
    expect(d.getElement(t)).toMatchObject({ x: 480, y: 360, w: 240, h: 180 });
    expect(seen).toContain(t);
    expect(d.getLayout('brand')!.elements![0]).toMatchObject({ x: 825, w: 90 });
    expect(d.getLayout('brand')!.placeholders[0].frame.x).toBe(75);
    d.undo();
    expect(d.getSlideSize()).toEqual({ w: 1280, h: 720 });
    expect(d.getElement(t)).toMatchObject({ x: 640, w: 320 });
    expect(d.getLayout('brand')!.elements![0].x).toBe(1100);
    expect(d.canUndo()).toBe(true); // earlier edits are still separate steps
  });

  it('is a plain size change when the size is the same or the current size is degenerate', () => {
    const d = new DeckImpl('d');
    const id = d.addElement(d.getActiveSlideId(), { type: 'text', x: 10, y: 10, w: 20, h: 20 }).id;
    d.setSlideSize({ w: 1280, h: 720 }, { scaleContent: true });
    expect(d.getElement(id)).toMatchObject({ x: 10, w: 20 });
    d.setData({ ...d.getData(), slideSize: { w: 0, h: 0 } });
    d.setSlideSize({ w: 960, h: 720 }, { scaleContent: true });
    expect(d.getSlideSize()).toEqual({ w: 960, h: 720 });
  });
});

describe('hidden slides', () => {
  it('toggles, emits, undoes, ignores no-ops and unknown slides, and serialises', () => {
    const d = new DeckImpl('d');
    const id = d.getActiveSlideId();
    const keys: string[][] = [];
    d.on('slideChange', (e) => keys.push((e.payload as { keys: string[] }).keys));
    d.setSlideHidden(id, false); // already visible
    d.setSlideHidden('nope', true);
    expect(keys).toEqual([]);
    d.setSlideHidden(id, true);
    expect(d.getSlide(id)!.hidden).toBe(true);
    expect(d.getData().slides[0].hidden).toBe(true);
    expect(keys).toEqual([['hidden']]);
    d.undo();
    expect(d.getSlide(id)!.hidden).toBeUndefined();
    d.redo();
    d.setSlideHidden(id, false);
    expect(d.getSlide(id)).not.toHaveProperty('hidden');
  });

  it('a duplicate of a hidden slide is hidden too', () => {
    const d = new DeckImpl('d');
    d.setSlideHidden(d.getActiveSlideId(), true);
    const copy = d.duplicateSlide(d.getActiveSlideId())!;
    expect(copy.hidden).toBe(true);
  });
});

describe('collaboration layout (schema v4)', () => {
  it('has no blockers and round-trips layouts (elements, textStyles, masterName) and hidden', () => {
    const d = deckWithLayout();
    const sid = d.addSlide({ layoutId: 'brand' }).id;
    d.setSlideHidden(sid, true);
    expect(getCollabBlockers(d.getData())).toEqual([]);
    const ydoc = new Y.Doc();
    hydrateYDocFromData(ydoc, d.getData());
    expect(ydoc.getMap('meta').get('schemaVersion')).toBe(COLLAB_SCHEMA_VERSION);
    const back = serializeYDocToData(ydoc, { elementIds: [] }, sid);
    expect(back.layouts!.find((l) => l.id === 'brand')).toEqual(layout);
    expect(back.slides.find((s) => s.id === sid)!.hidden).toBe(true);
    expect(back.slides[0].hidden).toBeUndefined();
  });

  it('reads hidden only when it is literally true', () => {
    const slides = new Y.Doc().getMap<Y.Map<unknown>>('slides');
    slides.set('on', createSlideYMap({ id: 'on', index: 'a0', hidden: true }));
    slides.set('off', createSlideYMap({ id: 'off', index: 'a0' }));
    const odd = new Y.Map<unknown>();
    slides.set('odd', odd);
    odd.set('index', 'a0');
    odd.set('hidden', 'yes');
    expect(yMapToSlide('on', slides.get('on')!).hidden).toBe(true);
    expect(yMapToSlide('odd', odd).hidden).toBeUndefined();
    expect(slides.get('off')!.has('hidden')).toBe(false);
  });

  it('a v3 client refuses a v4 room (refuse-newer-room)', () => {
    expect(COLLAB_SCHEMA_VERSION).toBeGreaterThanOrEqual(4);
    const ydoc = new Y.Doc();
    hydrateYDocFromData(ydoc, deckWithLayout().getData());
    expect(ydoc.getMap('meta').get('schemaVersion')).toBe(COLLAB_SCHEMA_VERSION);
    expect(() => assertSchemaVersion(ydoc, 3)).toThrow(CollabSchemaVersionError);
    expect(() => assertSchemaVersion(ydoc, 3)).toThrow(/up to version 3/);
    expect(() => assertSchemaVersion(ydoc, COLLAB_SCHEMA_VERSION)).not.toThrow();
  });

  it('syncs hidden, slide size with scaled content, and layouts between two clients', async () => {
    __resetInMemoryRooms();
    const a = deckWithLayout();
    const sid = a.addSlide({ layoutId: 'brand' }).id;
    const ha = await a.attachCollab(new InMemoryProvider(), identity('A'), { roomId: 'room' });
    const b = new DeckImpl('b');
    const hb = await b.attachCollab(new InMemoryProvider(), identity('B'), { roomId: 'room' });
    expect(b.getSlide(sid)!.hidden).toBeUndefined();
    expect(b.getLayout('brand')!.elements).toHaveLength(1);

    a.setSlideHidden(sid, true);
    expect(b.getSlide(sid)!.hidden).toBe(true);
    a.setSlideHidden(sid, false);
    expect(b.getSlide(sid)).not.toHaveProperty('hidden');

    const t = a.addElement(sid, { type: 'text', x: 640, y: 360, w: 320, h: 180 }).id;
    a.setSlideSize({ w: 960, h: 720 }, { scaleContent: true });
    expect(b.getSlideSize()).toEqual({ w: 960, h: 720 });
    expect(b.getElement(t)).toMatchObject({ x: 480, w: 240 });
    expect(b.getLayout('brand')!.elements![0].x).toBe(825);
    ha.detach();
    hb.detach();
  });
});
