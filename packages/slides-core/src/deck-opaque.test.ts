import { hashBytes, type OpaquePart } from '@b32nio/spindle-shared/opaque';
import { DeckImpl } from './deck';
import { createOpaqueElement } from './scene/elements';
import { normalizeDeckData } from './serialization';
import type { DeckData } from './types';
import type { OpaqueElement } from './scene/types';

// Opaque content in the deck (plan K11): elements, parts, the slide's transition and timing, and pruning on delete.

const enc = (s: string): Uint8Array => new TextEncoder().encode(s);
async function part(text: string, contentType = 'application/xml'): Promise<OpaquePart> {
  const bytes = enc(text);
  return { hash: await hashBytes(bytes), contentType, bytes };
}

const effect = (id: number, spid: string): string =>
  `<p:par><p:cTn id="${id}" presetClass="entr"><p:childTnLst><p:set><p:cBhvr><p:cTn id="${id + 1}"/><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl></p:cBhvr></p:set></p:childTnLst></p:cTn></p:par>`;
const TIMING = (...spids: string[]): string =>
  `<p:timing xmlns:p="urn:p"><p:tnLst><p:par><p:cTn id="1"><p:childTnLst>${spids.map((s, i) => effect(10 + i * 2, s)).join('')}</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>`;
const TRANSITION = '<p:transition xmlns:p="urn:p"><p:fade/></p:transition>';

function deckWith(timing: string | null = TIMING('3', '4')): { d: DeckImpl; slide: string; chart: OpaqueElement; box: string; other: string } {
  const d = new DeckImpl('deck', 'Deck');
  const slide = d.getActiveSlideId();
  const data = d.getData();
  const chart = createOpaqueElement({ containerId: slide, index: 'a0', kind: 'chart', partRefs: [{ hash: 'h'.repeat(64), role: 'chart' }], pptxId: '3', title: 'Sales' });
  const box = { id: 'box', containerId: slide, index: 'a1', type: 'text', x: 0, y: 0, w: 10, h: 10, rotation: 0, pptxId: '4', richText: { type: 'doc', content: [{ type: 'paragraph' }] } };
  const other = { ...box, id: 'other', index: 'a2', pptxId: '5' };
  data.slides[0] = { ...data.slides[0], opaque: { transition: TRANSITION, ...(timing ? { timing } : {}) }, elements: [chart, box, other] as never };
  d.setData(data);
  return { d, slide, chart, box: 'box', other: 'other' };
}

describe('opaque elements', () => {
  it('createOpaqueElement keeps what it is given, copies the refs and leaves out what is empty', () => {
    const refs = [{ hash: 'a'.repeat(64), role: 'chart', partName: 'ppt/charts/chart1.xml' }];
    const el = createOpaqueElement({ containerId: 's', index: 'a', kind: 'media', partRefs: refs, alt: 'A clip', poster: 'data:image/png;base64,AA==', props: { clip: 'videoFile' }, pptxId: '9', x: 5 });
    expect(el).toMatchObject({ type: 'opaque', kind: 'media', alt: 'A clip', poster: 'data:image/png;base64,AA==', props: { clip: 'videoFile' }, pptxId: '9', x: 5, w: 320, h: 240 });
    expect(el.partRefs).toEqual(refs);
    expect(el.partRefs[0]).not.toBe(refs[0]);
    const bare = createOpaqueElement({ containerId: 's', index: 'a', kind: 'ink', partRefs: [], props: {} });
    expect(Object.keys(bare).sort()).toEqual(['containerId', 'h', 'id', 'index', 'kind', 'partRefs', 'rotation', 'type', 'w', 'x', 'y']);
  });

  it('can be moved, resized and deleted like any element, and a copy has no PPTX id', () => {
    const { d, slide, chart } = deckWith();
    d.updateElement(chart.id, { x: 50, w: 400 });
    expect(d.getElement(chart.id)).toMatchObject({ x: 50, w: 400, title: 'Sales' });
    const copy = d.duplicateElement(chart.id) as OpaqueElement;
    expect(copy.pptxId).toBeUndefined();
    expect(copy.partRefs).toEqual(chart.partRefs);
    const copies = d.duplicateElements([chart.id, 'box']);
    expect(copies.map((c) => c.pptxId)).toEqual([undefined, undefined]);
    expect(d.getElementsForSlide(slide).filter((e) => e.pptxId === '3')).toHaveLength(1);
  });
});

describe('opaque parts in the deck', () => {
  it('round-trips through getData and setData, skipping malformed entries, and shares the bytes', async () => {
    const a = await part('<chart/>', 'application/vnd.chart+xml');
    const d = new DeckImpl();
    expect(d.getData().opaqueParts).toBeUndefined();
    const data = d.getData();
    d.setData({ ...data, opaqueParts: [a, { hash: 'nope', contentType: 'x', bytes: enc('x') }, { ...a, hash: 'b'.repeat(64), bytes: 'str' as never }, { ...a, hash: 'c'.repeat(64), contentType: 3 as never }, null as never] });
    expect(d.getOpaqueParts()).toEqual([a]);
    expect(d.getOpaquePart(a.hash)).toBe(a);
    expect(d.getOpaquePart('0'.repeat(64))).toBeUndefined();
    expect(d.getData().opaqueParts).toEqual([a]);
    expect(d.getData().opaqueParts?.[0].bytes).toBe(a.bytes);
    d.setData({ ...data });
    expect(d.getOpaqueParts()).toEqual([]);
  });

  it('normalizeDeckData passes them on, and a remote peer\'s parts are added to the deck\'s', async () => {
    const a = await part('<a/>');
    const b = await part('<b/>');
    expect(normalizeDeckData({ ...new DeckImpl().getData(), opaqueParts: [a] }).opaqueParts).toEqual([a]);
    const d = new DeckImpl();
    d.setData({ ...d.getData(), opaqueParts: [a] });
    const seen: unknown[] = [];
    d.on('deckChange', (e) => seen.push(e.payload));
    d._applyRemoteOpaqueParts([b, a]);
    expect(d.getOpaqueParts().map((p) => p.hash).sort()).toEqual([a.hash, b.hash].sort());
    expect(seen).toEqual([{ keys: ['opaqueParts'] }]);
  });
});

describe('the slide\'s transition and timing', () => {
  it('survive getData, setData and duplicating the slide', () => {
    const { d, slide } = deckWith();
    expect(d.getData().slides[0].opaque).toEqual({ transition: TRANSITION, timing: TIMING('3', '4') });
    const copy = d.duplicateSlide(slide) as { id: string };
    expect(d.getSlide(copy.id)?.opaque?.timing).toBe(TIMING('3', '4'));
    expect(d.getElementsForSlide(copy.id).map((e) => e.pptxId)).toEqual(['3', '4', '5']);
  });

  it('deleting a shape removes the effects that animate it, as part of the same undo step', () => {
    const { d, slide, box } = deckWith();
    const changes: unknown[] = [];
    d.on('slideChange', (e) => changes.push(e.payload));
    d.deleteElement(box);
    expect(d.getSlide(slide)?.opaque?.timing).toBe(TIMING('3'));
    expect(d.getSlide(slide)?.opaque?.transition).toBe(TRANSITION);
    expect(changes).toEqual([{ slideId: slide, keys: ['opaque'] }]);
    d.undo();
    expect(d.getSlide(slide)?.opaque?.timing).toBe(TIMING('3', '4'));
    expect(d.getElement(box)).toBeDefined();
    d.redo();
    expect(d.getSlide(slide)?.opaque?.timing).toBe(TIMING('3'));
  });

  it('deleting several shapes prunes once, and the last effect takes the timing (not the transition) with it', () => {
    const { d, slide, chart, box } = deckWith();
    d.deleteElements([chart.id, box]);
    expect(d.getSlide(slide)?.opaque).toEqual({ transition: TRANSITION });
    d.undo();
    expect(d.getSlide(slide)?.opaque?.timing).toBe(TIMING('3', '4'));
  });

  it('a slide left with nothing opaque loses the record', () => {
    const { d, slide, chart, box } = deckWith();
    d.setSlideHidden(slide, true);
    const data = d.getData();
    delete (data.slides[0].opaque as { transition?: string }).transition;
    d.setData(data);
    d.deleteElements([chart.id, box]);
    expect(d.getSlide(slide)).not.toHaveProperty('opaque');
    expect(d.getSlide(slide)?.hidden).toBe(true);
  });

  it('leaves the timing alone for a shape with no PPTX id, one that is not named, a slide without timing, and an id another shape still holds', () => {
    const { d, slide, other } = deckWith();
    d.deleteElement(other);
    expect(d.getSlide(slide)?.opaque?.timing).toBe(TIMING('3', '4'));
    const dup = deckWith();
    dup.d.updateElement('other', { pptxId: '4' });
    dup.d.deleteElement(dup.box);
    expect(dup.d.getSlide(dup.slide)?.opaque?.timing).toBe(TIMING('3', '4'));
    const none = deckWith(null);
    none.d.deleteElement(none.box);
    expect(none.d.getSlide(none.slide)?.opaque).toEqual({ transition: TRANSITION });
    const plain = new DeckImpl();
    const el = plain.addElement(plain.getActiveSlideId(), { type: 'text', x: 0, y: 0, w: 1, h: 1 });
    plain.deleteElement(el.id);
    expect(plain.getSlide(plain.getActiveSlideId())).not.toHaveProperty('opaque');
  });

  it('keeps a slide whose timing is untouched by the deleted shape as the same record', () => {
    const { d, slide } = deckWith(TIMING('3'));
    const before = d.getSlide(slide);
    d.deleteElement('box');
    expect(d.getSlide(slide)).toBe(before);
  });
});

describe('DeckData', () => {
  it('has the opaque fields as optional additions', () => {
    const data: DeckData = new DeckImpl().getData();
    expect(data.opaqueParts).toBeUndefined();
    expect(data.slides[0].opaque).toBeUndefined();
  });
});
