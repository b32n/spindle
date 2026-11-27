import * as Y from 'yjs';
import { InMemoryProvider, __resetInMemoryRooms, type CollabIdentity } from '@b32nio/spindle-shared';
import { DeckImpl } from '../deck';
import type { ImageElement } from '../scene/types';
import { hydrateYDocFromData, serializeYDocToData } from './y-schema';
import { getCollabBlockers } from './blockers';

const identity = (name: string): CollabIdentity => ({ userId: name, displayName: name, color: '#123456' });
const PNG = 'data:image/png;base64,iVBORw0KGgo=';
const crop = { l: 10, t: 20, r: 30, b: 5 };

function build(): DeckImpl {
  const d = new DeckImpl('room', 'Deck');
  d.addElement(d.getActiveSlideId(), { type: 'image', src: PNG, naturalW: 4, naturalH: 4, x: 0, y: 0, w: 40, h: 40, fit: 'cover', crop });
  return d;
}

describe('image crop in the collaboration layout', () => {
  it('has no collab blockers and survives hydrate -> serialize', () => {
    const d = build();
    expect(getCollabBlockers(d.getData())).toEqual([]);
    const ydoc = new Y.Doc();
    hydrateYDocFromData(ydoc, d.getData());
    const back = serializeYDocToData(ydoc, { elementIds: [] }, d.getActiveSlideId());
    const img = back.slides[0].elements.find((e) => e.type === 'image') as ImageElement;
    expect(img.crop).toEqual(crop);
    expect(img.fit).toBe('cover');
  });

  it('propagates between peers, including a later edit', async () => {
    __resetInMemoryRooms();
    const a = build();
    await a.attachCollab(new InMemoryProvider(), identity('A'), { roomId: 'room' });
    const b = new DeckImpl('joiner');
    await b.attachCollab(new InMemoryProvider(), identity('B'), { roomId: 'room' });
    const sid = b.getActiveSlideId();
    const img = b.getElementsForSlide(sid).find((e) => e.type === 'image') as ImageElement;
    expect(img.crop).toEqual(crop);
    b.updateElement(img.id, { crop: { l: 1, t: 2, r: 3, b: 4 } });
    expect((a.getElementsForSlide(sid).find((e) => e.id === img.id) as ImageElement).crop).toEqual({ l: 1, t: 2, r: 3, b: 4 });
  });
});
