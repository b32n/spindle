import * as Y from 'yjs';
import { InMemoryProvider, __resetInMemoryRooms, type CollabIdentity } from '@b32nio/spindle-shared';
import { DeckImpl } from '../deck';
import type { BodyStyle, TextElement } from '../scene/types';
import { hydrateYDocFromData, serializeYDocToData } from './y-schema';
import { getCollabBlockers } from './blockers';

const identity = (name: string): CollabIdentity => ({ userId: name, displayName: name, color: '#123456' });
const body: BodyStyle = {
  vAlign: 'middle',
  insets: { l: 1, t: 2, r: 3, b: 4 },
  anchorCtr: true,
  autofit: { mode: 'shrink', fontScale: 0.625, lnSpcReduction: 0.2 },
  vert: 'eaVert',
  columns: { count: 2, gap: 12 },
};

function build(): DeckImpl {
  const d = new DeckImpl('room', 'Deck');
  d.addElement(d.getActiveSlideId(), { type: 'text', x: 0, y: 0, w: 200, h: 80, bodyStyle: body });
  return d;
}

describe('text body attributes in the collaboration layout', () => {
  it('has no collab blockers and survives hydrate -> serialize', () => {
    const d = build();
    expect(getCollabBlockers(d.getData())).toEqual([]);
    const ydoc = new Y.Doc();
    hydrateYDocFromData(ydoc, d.getData());
    const back = serializeYDocToData(ydoc, { elementIds: [] }, d.getActiveSlideId());
    expect((back.slides[0].elements.find((e) => e.type === 'text') as TextElement).bodyStyle).toEqual(body);
  });

  it('propagates between peers, including a later autofit change', async () => {
    __resetInMemoryRooms();
    const a = build();
    await a.attachCollab(new InMemoryProvider(), identity('A'), { roomId: 'room' });
    const b = new DeckImpl('joiner');
    await b.attachCollab(new InMemoryProvider(), identity('B'), { roomId: 'room' });
    const sid = b.getActiveSlideId();
    const t = b.getElementsForSlide(sid).find((e) => e.type === 'text') as TextElement;
    expect(t.bodyStyle).toEqual(body);
    b.updateElement(t.id, { bodyStyle: { ...body, autofit: { mode: 'resize' } } });
    expect((a.getElement(t.id) as TextElement).bodyStyle?.autofit).toEqual({ mode: 'resize' });
  });
});
