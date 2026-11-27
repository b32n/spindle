import * as Y from 'yjs';
import { InMemoryProvider, __resetInMemoryRooms, type CollabIdentity } from '@b32nio/spindle-shared';
import { DeckImpl } from '../deck';
import type { Fill } from '../scene/types';
import { hydrateYDocFromData, serializeYDocToData } from './y-schema';
import { getCollabBlockers } from './blockers';

const identity = (name: string): CollabIdentity => ({ userId: name, displayName: name, color: '#123456' });
const PNG = 'data:image/png;base64,iVBORw0KGgo=';

const gradient: Fill = {
  kind: 'gradient',
  type: 'linear',
  angle: 30,
  stops: [
    { pos: 0, color: { kind: 'theme', slot: 'accent1', lumMod: 0.5 } },
    { pos: 1, color: { kind: 'rgb', hex: '#ffffff', alpha: 0.2 } },
  ],
};
const picture: Fill = { kind: 'picture', src: PNG, mode: 'tile', tile: { w: 10, h: 12 }, alpha: 0.5 };
const pattern: Fill = { kind: 'pattern', preset: 'dkHorz', fg: { kind: 'rgb', hex: '#000000' }, bg: { kind: 'rgb', hex: '#ffffff' } };

function build(): DeckImpl {
  const d = new DeckImpl('room', 'Deck');
  const s = d.getActiveSlideId();
  d.addElement(s, { type: 'shape', shape: 'rect', x: 0, y: 0, w: 10, h: 10, fill: gradient });
  d.addElement(s, { type: 'text', x: 0, y: 20, w: 10, h: 10, fill: picture });
  const table = d.addElement(s, { type: 'table', rows: 1, cols: 2, x: 0, y: 40, w: 100, h: 20 });
  d.updateTableCells(table.id, [[0, 1]], { fill: pattern });
  d.setSlideBackground(s, gradient);
  return d;
}

describe('rich fills in the collaboration layout', () => {
  it('have no collab blockers', () => {
    expect(getCollabBlockers(build().getData())).toEqual([]);
  });

  it('survive hydrate -> serialize unchanged', () => {
    const d = build();
    const ydoc = new Y.Doc();
    hydrateYDocFromData(ydoc, d.getData());
    const back = serializeYDocToData(ydoc, { elementIds: [] }, d.getActiveSlideId());
    const slide = back.slides[0];
    expect(slide.background).toEqual(gradient);
    const byType = Object.fromEntries(slide.elements.map((e) => [e.type, e]));
    expect((byType.shape as { fill: Fill }).fill).toEqual(gradient);
    expect((byType.text as { fill: Fill }).fill).toEqual(picture);
    expect((byType.table as { cells: Array<Array<{ fill?: Fill }>> }).cells[0][1].fill).toEqual(pattern);
  });

  it('propagate between two peers, including later edits and background changes', async () => {
    __resetInMemoryRooms();
    const a = build();
    await a.attachCollab(new InMemoryProvider(), identity('A'), { roomId: 'room' });
    const b = new DeckImpl('joiner');
    await b.attachCollab(new InMemoryProvider(), identity('B'), { roomId: 'room' });
    const sid = b.getActiveSlideId();
    expect(b.getSlide(sid)?.background).toEqual(gradient);
    const shape = b.getElementsForSlide(sid).find((e) => e.type === 'shape')!;
    expect((shape as { fill: Fill }).fill).toEqual(gradient);

    b.updateElement(shape.id, { fill: picture });
    b.setSlideBackground(sid, pattern);
    const aShape = a.getElementsForSlide(sid).find((e) => e.id === shape.id)!;
    expect((aShape as { fill: Fill }).fill).toEqual(picture);
    expect(a.getSlide(sid)?.background).toEqual(pattern);
  });
});
