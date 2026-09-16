import * as Y from 'yjs';
import { InMemoryProvider, __resetInMemoryRooms, type CollabIdentity } from '@b32nio/spindle-shared';
import { DeckImpl } from '../deck';
import { getBuiltinTheme } from '../theme/builtin';
import { richTextFromPlainText, richTextToPlainText, type RichTextDoc } from '../text/model';
import { hydrateYDocFromData, serializeYDocToData } from './y-schema';
import type { CollabHandle } from './binding';
import type { DeckData } from '../types';

const identity = (name: string): CollabIdentity => ({ userId: name, displayName: name, color: '#123456' });

function stripSelection(data: DeckData): Omit<DeckData, 'selection'> {
  const { selection: _selection, ...rest } = data;
  return rest;
}

function seedDeck(): DeckData {
  return {
    id: 'room',
    title: 'Deck',
    slideSize: { w: 1280, h: 720 },
    theme: getBuiltinTheme('Clean'),
    slides: [
      {
        id: 's1',
        index: 'V',
        layoutRef: 'blank',
        elements: [
          { id: 'e1', containerId: 's1', index: 'a', type: 'shape', shape: 'rect', x: 0, y: 0, w: 100, h: 100, rotation: 0, fill: { kind: 'solid', color: { kind: 'theme', slot: 'accent1' } } },
          { id: 'e2', containerId: 's1', index: 'b', type: 'text', x: 0, y: 200, w: 300, h: 80, rotation: 0, richText: richTextFromPlainText('hello') },
        ],
      },
    ],
  };
}

async function attachedPair(seed?: DeckData): Promise<{ a: DeckImpl; b: DeckImpl; ha: CollabHandle; hb: CollabHandle }> {
  __resetInMemoryRooms();
  const a = new DeckImpl('room', 'Deck');
  if (seed) a.setData(seed);
  const ha = await a.attachCollab(new InMemoryProvider(), identity('A'), { roomId: 'room' });
  const b = new DeckImpl('joiner');
  const hb = await b.attachCollab(new InMemoryProvider(), identity('B'), { roomId: 'room' });
  return { a, b, ha, hb };
}

// A CollabProvider that delivers everything on a later microtask, the way a
// real socket does — the reply to a sync request never lands synchronously.
// It honours the contract: connect() replays the room's doc log to the new
// peer before resolving. Used to prove seeding is decided from replayed state,
// not from a synchronous handshake that only works in-process.
const asyncRooms = new Map<string, { peers: Set<AsyncRelayProvider>; log: Uint8Array[] }>();
class AsyncRelayProvider {
  private roomId: string | null = null;
  private subs = new Map<string, Set<(p: Uint8Array, from?: string) => void>>();
  async connect(roomId: string): Promise<void> {
    this.roomId = roomId;
    let room = asyncRooms.get(roomId);
    if (!room) {
      room = { peers: new Set(), log: [] };
      asyncRooms.set(roomId, room);
    }
    room.peers.add(this);
    const handlers = this.subs.get('doc');
    for (const payload of [...room.log]) {
      await Promise.resolve(); // defer, so replay is genuinely asynchronous
      if (handlers) for (const h of handlers) h(payload, 'replay');
    }
  }
  disconnect(): void {
    asyncRooms.get(this.roomId ?? '')?.peers.delete(this);
    this.subs.clear();
    this.roomId = null;
  }
  send(channel: string, payload: Uint8Array): void {
    const room = asyncRooms.get(this.roomId ?? '');
    if (!room) return;
    if (channel === 'doc') room.log.push(payload);
    for (const peer of room.peers) {
      if (peer === this) continue;
      const hs = peer.subs.get(channel);
      if (hs) for (const h of hs) void Promise.resolve().then(() => h(payload, 'peer'));
    }
  }
  onMessage(channel: string, handler: (p: Uint8Array, from?: string) => void): () => void {
    let s = this.subs.get(channel);
    if (!s) {
      s = new Set();
      this.subs.set(channel, s);
    }
    s.add(handler);
    return () => s!.delete(handler);
  }
}

afterEach(() => {
  __resetInMemoryRooms();
  asyncRooms.clear();
});

describe('y-schema round-trip', () => {
  it('is idempotent through hydrate → serialize', () => {
    const a = new DeckImpl();
    a.setData(seedDeck());
    const d1 = new Y.Doc();
    hydrateYDocFromData(d1, a.getData());
    const round = serializeYDocToData(d1, { elementIds: [] }, a.getActiveSlideId());

    const b = new DeckImpl();
    b.setData(round);
    const d2 = new Y.Doc();
    hydrateYDocFromData(d2, b.getData());
    const round2 = serializeYDocToData(d2, { elementIds: [] }, b.getActiveSlideId());

    expect(stripSelection(round2)).toEqual(stripSelection(round));
    // Text body survived the fragment round-trip.
    const textEl = round.slides[0].elements.find((e) => e.type === 'text')!;
    expect(richTextToPlainText((textEl as unknown as { richText: RichTextDoc }).richText)).toBe('hello');
  });
});

describe('two-peer convergence over InMemoryProvider', () => {
  it('propagates a joiner into an existing room without duplicating', async () => {
    const { a, b } = await attachedPair(seedDeck());
    expect(b.slideCount()).toBe(1);
    expect(b.getElementsForSlide(b.getActiveSlideId())).toHaveLength(2);
    expect(stripSelection(a.getData())).toEqual(stripSelection(b.getData()));
  });

  it('a joiner that also holds the same seed data does not duplicate', async () => {
    __resetInMemoryRooms();
    const a = new DeckImpl('room', 'Deck');
    a.setData(seedDeck());
    const ha = await a.attachCollab(new InMemoryProvider(), identity('A'), { roomId: 'room' });
    // B loads the document from its own backend before joining — the exact
    // situation that used to stack a second copy of every slide/element.
    const b = new DeckImpl('room', 'Deck');
    b.setData(seedDeck());
    const hb = await b.attachCollab(new InMemoryProvider(), identity('B'), { roomId: 'room' });

    expect(b.slideCount()).toBe(1);
    expect(b.getElementsForSlide(b.getActiveSlideId())).toHaveLength(2);
    expect(a.slideCount()).toBe(1);
    expect(stripSelection(a.getData())).toEqual(stripSelection(b.getData()));
    ha.detach();
    hb.detach();
  });

  it('a seeded joiner does not duplicate over an async (real-socket-like) transport', async () => {
    asyncRooms.clear();
    const a = new DeckImpl('room', 'Deck');
    a.setData(seedDeck());
    const ha = await a.attachCollab(new AsyncRelayProvider() as never, identity('A'), { roomId: 'room' });
    const b = new DeckImpl('room', 'Deck');
    b.setData(seedDeck());
    const hb = await b.attachCollab(new AsyncRelayProvider() as never, identity('B'), { roomId: 'room' });

    // Over an async transport the old code hydrated before the sync reply
    // arrived, producing two of everything. connect()'s replay-before-resolve
    // makes the seed decision correct here too.
    expect(b.slideCount()).toBe(1);
    expect(b.getElementsForSlide(b.getActiveSlideId())).toHaveLength(2);
    ha.detach();
    hb.detach();
  });

  it('converges on concurrent moves and reorders of different elements', async () => {
    const { a, b } = await attachedPair(seedDeck());
    const slideId = a.getActiveSlideId();
    const [e1, e2] = a.getElementIdsForSlide(slideId);

    a.moveElements([e1], 40, 0);
    b.moveElements([e2], 0, 25);
    a.bringToFront([e1]);

    expect(stripSelection(a.getData())).toEqual(stripSelection(b.getData()));
    expect(a.getElement(e1)!.x).toBe(40);
    expect(a.getElement(e2)!.y).toBe(225);
  });

  it('converges on rich-text edits in different elements', async () => {
    const { a, b } = await attachedPair(seedDeck());
    const slideId = a.getActiveSlideId();
    const textId = a.getElementsForSlide(slideId).find((e) => e.type === 'text')!.id;

    a.setElementRichText(textId, richTextFromPlainText('hello world'));
    expect(richTextToPlainText((b.getElement(textId) as unknown as { richText: RichTextDoc }).richText)).toBe('hello world');
    expect(stripSelection(a.getData())).toEqual(stripSelection(b.getData()));
  });

  it('converges on a new element added by a peer', async () => {
    const { a, b } = await attachedPair(seedDeck());
    const slideId = a.getActiveSlideId();
    const created = b.addElement(slideId, { type: 'shape', shape: 'ellipse', x: 500, y: 100 });
    expect(a.getElement(created.id)).toBeDefined();
    expect(stripSelection(a.getData())).toEqual(stripSelection(b.getData()));
  });

  it('syncs comment threads between peers', async () => {
    const { a, b } = await attachedPair(seedDeck());
    const t = a.getComments().addThread({ slideId: 's1', elementId: 'e1' }, 'looks good', { id: 'A', name: 'Alice' });
    const remote = b.getComments().getThread(t.id);
    expect(remote?.comments[0].body).toBe('looks good');
    a.getComments().resolveThread(t.id, { id: 'A', name: 'Alice' });
    expect(b.getComments().getThread(t.id)?.status).toBe('resolved');
  });
});

describe('collab undo', () => {
  it("undo reverts only the local peer's edit, not a remote one", async () => {
    const { a, b } = await attachedPair(seedDeck());
    const slideId = a.getActiveSlideId();
    const ea = a.addElement(slideId, { type: 'shape', x: 10, y: 10 });
    const eb = b.addElement(slideId, { type: 'shape', x: 20, y: 20 });
    expect(a.getElement(eb.id)).toBeDefined(); // remote add visible locally

    a.undo();
    expect(a.getElement(ea.id)).toBeUndefined(); // own add undone
    expect(a.getElement(eb.id)).toBeDefined(); // remote add survives
    expect(stripSelection(a.getData())).toEqual(stripSelection(b.getData()));
  });

  it('tracks a rich-text fragment created after the UndoManager exists', async () => {
    const { a, b, ha } = await attachedPair(seedDeck());
    const slideId = a.getActiveSlideId();
    const t = a.addElement(slideId, { type: 'text', x: 0, y: 400, richText: richTextFromPlainText('one') });
    ha.undoManager.stopCapturing(); // separate the text edit into its own undo step
    a.setElementRichText(t.id, richTextFromPlainText('one two'));
    expect(richTextToPlainText((b.getElement(t.id) as unknown as { richText: RichTextDoc }).richText)).toBe('one two');

    a.undo(); // should revert just the text edit
    expect(richTextToPlainText((a.getElement(t.id) as unknown as { richText: RichTextDoc }).richText)).toBe('one');
    expect(a.getElement(t.id)).toBeDefined();
  });
});
