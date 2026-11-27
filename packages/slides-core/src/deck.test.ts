import { DeckImpl } from './deck';
import { normalizeDeckData } from './serialization';
import type { DeckData, DeckEventType, ElementChangePayload } from './types';
import type { TableElement } from './scene/types';
import { richTextFromPlainText, docHasMark } from './text/model';

const solidFill = { kind: 'solid', color: { kind: 'theme', slot: 'accent1' } } as const;

function capture(deck: DeckImpl, event: DeckEventType): unknown[] {
  const seen: unknown[] = [];
  deck.on(event, (e) => seen.push(e.payload));
  return seen;
}

describe('DeckImpl construction', () => {
  it('starts with one slide and a valid active slide', () => {
    const deck = new DeckImpl();
    expect(deck.slideCount()).toBe(1);
    expect(deck.getSlideIds()).toContain(deck.getActiveSlideId());
  });
});

describe('slides', () => {
  it('adds slides in order and keeps them sorted by fractional index', () => {
    const deck = new DeckImpl();
    const [first] = deck.getSlideIds();
    const b = deck.addSlide({ afterSlideId: first });
    const c = deck.addSlide({ afterSlideId: first }); // between first and b
    expect(deck.getSlideIds()).toEqual([first, c.id, b.id]);
  });

  it('refuses to delete the last slide', () => {
    const deck = new DeckImpl();
    expect(() => deck.deleteSlide(deck.getActiveSlideId())).toThrow();
  });

  it('reassigns active slide when the active one is deleted', () => {
    const deck = new DeckImpl();
    const first = deck.getActiveSlideId();
    const b = deck.addSlide();
    deck.setActiveSlide(b.id);
    deck.deleteSlide(b.id);
    expect(deck.getActiveSlideId()).toBe(first);
  });

  it('moves a slide to a new position', () => {
    const deck = new DeckImpl();
    const first = deck.getActiveSlideId();
    const b = deck.addSlide();
    const c = deck.addSlide();
    // Order: first, b, c → move c to front
    deck.moveSlide(c.id, {});
    expect(deck.getSlideIds()).toEqual([c.id, first, b.id]);
  });

  it('duplicates a slide together with its elements', () => {
    const deck = new DeckImpl();
    const slide = deck.getActiveSlideId();
    deck.addElement(slide, { type: 'shape', shape: 'ellipse' });
    const copy = deck.duplicateSlide(slide);
    expect(copy).toBeDefined();
    const copyEls = deck.getElementsForSlide(copy!.id);
    expect(copyEls).toHaveLength(1);
    expect(copyEls[0].containerId).toBe(copy!.id);
    // Duplicated elements are independent records with new ids.
    expect(copyEls[0].id).not.toBe(deck.getElementsForSlide(slide)[0].id);
  });

  it('emits elementAdd for every copied element (so collab mirrors them)', () => {
    const deck = new DeckImpl();
    const slide = deck.getActiveSlideId();
    deck.addElement(slide, { type: 'shape', shape: 'ellipse' });
    deck.addElement(slide, { type: 'shape', shape: 'rect' });
    const adds = capture(deck, 'elementAdd') as { slideId: string; elementId: string }[];
    const copy = deck.duplicateSlide(slide)!;
    const copyEls = deck.getElementsForSlide(copy.id);
    // One elementAdd per copied element, all targeting the new slide.
    expect(adds).toHaveLength(2);
    expect(adds.every((a) => a.slideId === copy.id)).toBe(true);
    expect(new Set(adds.map((a) => a.elementId))).toEqual(new Set(copyEls.map((e) => e.id)));
  });

  it('remaps internal references (group + connector binds) when duplicating a slide', () => {
    const deck = new DeckImpl();
    const slide = deck.getActiveSlideId();
    const a = deck.addElement(slide, { type: 'shape', x: 0, y: 0, w: 100, h: 100 });
    const b = deck.addElement(slide, { type: 'shape', x: 400, y: 0, w: 100, h: 100 });
    const groupId = deck.groupElements([a.id, b.id])!;
    const conn = deck.addElement(slide, {
      type: 'line', x: 100, y: 50, w: 300, h: 0,
      startBind: { elementId: a.id, anchor: 'e' },
    });

    const copy = deck.duplicateSlide(slide)!;
    const copyEls = deck.getElementsForSlide(copy.id);
    expect(copyEls).toHaveLength(3);
    const srcIds = new Set([a.id, b.id, conn.id]);
    for (const el of copyEls) expect(srcIds.has(el.id)).toBe(false);

    // The copied group is a fresh group made of only the copies.
    const copyGroupId = copyEls.find((e) => e.type !== 'line')!.groupId!;
    expect(copyGroupId).not.toBe(groupId);
    const groupMembers = deck.getGroupMembers(copyGroupId);
    expect(groupMembers.sort()).toEqual(copyEls.filter((e) => e.type !== 'line').map((e) => e.id).sort());
    // The original group is untouched (only its two source elements).
    expect(deck.getGroupMembers(groupId).sort()).toEqual([a.id, b.id].sort());

    // The copied connector binds to the COPIED shape, not the original.
    const copyConn = copyEls.find((e) => e.type === 'line')!;
    if (copyConn.type !== 'line') throw new Error('not a line');
    const copyA = copyEls.find((e) => e.type !== 'line' && e.x === 0)!;
    expect(copyConn.startBind?.elementId).toBe(copyA.id);
    expect([a.id, b.id]).not.toContain(copyConn.startBind?.elementId);
  });
});

describe('elements', () => {
  it('adds elements at the top of the z-order', () => {
    const deck = new DeckImpl();
    const slide = deck.getActiveSlideId();
    const a = deck.addElement(slide, { type: 'shape' });
    const b = deck.addElement(slide, { type: 'text' });
    expect(deck.getElementIdsForSlide(slide)).toEqual([a.id, b.id]);
    expect(a.index < b.index).toBe(true);
  });

  it('replaces the record on update (immutable) and emits keys', () => {
    const deck = new DeckImpl();
    const slide = deck.getActiveSlideId();
    const el = deck.addElement(slide, { type: 'shape' });
    const changes = capture(deck, 'elementChange') as ElementChangePayload[];
    deck.updateElement(el.id, { x: 500 });
    const updated = deck.getElement(el.id)!;
    expect(updated).not.toBe(el); // new object reference
    expect(updated.x).toBe(500);
    expect(changes[0]).toMatchObject({ elementId: el.id, slideId: slide, keys: ['x'] });
  });

  it('batches multi-element updates', () => {
    const deck = new DeckImpl();
    const slide = deck.getActiveSlideId();
    const a = deck.addElement(slide, { type: 'shape' });
    const b = deck.addElement(slide, { type: 'shape' });
    const changes = capture(deck, 'elementChange');
    deck.updateElements([
      { id: a.id, patch: { x: 10 } },
      { id: b.id, patch: { x: 20 } },
    ]);
    expect(changes).toHaveLength(2);
    expect(deck.getElement(a.id)!.x).toBe(10);
    expect(deck.getElement(b.id)!.x).toBe(20);
  });

  it('deletes elements and duplicates with an offset', () => {
    const deck = new DeckImpl();
    const slide = deck.getActiveSlideId();
    const el = deck.addElement(slide, { type: 'shape', x: 100, y: 100 });
    const copy = deck.duplicateElement(el.id)!;
    expect(copy.x).toBe(116);
    expect(copy.y).toBe(116);
    deck.deleteElement(el.id);
    expect(deck.getElement(el.id)).toBeUndefined();
    expect(deck.getElementsForSlide(slide)).toHaveLength(1);
  });

  it('shifts a duplicated line\'s explicit endpoints alongside its box', () => {
    const deck = new DeckImpl();
    const slide = deck.getActiveSlideId();
    const line = deck.addElement(slide, {
      type: 'line', x: 100, y: 100, w: 200, h: 100,
      startPoint: { x: 100, y: 100 }, endPoint: { x: 300, y: 200 },
    });
    const copy = deck.duplicateElement(line.id)!;
    if (copy.type !== 'line') throw new Error('not a line');
    expect(copy.x).toBe(116);
    expect(copy.y).toBe(116);
    // Endpoints move with the box so the diagonal stays in sync.
    expect(copy.startPoint).toEqual({ x: 116, y: 116 });
    expect(copy.endPoint).toEqual({ x: 316, y: 216 });
  });

  it('moves elements by a delta as one undo entry', () => {
    const deck = new DeckImpl();
    const slide = deck.getActiveSlideId();
    const a = deck.addElement(slide, { type: 'shape', x: 0, y: 0 });
    const b = deck.addElement(slide, { type: 'shape', x: 50, y: 50 });
    deck.moveElements([a.id, b.id], 10, -5);
    expect(deck.getElement(a.id)!).toMatchObject({ x: 10, y: -5 });
    expect(deck.getElement(b.id)!).toMatchObject({ x: 60, y: 45 });
  });
});

describe('events', () => {
  it('fires slideAdd / elementAdd / selectionChange', () => {
    const deck = new DeckImpl();
    const slide = deck.getActiveSlideId();
    const slideAdds = capture(deck, 'slideAdd');
    const elementAdds = capture(deck, 'elementAdd');
    const selChanges = capture(deck, 'selectionChange');
    deck.addSlide();
    deck.addElement(slide, { type: 'text' });
    deck.setSelection({ slideId: slide, elementIds: ['x'] });
    expect(slideAdds).toHaveLength(1);
    expect(elementAdds).toHaveLength(1);
    expect(selChanges).toHaveLength(1);
  });
});

describe('undo / redo', () => {
  it('undoes and redoes an element addition', () => {
    const deck = new DeckImpl();
    const slide = deck.getActiveSlideId();
    const el = deck.addElement(slide, { type: 'shape' });
    expect(deck.canUndo()).toBe(true);
    deck.undo();
    expect(deck.getElement(el.id)).toBeUndefined();
    expect(deck.canRedo()).toBe(true);
    deck.redo();
    expect(deck.getElement(el.id)).toBeDefined();
  });

  it('undoes a transform back to the prior frame', () => {
    const deck = new DeckImpl();
    const slide = deck.getActiveSlideId();
    const el = deck.addElement(slide, { type: 'shape', x: 100 });
    deck.updateElement(el.id, { x: 400 });
    deck.undo();
    expect(deck.getElement(el.id)!.x).toBe(100);
  });
});

describe('layout materialization', () => {
  it('materializes placeholder elements when adding a slide with a layout', () => {
    const deck = new DeckImpl();
    const slide = deck.addSlide({ layoutId: 'titleContent' });
    const els = deck.getElementsForSlide(slide.id);
    expect(els).toHaveLength(2); // title + body
    expect(els.every((e) => e.type === 'text' && !!e.placeholder)).toBe(true);
  });

  it('exposes the layout prompt for a placeholder element', () => {
    const deck = new DeckImpl();
    const slide = deck.addSlide({ layoutId: 'titleContent' });
    const title = deck.getElementsForSlide(slide.id).find((e) => e.placeholder?.type === 'title')!;
    expect(deck.getPlaceholderPrompt(slide.id, title.placeholder)).toBe('Click to add title');
    // Element carries the layout's default text style.
    expect((title as { bodyStyle?: { fontSize?: number } }).bodyStyle?.fontSize).toBe(44);
  });

  it('materializes as one undoable step with the slide', () => {
    const deck = new DeckImpl();
    const before = deck.slideCount();
    const slide = deck.addSlide({ layoutId: 'titleContent' });
    expect(deck.getElementsForSlide(slide.id)).toHaveLength(2);
    deck.undo();
    expect(deck.slideCount()).toBe(before);
    expect(deck.getElementsForSlide(slide.id)).toHaveLength(0);
  });
});

describe('getData / setData', () => {
  const handAuthored: DeckData = {
    id: 'deck1',
    title: 'Hand authored',
    slideSize: { w: 1280, h: 720 },
    theme: {
      name: 'Clean',
      colors: {
        dk1: '#000000', lt1: '#FFFFFF', dk2: '#333333', lt2: '#EEEEEE',
        accent1: '#2D7FF9', accent2: '#16B1A6', accent3: '#F5A623',
        accent4: '#E8543F', accent5: '#8C54FF', accent6: '#4CAF50',
        hlink: '#2D7FF9', folHlink: '#8C54FF',
      },
      fonts: { major: 'Inter', minor: 'Inter' },
    },
    // Deliberately omit slide/element ids and indices — normalization fills them.
    slides: [
      {
        elements: [
          { type: 'shape', shape: 'rect', x: 0, y: 0, w: 100, h: 100, rotation: 0, fill: { kind: 'solid', color: { kind: 'theme', slot: 'accent1' } }, containerId: '', id: '', index: '' },
          { type: 'text', x: 0, y: 0, w: 200, h: 50, rotation: 0, richText: richTextFromPlainText('Hi'), containerId: '', id: '', index: '' },
        ],
      } as never,
      { elements: [] } as never,
    ],
  };

  it('normalizes hand-authored data (ids, indices, containerIds)', () => {
    const deck = new DeckImpl();
    deck.setData(handAuthored);
    expect(deck.slideCount()).toBe(2);
    const slideId = deck.getSlideIds()[0];
    const els = deck.getElementsForSlide(slideId);
    expect(els).toHaveLength(2);
    for (const el of els) {
      expect(el.id.length).toBeGreaterThan(0);
      expect(el.index.length).toBeGreaterThan(0);
      expect(el.containerId).toBe(slideId);
    }
  });

  it('round-trips: getData → setData → getData is stable', () => {
    const a = new DeckImpl();
    a.setData(handAuthored);
    const dataA = a.getData();

    const b = new DeckImpl();
    b.setData(dataA);
    const dataB = b.getData();

    expect(dataB).toEqual(dataA);
  });

  it('preserves array order when only some elements have indices', () => {
    // A has a high index, C a lower one, B none — mixing given + generated keys
    // would sort them wrong. Normalization must keep the array order [A,B,C].
    const shape = (id: string, index: string) => ({
      type: 'shape', shape: 'rect', x: 0, y: 0, w: 10, h: 10, rotation: 0,
      fill: solidFill, id, containerId: 's', index,
    });
    const norm = normalizeDeckData({
      id: 'd', title: 't', slideSize: { w: 1280, h: 720 },
      slides: [{ id: 's', index: 'V', elements: [shape('A', 'zzz'), shape('B', ''), shape('C', 'zzy')] }],
    } as unknown as DeckData);
    const els = norm.slides[0].elements!;
    const byIndex = [...els].sort((x, y) => (x.index! < y.index! ? -1 : 1));
    expect(byIndex.map((e) => e.id)).toEqual(['A', 'B', 'C']);
  });

  it('preserves array order when only some slides have indices', () => {
    const slide = (id: string, index: string) => ({ id, index, elements: [] });
    const norm = normalizeDeckData({
      id: 'd', title: 't', slideSize: { w: 1280, h: 720 },
      slides: [slide('A', 'zzz'), slide('B', ''), slide('C', 'zzy')],
    } as unknown as DeckData);
    const bySlideIndex = [...norm.slides].sort((x, y) => (x.index! < y.index! ? -1 : 1));
    expect(bySlideIndex.map((s) => s.id)).toEqual(['A', 'B', 'C']);
  });
});

describe('connectors (bound lines track their shapes)', () => {
  function setup() {
    const deck = new DeckImpl();
    const slide = deck.getActiveSlideId();
    const a = deck.addElement(slide, { type: 'shape', x: 0, y: 0, w: 100, h: 100 });
    const b = deck.addElement(slide, { type: 'shape', x: 400, y: 400, w: 100, h: 100 });
    const conn = deck.addElement(slide, {
      type: 'line', x: 100, y: 100, w: 300, h: 300,
      startBind: { elementId: a.id, anchor: 'e' },
      endBind: { elementId: b.id, anchor: 'w' },
    });
    return { deck, slide, a, b, conn };
  }

  it('recomputes the connector box when a bound shape moves', () => {
    const { deck, a, conn } = setup();
    deck.moveElements([a.id], 50, 20); // A → (50,20,100,100); east anchor (150,70)
    const line = deck.getElement(conn.id)!;
    // B west anchor stays (400,450). Box spans (150,70)→(400,450).
    expect(line.x).toBe(150);
    expect(line.y).toBe(70);
    expect(line.w).toBe(250);
    expect(line.h).toBe(380);
  });

  it('refreshes the stored connector box when a bound shape is flipped (single and batch)', () => {
    const { deck, a, conn } = setup();
    deck.updateElement(a.id, { flipH: true }); // east site mirrors to the box's west edge (0,50)
    let line = deck.getElement(conn.id)!;
    expect(line.x).toBe(0);
    expect(line.w).toBe(400);
    deck.updateElement(a.id, { flipH: false });
    expect(deck.getElement(conn.id)!.x).toBe(100);
    deck.updateElements([{ id: a.id, patch: { flipH: true } }]);
    line = deck.getElement(conn.id)!;
    expect(line.x).toBe(0);
    expect(line.w).toBe(400);
  });

  it('mirrors a flipped plain rectangle: a bound nw site resolves to the ne corner', () => {
    const deck = new DeckImpl();
    const slide = deck.getActiveSlideId();
    const r = deck.addElement(slide, { type: 'shape', x: 0, y: 0, w: 100, h: 100, flipH: true });
    const t = deck.addElement(slide, { type: 'shape', x: 400, y: 400, w: 100, h: 100 });
    const l = deck.addElement(slide, {
      type: 'line', x: 0, y: 0, w: 400, h: 400,
      startBind: { elementId: r.id, anchor: 'nw' },
      endBind: { elementId: t.id, anchor: 'w' },
    });
    deck.updateElement(t.id, { x: 410 });
    const line = deck.getElement(l.id)!;
    expect(line.x).toBe(100); // start at the mirrored (100,0)
    expect(line.y).toBe(0);
  });

  it('tracks a bound shape through resize', () => {
    const { deck, b, conn } = setup();
    deck.updateElement(b.id, { x: 300, w: 200 }); // west anchor now (300,450)
    const line = deck.getElement(conn.id)!;
    expect(line.x).toBe(100); // A east (100,50)
    expect(line.w).toBe(200); // to (300,450)
  });

  it('detaches (clears the binding) when a bound shape is deleted, keeping the connector', () => {
    const { deck, a, conn } = setup();
    deck.deleteElements([a.id]);
    const line = deck.getElement(conn.id)!;
    expect(line).toBeDefined();
    expect(line.type).toBe('line');
    if (line.type !== 'line') return;
    expect(line.startBind).toBeUndefined();
    expect(line.endBind).toBeDefined(); // B still bound
  });

  it('keeps a mid-air connector (bound start, free end) non-degenerate as the source moves', () => {
    const deck = new DeckImpl();
    const slide = deck.getActiveSlideId();
    const a = deck.addElement(slide, { type: 'shape', x: 400, y: 300, w: 100, h: 100 });
    // Free end dropped up-left of the source (the case that used to collapse).
    const conn = deck.addElement(slide, {
      type: 'line', x: 200, y: 150, w: 300, h: 200,
      startBind: { elementId: a.id, anchor: 'w' }, endPoint: { x: 200, y: 150 },
    });
    let line = deck.getElement(conn.id)!;
    if (line.type !== 'line') throw new Error('not a line');
    expect(line.w !== 0 || line.h !== 0).toBe(true); // visible
    // Move the source — the bound end follows, the free end stays pinned.
    deck.moveElements([a.id], 100, 40);
    line = deck.getElement(conn.id)!;
    if (line.type !== 'line') throw new Error('not a line');
    expect(line.endPoint).toEqual({ x: 200, y: 150 }); // free end unchanged
    expect(line.w !== 0 || line.h !== 0).toBe(true); // still visible
  });

  it('setLineEndpoint pins a free tip and freezes the other end explicitly', () => {
    const deck = new DeckImpl();
    const slide = deck.getActiveSlideId();
    const line = deck.addElement(slide, { type: 'line', x: 100, y: 100, w: 200, h: 0 });
    deck.setLineEndpoint(line.id, 'end', { point: { x: 400, y: 250 } });
    const l = deck.getElement(line.id)!;
    if (l.type !== 'line') throw new Error('not a line');
    expect(l.endPoint).toEqual({ x: 400, y: 250 });
    expect(l.startPoint).toEqual({ x: 100, y: 100 }); // frozen from the old box corner
    expect(l.endBind).toBeUndefined();
    // box spans the two points
    expect({ x: l.x, y: l.y, w: l.w, h: l.h }).toEqual({ x: 100, y: 100, w: 300, h: 150 });
  });

  it('setLineEndpoint can bind a tip to a shape anchor (clearing its free point)', () => {
    const deck = new DeckImpl();
    const slide = deck.getActiveSlideId();
    const shape = deck.addElement(slide, { type: 'shape', x: 500, y: 500, w: 100, h: 100 });
    const line = deck.addElement(slide, { type: 'line', x: 100, y: 100, w: 200, h: 0 });
    deck.setLineEndpoint(line.id, 'end', { bind: { elementId: shape.id, anchor: 'w' } });
    const l = deck.getElement(line.id)!;
    if (l.type !== 'line') throw new Error('not a line');
    expect(l.endBind).toEqual({ elementId: shape.id, anchor: 'w' });
    expect(l.endPoint).toBeUndefined();
    // Moving the shape now drags the bound tip (start stays at x=100, so the
    // width grows as the shape's west anchor moves right).
    const beforeW = l.w;
    deck.moveElements([shape.id], 100, 0);
    expect(deck.getElement(line.id)!.w).not.toBe(beforeW);
  });

  it('reconciliation is one undo entry with the move', () => {
    const { deck, a, conn } = setup();
    const before = deck.getElement(conn.id)!.x;
    deck.moveElements([a.id], 50, 20);
    expect(deck.getElement(conn.id)!.x).not.toBe(before);
    deck.undo();
    expect(deck.getElement(conn.id)!.x).toBe(before); // connector restored too
    expect(deck.getElement(a.id)!.x).toBe(0);
  });
});

describe('table multi-cell ops', () => {
  function withTable() {
    const deck = new DeckImpl();
    const slide = deck.getActiveSlideId();
    const t = deck.addElement(slide, { type: 'table', rows: 3, cols: 3 });
    return { deck, id: t.id };
  }
  const fillOf = (deck: DeckImpl, id: string, r: number, c: number) =>
    (deck.getElement(id) as TableElement).cells[r][c].fill;

  it('fills a whole row in one undo, leaving other rows untouched', () => {
    const { deck, id } = withTable();
    const row0: Array<[number, number]> = [[0, 0], [0, 1], [0, 2]];
    deck.updateTableCells(id, row0, { fill: { kind: 'solid', color: { kind: 'rgb', hex: '#112233' } } });
    for (let c = 0; c < 3; c++) expect(fillOf(deck, id, 0, c)).toBeDefined();
    expect(fillOf(deck, id, 1, 0)).toBeUndefined();
    deck.undo();
    expect(fillOf(deck, id, 0, 0)).toBeUndefined(); // single undo reverts the whole row
  });

  it('ignores cells outside the grid', () => {
    const { deck, id } = withTable();
    deck.updateTableCells(id, [[0, 0], [9, 9], [-1, 0]], { fill: { kind: 'solid', color: { kind: 'rgb', hex: '#abcdef' } } });
    expect(fillOf(deck, id, 0, 0)).toBeDefined();
  });

  it('bolds a header row across cells as one undo', () => {
    const { deck, id } = withTable();
    for (let c = 0; c < 3; c++) deck.setTableCellRichText(id, 0, c, richTextFromPlainText(`H${c}`));
    deck.applyTableCellsFormat(id, [[0, 0], [0, 1], [0, 2]], { toggleMark: 'bold' });
    const cells = (deck.getElement(id) as TableElement).cells;
    for (let c = 0; c < 3; c++) expect(docHasMark(cells[0][c].richText, 'bold')).toBe(true);
    expect(docHasMark(cells[1][0].richText, 'bold')).toBe(false);
    deck.undo();
    expect(docHasMark((deck.getElement(id) as TableElement).cells[0][0].richText, 'bold')).toBe(false);
  });
});

describe('table height sync', () => {
  it('sets the frame height to the measured content height (up or down)', () => {
    const deck = new DeckImpl();
    const slide = deck.getActiveSlideId();
    const t = deck.addElement(slide, { type: 'table', rows: 2, cols: 2, h: 100 });
    deck.syncElementHeight(t.id, 260);
    expect(deck.getElement(t.id)!.h).toBe(260);
    deck.syncElementHeight(t.id, 180); // content shrank (row deleted) → follows down
    expect(deck.getElement(t.id)!.h).toBe(180);
  });

  it('records no undo entry (a reflow is not a user edit)', () => {
    const deck = new DeckImpl();
    const slide = deck.getActiveSlideId();
    const t = deck.addElement(slide, { type: 'table', rows: 2, cols: 2, h: 100, x: 10 });
    deck.updateElement(t.id, { x: 50 }); // a real edit (records history)
    deck.syncElementHeight(t.id, 300);
    deck.undo(); // undoes the x move; the height sync left no separate entry
    expect(deck.getElement(t.id)!.x).toBe(10);
  });
});

describe('table batch row/column removal', () => {
  it('removes an inclusive row range in one undo', () => {
    const deck = new DeckImpl();
    const slide = deck.getActiveSlideId();
    const t = deck.addElement(slide, { type: 'table', rows: 4, cols: 2 });
    deck.removeTableRows(t.id, 1, 2); // remove rows 1..2
    expect((deck.getElement(t.id) as TableElement).rows).toBe(2);
    deck.undo();
    expect((deck.getElement(t.id) as TableElement).rows).toBe(4); // single undo restores both
  });
  it('never removes the last column', () => {
    const deck = new DeckImpl();
    const slide = deck.getActiveSlideId();
    const t = deck.addElement(slide, { type: 'table', rows: 2, cols: 3 });
    deck.removeTableColumns(t.id, 0, 2); // would remove all → clamps to 1
    expect((deck.getElement(t.id) as TableElement).cols).toBe(1);
  });
});
