import * as Y from 'yjs';
import { InMemoryProvider, __resetInMemoryRooms, type CollabIdentity } from '@b32nio/spindle-shared';
import { DeckImpl } from '../deck';
import { COLLAB_SCHEMA_VERSION, CollabSchemaVersionError, assertSchemaVersion, hydrateYDocFromData, serializeYDocToData } from './y-schema';
import { getCollabBlockers } from './blockers';

const identity = (name: string): CollabIdentity => ({ userId: name, displayName: name, color: '#123456' });

function build() {
  const d = new DeckImpl('room', 'Deck');
  const s = d.getActiveSlideId();
  const [a, b, c, e] = [0, 1, 2, 3].map((i) => d.addElement(s, { type: 'shape', x: i * 50, y: 0, w: 40, h: 40 }).id);
  return { d, s, a, b, c, e };
}

describe('nested groups in the deck', () => {
  it('groups groups: the outer group takes the whole inner group', () => {
    const { d, a, b, c } = build();
    const inner = d.groupElements([a, b])!;
    const outer = d.groupElements([a, c])!; // a stands for its whole root group (inner)
    expect(outer).not.toBe(inner);
    expect(d.getElement(a)).toMatchObject({ groupId: inner, parentGroupId: outer });
    expect(d.getElement(b)).toMatchObject({ groupId: inner, parentGroupId: outer });
    expect(d.getElement(c)).toMatchObject({ groupId: outer });
    expect(d.getElement(c)!.parentGroupId).toBeUndefined();
    expect(d.getGroupChain(a)).toEqual([inner, outer]);
    expect(d.getGroupChain(c)).toEqual([outer]);
    expect(d.getRootGroupId(b)).toBe(outer);
    expect(d.getRootGroupId('missing')).toBeUndefined();
    expect(d.getGroupMembers(outer).sort()).toEqual([a, b, c].sort());
    expect(d.getGroupMembers(inner).sort()).toEqual([a, b].sort());
  });

  it('nests three deep and refuses to group a single unit', () => {
    const { d, a, b, c, e } = build();
    const g1 = d.groupElements([a, b])!;
    const g2 = d.groupElements([b, c])!;
    const g3 = d.groupElements([a, e])!;
    expect(d.getGroupChain(b)).toEqual([g1, g2, g3]);
    expect(d.getGroupMembers(g3)).toHaveLength(4);
    // Everything in one root group is one unit.
    expect(d.groupElements([a, b, c, e])).toBeUndefined();
    expect(d.groupElements([a])).toBeUndefined();
    expect(d.groupElements(['missing', a])).toBeUndefined();
  });

  it('ungroup dissolves only the outermost group; inner groups survive', () => {
    const { d, a, b, c } = build();
    const inner = d.groupElements([a, b])!;
    const outer = d.groupElements([a, c])!;
    d.ungroupElements([b]);
    expect(d.getElement(c)!.groupId).toBeUndefined();
    expect(d.getElement(a)).toMatchObject({ groupId: inner });
    expect(d.getElement(a)!.parentGroupId).toBeUndefined();
    expect(d.getGroupMembers(outer)).toEqual([]);
    expect(d.getRootGroupId(a)).toBe(inner);
    d.ungroupElements([a]);
    expect(d.getElement(a)!.groupId).toBeUndefined();
    d.ungroupElements([a]); // nothing left to ungroup: no-op
    d.ungroupElements(['missing']);
  });

  it('ungroup leaves other groups and their nesting alone', () => {
    const { d, a, b, c, e } = build();
    const g1 = d.groupElements([a, b])!;
    const g2 = d.groupElements([c, e])!;
    const top = d.groupElements([a, c])!;
    d.ungroupElements([a]);
    expect(d.getElement(a)!.groupId).toBe(g1);
    expect(d.getElement(c)!.groupId).toBe(g2);
    expect(d.getElement(a)!.parentGroupId).toBeUndefined();
    expect(d.getGroupMembers(top)).toEqual([]);
  });

  it('undo restores the nesting in one step', () => {
    const { d, a, b, c } = build();
    const inner = d.groupElements([a, b])!;
    d.groupElements([a, c]);
    d.undo();
    expect(d.getGroupChain(a)).toEqual([inner]);
    expect(d.getElement(c)!.groupId).toBeUndefined();
  });

  it('duplicating a slide keeps the nesting among fresh group ids', () => {
    const { d, s, a, b, c } = build();
    const inner = d.groupElements([a, b])!;
    const outer = d.groupElements([a, c])!;
    const copy = d.duplicateSlide(s)!;
    const els = d.getElementsForSlide(copy.id);
    const grouped = els.filter((e) => e.groupId);
    expect(grouped).toHaveLength(3);
    const copyOuter = grouped.find((e) => !e.parentGroupId)!.groupId!;
    const copyInner = grouped.find((e) => e.parentGroupId)!.groupId!;
    expect([copyInner, copyOuter]).not.toContain(inner);
    expect([copyInner, copyOuter]).not.toContain(outer);
    expect(grouped.filter((e) => e.parentGroupId).every((e) => e.parentGroupId === copyOuter)).toBe(true);
    expect(d.getGroupMembers(copyOuter)).toHaveLength(3);
    expect(d.getGroupMembers(outer).sort()).toEqual([a, b, c].sort());
  });

  it('duplicateElements keeps nesting; a parent outside the copied set is dropped; single duplicate is loose', () => {
    const { d, a, b, c } = build();
    const inner = d.groupElements([a, b])!;
    const outer = d.groupElements([a, c])!;
    const all = d.duplicateElements([a, b, c]);
    const copyOuter = all.find((e) => !e.parentGroupId)!.groupId!;
    expect(copyOuter).not.toBe(outer);
    expect(all.filter((e) => e.parentGroupId).map((e) => e.parentGroupId)).toEqual([copyOuter, copyOuter]);
    const partial = d.duplicateElements([a, b]);
    expect(partial.map((e) => e.groupId)).toEqual([partial[0].groupId, partial[0].groupId]);
    expect(partial[0].groupId).not.toBe(inner);
    expect(partial.every((e) => e.parentGroupId === undefined)).toBe(true);
    const one = d.duplicateElement(a)!;
    expect(one.groupId).toBeUndefined();
    expect(one.parentGroupId).toBeUndefined();
  });
});

describe('nested groups in the collaboration layout (schema v6)', () => {
  it('has no collab blockers and survives hydrate -> serialize', () => {
    const { d, s, a, b, c } = build();
    const inner = d.groupElements([a, b])!;
    const outer = d.groupElements([a, c])!;
    expect(getCollabBlockers(d.getData())).toEqual([]);
    const ydoc = new Y.Doc();
    hydrateYDocFromData(ydoc, d.getData());
    const back = serializeYDocToData(ydoc, { elementIds: [] }, s);
    const els = back.slides[0].elements;
    expect(els.find((e) => e.id === a)).toMatchObject({ groupId: inner, parentGroupId: outer });
    expect(els.find((e) => e.id === c)).toMatchObject({ groupId: outer });
    expect(els.find((e) => e.id === c)!.parentGroupId).toBeUndefined();
  });

  it('propagates grouping and ungrouping between peers', async () => {
    __resetInMemoryRooms();
    const { d: a, s, a: ea, b: eb, c: ec } = build();
    await a.attachCollab(new InMemoryProvider(), identity('A'), { roomId: 'room' });
    const b = new DeckImpl('joiner');
    await b.attachCollab(new InMemoryProvider(), identity('B'), { roomId: 'room' });
    const inner = b.groupElements([ea, eb])!;
    const outer = b.groupElements([ea, ec])!;
    expect(a.getGroupChain(ea)).toEqual([inner, outer]);
    expect(a.getElementsForSlide(s).find((e) => e.id === ec)!.groupId).toBe(outer);
    a.ungroupElements([ea]);
    expect(b.getGroupChain(ea)).toEqual([inner]);
    expect(b.getElement(ea)!.parentGroupId).toBeUndefined();
    expect(b.getElement(ec)!.groupId).toBeUndefined();
  });

  it('a v5 client refuses a v6 room (refuse-newer-room)', () => {
    expect(COLLAB_SCHEMA_VERSION).toBeGreaterThanOrEqual(6);
    const ydoc = new Y.Doc();
    hydrateYDocFromData(ydoc, new DeckImpl('d').getData());
    expect(ydoc.getMap('meta').get('schemaVersion')).toBe(COLLAB_SCHEMA_VERSION);
    expect(() => assertSchemaVersion(ydoc, 5)).toThrow(CollabSchemaVersionError);
    expect(() => assertSchemaVersion(ydoc, 5)).toThrow(/up to version 5/);
    expect(() => assertSchemaVersion(ydoc, COLLAB_SCHEMA_VERSION)).not.toThrow();
  });
});
