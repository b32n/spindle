// Tracked changes between collaborators: suggestions by several people, settled by another, converge; offline suggestions
// merge; the room's schema version keeps older clients out; presence gives people their colours.

import * as Y from 'yjs';
import { Awareness } from 'y-protocols/awareness';
import { EditorState, Transaction } from 'prosemirror-state';
import { updateYFragment, yXmlFragmentToProseMirrorRootNode } from 'y-prosemirror';
import type { Node as PmNode } from 'prosemirror-model';
import { docsSchema } from '../prosemirror/schema';
import { listChanges, resolveAll, resolveChange } from '../prosemirror/review';
import { suggestingPlugin } from '../prosemirror/suggesting';
import { COLLAB_SCHEMA_VERSION, CollabSchemaVersionError, assertSchemaVersion, getYDocFields, hydrateYDocFromData, readSchemaVersion, serializeYDocToData } from './y-schema';
import { collabAuthorColors } from './presence';
import { assignAuthorColors } from '../revisions';
import { DEFAULT_PAGE_CONFIG } from '../types';
import type { Block, DocumentData } from '../types';

class Client {
  readonly ydoc = new Y.Doc();
  private ids = 0;
  constructor(readonly name: string) {}
  get fragment(): Y.XmlFragment {
    return getYDocFields(this.ydoc).content;
  }
  pm(): PmNode {
    return yXmlFragmentToProseMirrorRootNode(this.fragment, docsSchema);
  }
  private push(doc: PmNode): void {
    updateYFragment(this.ydoc, this.fragment, doc, { mapping: new Map(), isOMark: new Map() });
  }
  /** What the user does in Suggesting mode: the transaction goes through the plugin, and the result is what the Y document gets. */
  suggest(fn: (tr: Transaction) => void): void {
    const state = EditorState.create({
      doc: this.pm(),
      plugins: [suggestingPlugin({ isActive: () => true, author: () => this.name, now: () => '2024-06-01T10:00:00Z', newId: () => `${this.name}-${++this.ids}` })],
    });
    const tr = state.tr;
    fn(tr);
    this.push(state.apply(tr).doc);
  }
  /** A plain edit, or a review command, run as the editor would and pushed. */
  edit(fn: (state: EditorState) => Transaction | null): void {
    const state = EditorState.create({ doc: this.pm() });
    const tr = fn(state);
    if (tr) this.push(state.apply(tr).doc);
  }
  review(action: 'accept' | 'reject', author?: string): void {
    this.edit((state) => {
      let out: Transaction | null = null;
      expect(resolveAll(action, author)(state, (tr) => (out = tr))).toBe(true);
      return out;
    });
  }
  text(): string[] {
    const out: string[] = [];
    this.pm().forEach((b) => out.push(b.textContent));
    return out;
  }
}

function sync(...clients: Client[]): void {
  for (let round = 0; round < 2; round++) {
    for (const a of clients) {
      for (const b of clients) {
        if (a === b) continue;
        Y.applyUpdate(b.ydoc, Y.encodeStateAsUpdate(a.ydoc, Y.encodeStateVector(b.ydoc)));
      }
    }
  }
}

const para = (text: string): Block => ({ id: 'p', type: 'paragraph', content: [{ type: 'text', text }] }) as Block;
function data(blocks: Block[]): DocumentData {
  return { id: 'doc', title: 'Doc', sections: [{ id: 's1', pageConfig: DEFAULT_PAGE_CONFIG, blocks }], defaultPageConfig: DEFAULT_PAGE_CONFIG, textStylePool: {}, paragraphStylePool: {} };
}
function room(names: string[], blocks: Block[]): Client[] {
  const [first, ...rest] = names.map((n) => new Client(n));
  hydrateYDocFromData(first.ydoc, data(blocks));
  const clients = [first, ...rest];
  sync(...clients);
  return clients;
}
const same = (clients: Client[]) => {
  const [a, ...rest] = clients;
  for (const b of rest) expect(b.pm().toJSON()).toEqual(a.pm().toJSON());
};
const changesOf = (c: Client) => listChanges(c.pm()).map((x) => [x.kind, x.author, x.text]);

describe('suggestions by several people', () => {
  it('two people suggest at once in different places: both are kept, by their authors, and everybody sees the same', () => {
    const [a, b, c] = room(['Ada', 'Bo', 'Cy'], [para('one two three')]);
    a.suggest((tr) => tr.insertText('X', 4));
    b.suggest((tr) => tr.delete(11, 14));
    sync(a, b, c);
    same([a, b, c]);
    expect(changesOf(c)).toEqual([['insertion', 'Ada', 'X'], ['deletion', 'Bo', 'ree']]);
    expect(c.text()).toEqual(['oneX two three']);
  });

  it('two people typing at the same place both keep their text, each marked as theirs', () => {
    const [a, b] = room(['Ada', 'Bo'], [para('abcd')]);
    a.suggest((tr) => tr.insertText('X', 3));
    b.suggest((tr) => tr.insertText('Y', 3));
    sync(a, b);
    same([a, b]);
    const found = changesOf(a);
    expect(found.map((x) => x[1]).sort()).toEqual(['Ada', 'Bo']);
    expect(a.text()[0]).toMatch(/^ab(XY|YX)cd$/);
  });

  it('a third person accepts all: everybody ends with the suggested text and no changes', () => {
    const [a, b, c] = room(['Ada', 'Bo', 'Cy'], [para('one two three')]);
    a.suggest((tr) => tr.insertText('X', 4));
    b.suggest((tr) => tr.delete(11, 14));
    sync(a, b, c);
    c.review('accept');
    sync(a, b, c);
    same([a, b, c]);
    expect(c.text()).toEqual(['oneX two th']);
    expect(changesOf(a)).toEqual([]);
  });

  it('a third person rejects one author’s changes and accepts the other’s', () => {
    const [a, b, c] = room(['Ada', 'Bo', 'Cy'], [para('one two three')]);
    a.suggest((tr) => tr.insertText('X', 4));
    b.suggest((tr) => tr.delete(11, 14));
    sync(a, b, c);
    c.review('reject', 'Ada');
    c.review('accept', 'Bo');
    sync(a, b, c);
    same([a, b, c]);
    expect(a.text()).toEqual(['one two th']);
    expect(changesOf(b)).toEqual([]);
  });

  it('deleting inside somebody else’s insertion marks it deleted by the second author, and accepting removes it', () => {
    const [a, b, c] = room(['Ada', 'Bo', 'Cy'], [para('abcd')]);
    a.suggest((tr) => tr.insertText('WXYZ', 3));
    sync(a, b, c);
    b.suggest((tr) => tr.delete(4, 6));
    sync(a, b, c);
    same([a, b, c]);
    expect(changesOf(c).map((x) => [x[0], x[1]])).toEqual([['insertion', 'Ada'], ['deletion', 'Bo']]);
    expect(c.text()).toEqual(['abWXYZcd']);
    c.review('accept');
    sync(a, b, c);
    same([a, b, c]);
    expect(a.text()).toEqual(['abWZcd']);
  });
});

describe('settling at the same time', () => {
  const setup = () => {
    const clients = room(['Ada', 'Bo', 'Cy', 'Di'], [para('one two three')]);
    clients[0].suggest((tr) => tr.insertText('XY', 4));
    sync(...clients);
    return clients;
  };

  it('two people accepting the same change leave it accepted', () => {
    const [a, b, c, d] = setup();
    c.review('accept');
    d.review('accept');
    sync(a, b, c, d);
    same([a, b, c, d]);
    expect(a.text()).toEqual(['oneXY two three']);
    expect(changesOf(a)).toEqual([]);
  });

  it('two people rejecting the same change leave it rejected', () => {
    const [a, b, c, d] = setup();
    c.review('reject');
    d.review('reject');
    sync(a, b, c, d);
    same([a, b, c, d]);
    expect(a.text()).toEqual(['one two three']);
  });

  it('one accepting and the other rejecting the same change leave one state that everybody shares', () => {
    const [a, b, c, d] = setup();
    c.review('accept');
    d.review('reject');
    sync(a, b, c, d);
    same([a, b, c, d]);
    // One of the two wins, for everybody, and no mark or text of the change is left behind.
    expect(['one two three', 'oneXY two three']).toContain(a.text()[0]);
    expect(changesOf(a)).toEqual([]);
    expect(changesOf(d)).toEqual([]);
  });

  it('settling a change while its author types next to it converges', () => {
    const [a, b, c] = setup();
    c.review('accept');
    a.suggest((tr) => tr.insertText('Z', 6));
    sync(a, b, c);
    same([a, b, c]);
    // Nothing is lost: the text the author added meanwhile continued the change that was settled, and is settled with it.
    expect(a.text()).toEqual(['oneXYZ two three']);
    expect(changesOf(a)).toEqual([]);
  });

  it('settling one change by its id leaves the others', () => {
    const [a, b, c] = room(['Ada', 'Bo', 'Cy'], [para('one two three')]);
    a.suggest((tr) => tr.insertText('X', 4));
    b.suggest((tr) => tr.insertText('Y', 8));
    sync(a, b, c);
    c.edit((state) => {
      const target = listChanges(state.doc).find((ch) => ch.author === 'Bo')!;
      let out: Transaction | null = null;
      expect(resolveChange(target.id, 'reject')(state, (tr) => (out = tr))).toBe(true);
      return out;
    });
    sync(a, b, c);
    same([a, b, c]);
    expect(a.text()).toEqual(['oneX two three']);
    expect(changesOf(b)).toEqual([['insertion', 'Ada', 'X']]);
  });
});

describe('offline suggestions', () => {
  it('suggestions made while apart merge when the people meet again, whatever the order they meet in', () => {
    const build = () => {
      const [a, b, c] = room(['Ada', 'Bo', 'Cy'], [para('one two three'), para('four five')]);
      // Ada and Bo are offline from here; each makes several suggestions.
      a.suggest((tr) => tr.insertText('A1', 4));
      a.suggest((tr) => tr.insertText('A2', 1));
      b.suggest((tr) => tr.delete(8, 11));
      b.suggest((tr) => tr.insertText('B', 21));
      c.edit((state) => state.tr.insertText('!', state.doc.content.size - 1));
      return [a, b, c];
    };
    const [a1, b1, c1] = build();
    sync(a1, b1, c1);
    same([a1, b1, c1]);
    const [a2, b2, c2] = build();
    // The other way round: Cy meets Bo first, and then Ada.
    sync(c2, b2);
    sync(c2, a2);
    sync(a2, b2, c2);
    same([a2, b2, c2]);
    expect(a2.pm().toJSON()).toEqual(a1.pm().toJSON());
    expect(changesOf(a1).map((x) => x[1]).sort()).toEqual(['Ada', 'Ada', 'Bo', 'Bo']);
    // And a person who was never part of it settles them all.
    c1.review('accept');
    sync(a1, b1, c1);
    same([a1, b1, c1]);
    expect(changesOf(c1)).toEqual([]);
  });

  it('paragraph breaks and joins suggested at once converge, and accepting joins or splits the same everywhere', () => {
    const [a, b, c] = room(['Ada', 'Bo', 'Cy'], [para('one two'), para('three')]);
    a.suggest((tr) => tr.split(4));
    b.suggest((tr) => tr.join(9));
    sync(a, b, c);
    same([a, b, c]);
    // The paragraph mark keeps one record of its change, so of two people's suggestions on the same mark one stays (which one
    // depends on the order Yjs settles concurrent writes in); the document is the same everywhere and no text is lost.
    const open = changesOf(c).map((x) => [x[0], x[1]]);
    expect(open).toHaveLength(1);
    expect([['paragraphInserted', 'Ada'], ['paragraphDeleted', 'Bo']]).toContainEqual(open[0]);
    c.review('accept');
    sync(a, b, c);
    same([a, b, c]);
    expect(changesOf(a)).toEqual([]);
    // (A record that stays with the paragraph element when another person splits it ends up on the first half.)
    expect([['one', ' two', 'three'], ['one two', 'three'], ['one twothree']]).toContainEqual(a.text());
  });
});

describe('the room’s version and people’s colours', () => {
  it('rooms carry the current version, and a client that supports an older one refuses to join', () => {
    // 7 is the version that added the formatChange mark and the revision attributes on top of the tracked insertions and deletions of 6.
    expect(COLLAB_SCHEMA_VERSION).toBe(7);
    const ydoc = new Y.Doc();
    hydrateYDocFromData(ydoc, data([para('x')]));
    expect(readSchemaVersion(ydoc)).toBe(COLLAB_SCHEMA_VERSION);
    expect(() => assertSchemaVersion(ydoc)).not.toThrow();
    // A build that supports 6 (it would drop the formatChange marks) is told the room is newer; one that supports 7 is not.
    expect(() => assertSchemaVersion(ydoc, 6)).toThrow(CollabSchemaVersionError);
    expect(() => assertSchemaVersion(ydoc, 6)).toThrow(/version 7.*up to version 6/);
    expect(() => assertSchemaVersion(ydoc, 7)).not.toThrow();
  });

  it('a Y document with suggestions in it serialises to the model with them', () => {
    const [a] = room(['Ada'], [para('abcd')]);
    a.suggest((tr) => tr.insertText('X', 3));
    const block = serializeYDocToData(a.ydoc).sections[0].blocks[0] as Extract<Block, { type: 'paragraph' }>;
    expect(JSON.stringify(block.content)).toContain('"revision":{"insertion":{"id"');
    expect(JSON.stringify(block.content)).toContain('"author":"Ada"');
  });

  it('paragraph marks and formatting changes that were suggested are in the model the Y document serialises to', () => {
    const [a] = room(['Ada'], [para('abcd')]);
    a.suggest((tr) => tr.split(3));
    a.suggest((tr) => tr.addMark(1, 3, docsSchema.marks.bold.create()));
    const [first] = serializeYDocToData(a.ydoc).sections[0].blocks as Array<Extract<Block, { type: 'paragraph' }>>;
    expect(first.revision).toMatchObject({ insertion: { author: 'Ada' } });
    expect(JSON.stringify(first.content)).toContain('"formatChange":{');
    expect(JSON.stringify(first.content)).toContain('"previous":{}');
  });

  it('people in the room are drawn in the colour of their cursor, and nobody else takes that colour', () => {
    const awareness = new Awareness(new Y.Doc());
    awareness.setLocalStateField('user', { userId: 'u1', name: 'Ada', color: '#00AA00' });
    const colors = collabAuthorColors({ awareness });
    expect(colors.get('Ada')).toBe('#00AA00');
    const assigned = assignAuthorColors(['Ada', 'Bo', 'Cy'], colors);
    expect(assigned.get('Ada')).toBe('#00aa00');
    expect(new Set(assigned.values()).size).toBe(3);
    // A state with no user, or a colour that is not a six-digit hex colour, is ignored.
    awareness.setLocalStateField('user', { name: 'Bo', color: 'red' });
    expect(assignAuthorColors(['Bo'], collabAuthorColors({ awareness })).get('Bo')).not.toBe('red');
    const none = new Awareness(new Y.Doc());
    none.setLocalState({});
    expect(collabAuthorColors({ awareness: none }).size).toBe(0);
    // Two people under one name are one author, in the smallest of their colours, whatever order the room lists them in.
    const twins = new Awareness(new Y.Doc());
    twins.setLocalStateField('user', { name: 'Ann', color: '#BB0000' });
    const other = new Awareness(new Y.Doc());
    other.setLocalStateField('user', { name: 'Ann', color: '#aa0000' });
    const states = (...all: Awareness[]) => ({ awareness: { getStates: () => new Map(all.map((aw, i) => [i, aw.getLocalState() as unknown]))} as unknown as Awareness });
    expect(collabAuthorColors(states(twins, other)).get('Ann')).toBe('#aa0000');
    expect(collabAuthorColors(states(other, twins)).get('Ann')).toBe('#aa0000');
    // A colour that is not a hex colour never beats one that is.
    other.setLocalStateField('user', { name: 'Ann', color: '!!' });
    expect(collabAuthorColors(states(other, twins)).get('Ann')).toBe('#BB0000');
    expect(collabAuthorColors(states(twins, other)).get('Ann')).toBe('#BB0000');
    for (const aw of [awareness, none, twins, other]) aw.destroy();
    // The same colour for two people goes to the first (by name) only.
    expect(new Set(assignAuthorColors(['A', 'B'], new Map([['A', '#112233'], ['B', '#112233']])).values()).size).toBe(2);
  });
});
