import { Awareness } from 'y-protocols/awareness';
import * as Y from 'yjs';
import { sameAuthorColors, suggestionAuthor, watchAuthorColors } from './collab-authors';

describe('suggestionAuthor', () => {
  it('is the name the user joined the room under, else the current user', () => {
    expect(suggestionAuthor({ identity: { displayName: 'Ada (room)' } } as never, 'Alice')).toBe('Ada (room)');
    expect(suggestionAuthor(null, 'Alice')).toBe('Alice');
    expect(suggestionAuthor(undefined, 'Alice')).toBe('Alice');
    expect(suggestionAuthor({ identity: {} } as never, 'Alice')).toBe('Alice');
  });
});

describe('sameAuthorColors', () => {
  it('compares names and colours', () => {
    const a = new Map([['Ada', '#111111'], ['Bo', '#222222']]);
    expect(sameAuthorColors(a, new Map([['Bo', '#222222'], ['Ada', '#111111']]))).toBe(true);
    expect(sameAuthorColors(a, new Map([['Ada', '#111111']]))).toBe(false);
    expect(sameAuthorColors(a, new Map([['Ada', '#111111'], ['Bo', '#333333']]))).toBe(false);
    expect(sameAuthorColors(new Map(), new Map())).toBe(true);
  });
});

describe('watchAuthorColors', () => {
  const awareness = () => new Awareness(new Y.Doc());

  it('tells the colours of the room now, and again when somebody joins or changes colour', () => {
    const aw = awareness();
    aw.setLocalStateField('user', { name: 'Ada', color: '#aa0000' });
    const seen: Array<Record<string, string>> = [];
    const stop = watchAuthorColors({ awareness: aw }, (colors) => seen.push(Object.fromEntries(colors)));
    expect(seen).toEqual([{ Ada: '#aa0000' }]);
    aw.setLocalStateField('user', { name: 'Ada', color: '#00aa00' });
    expect(seen[seen.length - 1]).toEqual({ Ada: '#00aa00' });
    stop();
    aw.destroy();
  });

  it('stops telling once stopped, and the caller can ignore a cursor that only moved', () => {
    const aw = awareness();
    aw.setLocalStateField('user', { name: 'Ada', color: '#aa0000' });
    let current: ReadonlyMap<string, string> = new Map();
    let renders = 0;
    const stop = watchAuthorColors({ awareness: aw }, (colors) => {
      if (!sameAuthorColors(current, colors)) {
        current = colors;
        renders++;
      }
    });
    expect(renders).toBe(1);
    aw.setLocalStateField('cursor', { anchor: 1, head: 2 });
    aw.setLocalStateField('cursor', { anchor: 3, head: 3 });
    expect(renders).toBe(1);
    stop();
    aw.setLocalStateField('user', { name: 'Ada', color: '#0000aa' });
    expect(renders).toBe(1);
    aw.destroy();
  });
});
