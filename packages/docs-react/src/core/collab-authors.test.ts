import { Awareness } from 'y-protocols/awareness';
import * as Y from 'yjs';
import { authorColorsKey, authorsWithColors, sameAuthorColors, suggestionAuthor, watchAuthorColors } from './collab-authors';
import { assignAuthorColors } from '@b32nio/spindle-docs-core';

describe('suggestionAuthor', () => {
  it('is the name the user joined the room under, else the current user', () => {
    expect(suggestionAuthor({ identity: { displayName: 'Ada (room)' } } as never, 'Alice')).toBe('Ada (room)');
    expect(suggestionAuthor(null, 'Alice')).toBe('Alice');
    expect(suggestionAuthor(undefined, 'Alice')).toBe('Alice');
    expect(suggestionAuthor({ identity: {} } as never, 'Alice')).toBe('Alice');
    // A blank or whitespace name would write suggestions with no author.
    expect(suggestionAuthor({ identity: { displayName: '' } } as never, 'Alice')).toBe('Alice');
    expect(suggestionAuthor({ identity: { displayName: '   ' } } as never, 'Alice')).toBe('Alice');
    expect(suggestionAuthor({ identity: { displayName: ' Ada ' } } as never, 'Alice')).toBe('Ada');
    // Never an empty author: it would be written as w:author="" and shared with the next blank author.
    expect(suggestionAuthor({ identity: { displayName: ' ' } } as never, '')).toBe('Unknown');
    expect(suggestionAuthor(null, '  ')).toBe('Unknown');
  });
});

describe('authorColorsKey', () => {
  it('ignores the order the map was built in', () => {
    expect(authorColorsKey(new Map([['Ada', '#111111'], ['Bo', '#222222']]))).toBe(authorColorsKey(new Map([['Bo', '#222222'], ['Ada', '#111111']])));
  });

  it('differs when a colour or a name differs', () => {
    const base = authorColorsKey(new Map([['Ada', '#111111']]));
    expect(authorColorsKey(new Map([['Ada', '#111112']]))).not.toBe(base);
    expect(authorColorsKey(new Map([['Ade', '#111111']]))).not.toBe(base);
    expect(authorColorsKey(new Map([['Ada', '#111111'], ['Bo', '#222222']]))).not.toBe(base);
  });

  it('is empty for no map', () => {
    expect(authorColorsKey(undefined)).toBe('');
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

describe('authorsWithColors', () => {
  const counts = [
    { name: 'Ada', color: '#000000', count: 2 },
    { name: 'Bo', color: '#000000', count: 1 },
  ];

  it('draws people in the room in their cursor colour, and the others in a colour the document’s authors share out', () => {
    const present = authorsWithColors(counts, ['Ada', 'Bo'], new Map([['Ada', '#00AA00']]));
    expect(present.map((a) => a.color)[0]).toBe('#00aa00');
    expect(present[1].color).toBe(assignAuthorColors(['Ada', 'Bo'], new Map([['Ada', '#00AA00']])).get('Bo'));
    expect(present.map((a) => a.count)).toEqual([2, 1]);
  });

  it('gives an author a colour from the palette once they have left the room, so the colour can change', () => {
    const inRoom = authorsWithColors(counts, ['Ada', 'Bo'], new Map([['Bo', '#00AA00']]));
    const left = authorsWithColors(counts, ['Ada', 'Bo'], new Map());
    expect(inRoom[1].color).toBe('#00aa00');
    expect(left[1].color).toBe(assignAuthorColors(['Ada', 'Bo']).get('Bo'));
    expect(left[1].color).not.toBe(inRoom[1].color);
  });

  it('keeps the colour it was given for an author the palette does not list', () => {
    expect(authorsWithColors([{ name: 'Cy', color: '#123456', count: 1 }], [], new Map())[0].color).toBe('#123456');
  });
});
