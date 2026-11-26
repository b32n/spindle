import { Schema } from 'prosemirror-model';
import { findBookmarkPos, findElementAtPos, internalLinkTarget } from './internal-links';

const schema = new Schema({
  nodes: {
    doc: { content: 'block+' },
    paragraph: { group: 'block', content: 'inline*', toDOM: () => ['p', 0] },
    text: { group: 'inline' },
    bookmark: { inline: true, group: 'inline', atom: true, attrs: { name: {} } },
  },
});

describe('internalLinkTarget', () => {
  it('returns the bookmark name of a #name href and null for anything else', () => {
    expect(internalLinkTarget('#_Toc1')).toBe('_Toc1');
    expect(internalLinkTarget('#a%20b')).toBe('a b');
    expect(internalLinkTarget('#%E0%A4%A')).toBe('%E0%A4%A'); // not decodable: used as written
    for (const href of ['#', '', null, undefined, 'https://e.com/#x', 'page.html#x']) expect(internalLinkTarget(href)).toBeNull();
  });
});

describe('findBookmarkPos', () => {
  const doc = schema.nodeFromJSON({
    type: 'doc',
    content: [
      { type: 'paragraph', content: [{ type: 'text', text: 'ab' }] },
      { type: 'paragraph', content: [{ type: 'text', text: 'c' }, { type: 'bookmark', attrs: { name: 'x' } }, { type: 'bookmark', attrs: { name: 'x' } }] },
    ],
  });

  it('finds the position of the first bookmark with the name', () => {
    // paragraph 1 occupies 0..4, paragraph 2 opens at 4, its text at 5, the bookmark at 6
    expect(findBookmarkPos(doc, 'x')).toBe(6);
  });

  it('returns null for an unknown name', () => {
    expect(findBookmarkPos(doc, 'nope')).toBeNull();
  });
});

describe('findElementAtPos', () => {
  // No DOM in this test environment: stand-ins with the attributes the painter writes.
  const el = (id: string, start: number, end: number) => ({ id, dataset: { pmStart: String(start), pmEnd: String(end) } });
  const container = {
    querySelectorAll: (_selector: string) => ({
      forEach: (fn: (e: unknown) => void) => [el('block', 4, 10), el('a', 5, 6), el('b', 6, 7)].forEach(fn),
    }),
  } as unknown as ParentNode;
  const idAt = (pos: number) => (findElementAtPos(container, pos) as unknown as { id: string } | null)?.id ?? null;

  it('picks the smallest painted element covering the position, and null when none does', () => {
    expect(idAt(6)).toBe('b');
    expect(idAt(5)).toBe('a');
    expect(idAt(9)).toBe('block');
    expect(idAt(20)).toBeNull();
  });
});
