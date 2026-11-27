import { DeckImpl } from '@b32nio/spindle-slides-core';
import { expandGroups } from './gesture';
import { copyElements, pasteElements } from '../hooks/useClipboard';

function build() {
  const deck = new DeckImpl('d');
  const s = deck.getActiveSlideId();
  const ids = [0, 1, 2, 3].map((i) => deck.addElement(s, { type: 'shape', x: i * 60, y: 0, w: 40, h: 40 }).id);
  return { deck, s, ids };
}

describe('nested groups in the editor', () => {
  it('expandGroups selects the outermost group as one', () => {
    const { deck, ids: [a, b, c, e] } = build();
    deck.groupElements([a, b]);
    deck.groupElements([a, c]);
    expect(expandGroups(deck, [b]).sort()).toEqual([a, b, c].sort());
    expect(expandGroups(deck, [e])).toEqual([e]);
    expect(expandGroups(deck, ['missing'])).toEqual([]);
  });

  it('copy and paste re-forms the nesting with fresh group ids', () => {
    const { deck, s, ids: [a, b, c, e] } = build();
    const inner = deck.groupElements([a, b])!;
    const outer = deck.groupElements([a, c])!;
    deck.setSelection({ slideId: s, elementIds: [a, b, c, e] });
    copyElements(deck);
    const before = new Set(deck.getElementsForSlide(s).map((x) => x.id));
    pasteElements(deck);
    const pasted = deck.getElementsForSlide(s).filter((x) => !before.has(x.id));
    expect(pasted).toHaveLength(4);
    const grouped = pasted.filter((x) => x.groupId);
    expect(grouped).toHaveLength(3);
    const [pa, pb, pc] = pasted;
    expect(deck.getGroupChain(pa.id)).toHaveLength(2);
    expect(deck.getGroupChain(pa.id)).toEqual(deck.getGroupChain(pb.id));
    expect(deck.getGroupChain(pc.id)).toEqual([deck.getGroupChain(pa.id)[1]]);
    expect(deck.getGroupChain(pa.id)).not.toContain(inner);
    expect(deck.getGroupChain(pa.id)).not.toContain(outer);
    expect(pasted[3].groupId).toBeUndefined();
    // The source nesting is untouched.
    expect(deck.getGroupChain(a)).toEqual([inner, outer]);
  });

  it('pasting nothing and copying an empty selection do nothing', () => {
    const { deck, s } = build();
    deck.setSelection({ slideId: s, elementIds: [] });
    copyElements(deck);
    expect(() => pasteElements(deck)).not.toThrow();
  });
});
