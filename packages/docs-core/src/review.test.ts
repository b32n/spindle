// Accepting and rejecting tracked changes: what each kind does, in one transaction, and the list of changes.

import { EditorState, TextSelection } from 'prosemirror-state';
import type { Node as PmNode } from 'prosemirror-model';
import { docsSchema } from './prosemirror/schema';
import { blocksToPmDoc } from './prosemirror/sync';
import { createPlugins } from './prosemirror/plugins';
import { readBlockRevision } from './prosemirror/block-revision';
import { changeUnits, listChanges, resolveAll, resolveChange, resolveSelection, settle } from './prosemirror/review';
import { assignAuthorColors, documentRevisionAuthors, revisionAuthors } from './revisions';
import type { Block, InlineContent, RevisionMeta } from './types';

const ADA: RevisionMeta = { id: '1', author: 'Ada', date: '2024-01-01T00:00:00Z' };
const BO: RevisionMeta = { id: '2', author: 'Bo' };
const run = (text: string, revision?: object, extra: object = {}): InlineContent => ({ type: 'text', text, ...extra, ...(revision && { revision }) }) as InlineContent;
const para = (content: InlineContent[], extra: object = {}): Block => ({ id: 'p', type: 'paragraph', content, ...extra }) as Block;

const stateOf = (blocks: Block[]) => EditorState.create({ doc: blocksToPmDoc(blocks) });
/** Apply a command (which must apply) and return the new state. */
function apply(state: EditorState, command: (s: EditorState, d?: (tr: import('prosemirror-state').Transaction) => void) => boolean): EditorState {
  let next = state;
  expect(command(state, (tr) => (next = state.apply(tr)))).toBe(true);
  return next;
}
const texts = (doc: PmNode) => {
  const out: string[] = [];
  doc.forEach((block) => out.push(block.textContent));
  return out;
};
const marks = (doc: PmNode) => {
  const out: string[] = [];
  doc.descendants((node) => {
    if (node.isText) out.push(`${node.text}:${node.marks.map((m) => m.type.name).filter((n) => ['insertion', 'deletion', 'formatChange'].includes(n)).join('+')}`);
  });
  return out;
};

describe('accepting and rejecting changes to text', () => {
  const blocks = [para([run('keep '), run('new', { insertion: ADA }), run('old', { deletion: BO }), run(' end')])];

  it('accepting all keeps inserted text and drops deleted text, with no marks left', () => {
    const next = apply(stateOf(blocks), resolveAll('accept'));
    expect(texts(next.doc)).toEqual(['keep new end']);
    expect(marks(next.doc).every((m) => m.endsWith(':'))).toBe(true);
    expect(changeUnits(next.doc)).toEqual([]);
  });

  it('rejecting all drops inserted text and keeps deleted text, with no marks left', () => {
    const next = apply(stateOf(blocks), resolveAll('reject'));
    expect(texts(next.doc)).toEqual(['keep old end']);
    expect(changeUnits(next.doc)).toEqual([]);
  });

  it('settles text that one author inserted and another deleted, whichever way', () => {
    const both = [para([run('a'), run('x', { insertion: ADA, deletion: BO }), run('b')])];
    expect(texts(apply(stateOf(both), resolveAll('accept')).doc)).toEqual(['ab']);
    // Rejecting the insertion takes the text away whatever else happened to it.
    expect(texts(apply(stateOf(both), resolveAll('reject')).doc)).toEqual(['ab']);
    // Accepting only the deletion removes the text, and with it the insertion.
    const byBo = apply(stateOf(both), resolveAll('accept', 'Bo'));
    expect(texts(byBo.doc)).toEqual(['ab']);
    // Rejecting only the insertion removes the text; rejecting only the deletion keeps it inserted.
    expect(texts(apply(stateOf(both), resolveAll('reject', 'Ada')).doc)).toEqual(['ab']);
    expect(marks(apply(stateOf(both), resolveAll('reject', 'Bo')).doc)).toEqual(['a:', 'x:insertion', 'b:']);
  });

  it('acts on one author, one change, or the selection', () => {
    const state = stateOf(blocks);
    const byAda = apply(state, resolveAll('accept', 'Ada'));
    expect(marks(byAda.doc)).toEqual(['keep new:', 'old:deletion', ' end:']);
    const changes = listChanges(state.doc);
    expect(changes.map((c) => [c.kind, c.author, c.text])).toEqual([['insertion', 'Ada', 'new'], ['deletion', 'Bo', 'old']]);
    const second = apply(state, resolveChange(changes[1].id, 'reject'));
    expect(marks(second.doc)).toEqual(['keep :', 'new:insertion', 'old end:']);
    // A selection that touches only the deletion.
    const at = state.doc.textContent.indexOf('old') + 1 + 1;
    const selected = state.apply(state.tr.setSelection(TextSelection.create(state.doc, at, at + 1)));
    expect(texts(apply(selected, resolveSelection('accept')).doc)).toEqual(['keep new end']);
    // A caret inside an insertion settles it.
    const caret = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 8)));
    expect(texts(apply(caret, resolveSelection('reject')).doc)).toEqual(['keep old end']);
  });

  it('a selection that exactly covers one change leaves the changes next to it alone', () => {
    const state = stateOf(blocks); // "keep " 1-6, "new" 6-9 (insertion), "old" 9-12 (deletion), " end"
    const covering = (from: number, to: number) => state.apply(state.tr.setSelection(TextSelection.create(state.doc, from, to)));
    expect(marks(apply(covering(6, 9), resolveSelection('accept')).doc)).toEqual(['keep new:', 'old:deletion', ' end:']);
    expect(marks(apply(covering(9, 12), resolveSelection('accept')).doc)).toEqual(['keep :', 'new:insertion', ' end:']);
    // A range that only reaches the edge of a change from outside does not touch it.
    expect(resolveSelection('accept')(covering(1, 6))).toBe(false);
    expect(resolveSelection('accept')(covering(12, 15))).toBe(false);
    // A range that overlaps two changes settles both; a caret at the edge between two settles both (it is at the edge of each).
    expect(texts(apply(covering(8, 10), resolveSelection('accept')).doc)).toEqual(['keep new end']);
    expect(texts(apply(covering(9, 9), resolveSelection('accept')).doc)).toEqual(['keep new end']);
  });

  it('a selection that ends where a paragraph starts does not settle the change of that paragraph mark', () => {
    const state = stateOf([para([run('one')]), para([run('two')], { revision: { insertion: ADA } })]); // "one" 0-5, "two" 5-10
    const selecting = (from: number, to: number) => state.apply(state.tr.setSelection(TextSelection.create(state.doc, from, to)));
    expect(listChanges(state.doc).map((c) => c.kind)).toEqual(['paragraphInserted']);
    expect(resolveSelection('accept')(selecting(1, 5))).toBe(false);
    const inside = apply(selecting(1, 7), resolveSelection('accept'));
    expect(readBlockRevision(inside.doc.child(1).attrs)).toBeUndefined();
  });

  it('does nothing, and says so, when there is nothing to settle', () => {
    const plain = stateOf([para([run('plain')])]);
    expect(resolveAll('accept')(plain, () => undefined)).toBe(false);
    expect(resolveChange('ins:missing@0', 'accept')(plain)).toBe(false);
    expect(settle(plain, [], 'accept')).toBeNull();
    expect(resolveSelection('accept')(plain)).toBe(false);
  });

  it('is one transaction, so one undo takes it back', () => {
    const state = stateOf(blocks);
    let count = 0;
    resolveAll('accept')(state, () => count++);
    expect(count).toBe(1);
  });
});

describe('moves', () => {
  const blocks = [
    para([run('a '), run('moved', { deletion: { id: '5', author: 'Cy', moveId: 'm1' } })]),
    para([run('b '), run('moved', { insertion: { id: '6', author: 'Cy', moveId: 'm1' } })]),
  ];

  it('are one change whose two ends are settled together', () => {
    const state = stateOf(blocks);
    const changes = listChanges(state.doc);
    expect(changes.map((c) => [c.kind, c.text, c.units.length])).toEqual([['move', 'moved', 2]]);
    // The change is where the text went to.
    expect(state.doc.textBetween(changes[0].from, changes[0].to)).toBe('moved');
    expect(changes[0].from).toBeGreaterThan(state.doc.child(0).nodeSize);
    expect(texts(apply(state, resolveChange(changes[0].id, 'accept')).doc)).toEqual(['a ', 'b moved']);
    expect(texts(apply(state, resolveChange(changes[0].id, 'reject')).doc)).toEqual(['a moved', 'b ']);
    expect(texts(apply(state, resolveAll('accept')).doc)).toEqual(['a ', 'b moved']);
  });

  it('stay separate changes when the other end is missing', () => {
    const state = stateOf([blocks[0]]);
    expect(listChanges(state.doc).map((c) => c.kind)).toEqual(['move']);
    expect(texts(apply(state, resolveAll('accept')).doc)).toEqual(['a ']);
  });
});

describe('formatting changes of text', () => {
  const blocks = [para([run('now', { formatChange: { id: '7', author: 'Di', previous: { italic: true, fontSize: 14, color: '#c00000' } } }, { bold: true, color: '#00ff00' })])];
  const markNames = (state: EditorState) => state.doc.firstChild!.firstChild!.marks.map((m) => m.type.name).sort();

  it('are accepted by dropping the record, and the new formatting stays', () => {
    const next = apply(stateOf(blocks), resolveAll('accept'));
    expect(markNames(next)).toEqual(['bold', 'textStyle']);
    expect(next.doc.firstChild!.firstChild!.marks.find((m) => m.type.name === 'textStyle')!.attrs.color).toBe('#00ff00');
  });

  it('are rejected by putting the old formatting back', () => {
    const next = apply(stateOf(blocks), resolveAll('reject'));
    expect(markNames(next)).toEqual(['italic', 'textStyle']);
    const style = next.doc.firstChild!.firstChild!.marks.find((m) => m.type.name === 'textStyle')!.attrs;
    expect(style).toMatchObject({ color: '#c00000', fontSize: 14 });
  });

  it('bring back a character style, and leave links and comments alone', () => {
    const doc = [
      para([{ type: 'link', text: 'go', href: 'https://example.com/', commentThreadId: 't1', bold: true, revision: { formatChange: { id: '8', author: 'Di', previous: { namedStyleId: 'Strong' } } } } as InlineContent]),
    ];
    const next = apply(stateOf(doc), resolveAll('reject'));
    expect(markNames(next)).toEqual(['comment', 'link', 'namedStyle']);
  });
});

describe('changes of paragraph marks', () => {
  const two = (first: object, firstExtra: object = {}) => [
    para([run('one')], { revision: first, ...firstExtra }),
    para([run('two')]),
  ];

  it('an inserted mark stays when accepted and is joined away when rejected', () => {
    const blocks = two({ insertion: ADA });
    expect(texts(apply(stateOf(blocks), resolveAll('accept')).doc)).toEqual(['one', 'two']);
    expect(readBlockRevision(apply(stateOf(blocks), resolveAll('accept')).doc.child(0).attrs)).toBeUndefined();
    expect(texts(apply(stateOf(blocks), resolveAll('reject')).doc)).toEqual(['onetwo']);
  });

  it('a deleted mark is joined away when accepted and stays when rejected', () => {
    const blocks = two({ deletion: BO });
    expect(texts(apply(stateOf(blocks), resolveAll('accept')).doc)).toEqual(['onetwo']);
    expect(texts(apply(stateOf(blocks), resolveAll('reject')).doc)).toEqual(['one', 'two']);
    expect(readBlockRevision(apply(stateOf(blocks), resolveAll('reject')).doc.child(0).attrs)).toBeUndefined();
  });

  it('joining keeps the first paragraph, also when the second is a heading', () => {
    const blocks: Block[] = [para([run('one')], { revision: { deletion: BO }, alignment: 'center' }), { id: 'h', type: 'heading', level: 2, content: [run('Two')] } as Block];
    const next = apply(stateOf(blocks), resolveAll('accept'));
    expect(next.doc.childCount).toBe(1);
    expect(next.doc.child(0)).toMatchObject({ type: { name: 'paragraph' }, textContent: 'oneTwo' });
    expect(next.doc.child(0).attrs.alignment).toBe('center');
  });

  it('only clears the change when there is no paragraph after it to join', () => {
    const blocks: Block[] = [para([run('last')], { revision: { deletion: BO } })];
    const next = apply(stateOf(blocks), resolveAll('accept'));
    expect(texts(next.doc)).toEqual(['last']);
    expect(readBlockRevision(next.doc.child(0).attrs)).toBeUndefined();
    // In a list item, the next paragraph is in another item.
    const list: Block[] = [
      { id: 'l1', type: 'list-item', listType: 'bullet', level: 0, content: [run('a')], revision: { insertion: ADA } } as Block,
      { id: 'l2', type: 'list-item', listType: 'bullet', level: 0, content: [run('b')] } as Block,
    ];
    const next2 = apply(stateOf(list), resolveAll('reject'));
    // The result of the command, not the fixture: both items stay and the first one's change is cleared.
    expect(next2.doc.child(0).childCount).toBe(2);
    expect(readBlockRevision(next2.doc.child(0).child(0).child(0).attrs)).toBeUndefined();
    expect(next2.doc.child(0).textContent).toBe('ab');
    expect(readBlockRevision(stateOf(list).doc.child(0).child(0).child(0).attrs)).toMatchObject({ insertion: { id: '1' } });
  });

  it('settling one deleted mark of several leaves the next one open, with its change', () => {
    const CY = { id: '3', author: 'Cy' };
    const blocks: Block[] = [
      para([run('one')], { revision: { deletion: BO } }),
      para([run('two')], { revision: { deletion: CY } }),
      para([run('three')]),
    ];
    const state = stateOf(blocks);
    const first = listChanges(state.doc).find((c) => c.author === 'Bo')!;
    const next = apply(state, resolveChange(first.id, 'accept'));
    expect(texts(next.doc)).toEqual(['onetwo', 'three']);
    // The paragraph that now ends with the second mark still carries its deletion, and the panel still lists it.
    expect(readBlockRevision(next.doc.child(0).attrs)).toEqual({ deletion: CY });
    expect(listChanges(next.doc).map((c) => [c.kind, c.author])).toEqual([['paragraphDeleted', 'Cy']]);
    expect(texts(apply(next, resolveAll('accept')).doc)).toEqual(['onetwothree']);
  });

  it('joining takes the second mark\'s changes and the first one\'s formatting change, and drops the second one\'s formatting change', () => {
    const CY = { id: '3', author: 'Cy' };
    const fmt = (author: string) => ({ id: author, author, previous: { alignment: 'right' } });
    const blocks: Block[] = [
      para([run('one')], { revision: { deletion: BO, formatChange: fmt('Ann') } }),
      para([run('two')], { revision: { insertion: CY, formatChange: fmt('Dee') } }),
    ];
    const next = apply(stateOf(blocks), resolveChange(listChanges(stateOf(blocks).doc).find((c) => c.kind === 'paragraphDeleted')!.id, 'accept'));
    expect(texts(next.doc)).toEqual(['onetwo']);
    expect(readBlockRevision(next.doc.child(0).attrs)).toEqual({ insertion: CY, formatChange: fmt('Ann') });
  });

  it('keeps a mark that is both inserted and deleted until both are settled', () => {
    const blocks = two({ insertion: ADA, deletion: BO });
    const byAda = apply(stateOf(blocks), resolveAll('accept', 'Ada'));
    expect(readBlockRevision(byAda.doc.child(0).attrs)).toEqual({ deletion: { id: '2', author: 'Bo' } });
    expect(texts(apply(stateOf(blocks), resolveAll('accept')).doc)).toEqual(['onetwo']);
  });
});

describe('changes of a paragraph’s formatting', () => {
  const formatted = (previous: object): Block[] => [
    para([run('text')], { alignment: 'center', revision: { formatChange: { id: '9', author: 'Ada', previous } } }),
  ];

  it('are accepted by dropping the record', () => {
    const next = apply(stateOf(formatted({ alignment: 'right' })), resolveAll('accept'));
    expect(next.doc.child(0).attrs).toMatchObject({ alignment: 'center', revisionFormat: null });
  });

  it('are rejected by putting the old formatting back', () => {
    const next = apply(stateOf(formatted({ alignment: 'right', leftIndent: 24, spaceAfter: 0 })), resolveAll('reject'));
    expect(next.doc.child(0).attrs).toMatchObject({ alignment: 'right', leftIndent: 24, spaceAfter: 0, revisionFormat: null });
    expect(next.doc.child(0).textContent).toBe('text');
  });

  it('turn a paragraph back into the heading it was, and a heading back into a paragraph', () => {
    const toHeading = apply(stateOf(formatted({ headingLevel: 3 })), resolveAll('reject'));
    expect(toHeading.doc.child(0)).toMatchObject({ type: { name: 'heading' }, attrs: { level: 3, revisionFormat: null }, textContent: 'text' });
    const heading: Block[] = [{ id: 'h', type: 'heading', level: 1, content: [run('Title')], revision: { formatChange: { id: '1', author: 'Ada', previous: {} } } } as Block];
    const toParagraph = apply(stateOf(heading), resolveAll('reject'));
    expect(toParagraph.doc.child(0).type.name).toBe('paragraph');
    expect(toParagraph.doc.child(0).attrs.indent).toBe(0);
  });

  it('rejecting a change from a heading in a list item keeps a paragraph there (a list item starts with one) and does not fail', () => {
    const blocks: Block[] = [
      { id: 'l1', type: 'list-item', listType: 'bullet', level: 0, content: [run('a')], revision: { formatChange: { id: '9', author: 'Ada', previous: { headingLevel: 2, alignment: 'center' } } } } as Block,
    ];
    const next = apply(stateOf(blocks), resolveAll('reject'));
    const item = next.doc.child(0).child(0).child(0);
    expect(item.type.name).toBe('paragraph');
    expect(item.attrs).toMatchObject({ alignment: 'center', revisionFormat: null });
    expect(item.textContent).toBe('a');
  });

  it('rejecting a change in a paragraph that stays a paragraph keeps its indent level (it is not part of the recorded formatting)', () => {
    const blocks: Block[] = [para([run('t')], { indent: 2, revision: { formatChange: { id: '9', author: 'Ada', previous: { alignment: 'right' } } } })];
    const next = apply(stateOf(blocks), resolveAll('reject'));
    expect(next.doc.child(0).attrs).toMatchObject({ alignment: 'right', indent: 2 });
  });

  it('rejecting a format change leaves the paragraph mark records of the other kinds alone', () => {
    const blocks: Block[] = [para([run('t')], { revision: { insertion: ADA, deletion: BO, formatChange: { id: '9', author: 'Cy', previous: { alignment: 'right' } } } }), para([run('u')])];
    const next = apply(stateOf(blocks), resolveAll('reject', 'Cy'));
    expect(next.doc.child(0).attrs).toMatchObject({ alignment: 'right', revisionFormat: null });
    expect(readBlockRevision(next.doc.child(0).attrs)).toEqual({ insertion: ADA, deletion: BO });
  });

  it('keep the paragraph mark change that is still open', () => {
    const blocks: Block[] = [para([run('t')], { revision: { insertion: ADA, formatChange: { id: '9', author: 'Bo', previous: { alignment: 'right' } } } }), para([run('u')])];
    const next = apply(stateOf(blocks), resolveAll('reject', 'Bo'));
    expect(readBlockRevision(next.doc.child(0).attrs)).toEqual({ insertion: { id: '1', author: 'Ada', date: '2024-01-01T00:00:00Z' } });
    expect(next.doc.child(0).attrs.alignment).toBe('right');
  });
});

describe('line breaks inside tracked text', () => {
  const blocks = [para([run('a'), run('x\ny', { deletion: BO }), run('z', { insertion: ADA }), run('p\nq', { insertion: ADA })])];

  it('are part of the change: one change for the text and its break, and the break goes with the text', () => {
    const changes = listChanges(stateOf(blocks).doc);
    expect(changes.map((c) => [c.kind, c.text])).toEqual([['deletion', 'x\ny'], ['insertion', 'zp\nq']]);
    expect(texts(apply(stateOf(blocks), resolveAll('accept')).doc)).toEqual(['azp\nq']);
    expect(texts(apply(stateOf(blocks), resolveAll('reject')).doc)).toEqual(['ax\ny']);
  });
});

describe('settling the changes of one author', () => {
  it('settles both ends of a move whose ends are by different authors, never one end alone', () => {
    const blocks = [
      para([run('a '), run('moved', { deletion: { id: '5', author: 'Bo', moveId: 'm1' } })]),
      para([run('b '), run('moved', { insertion: { id: '6', author: 'Ada', moveId: 'm1' } })]),
    ];
    // The move is Ada's (the end it went to): whoever is settled, it goes whole or not at all.
    expect(texts(apply(stateOf(blocks), resolveAll('accept', 'Ada')).doc)).toEqual(['a ', 'b moved']);
    expect(texts(apply(stateOf(blocks), resolveAll('reject', 'Ada')).doc)).toEqual(['a moved', 'b ']);
    expect(resolveAll('accept', 'Bo')(stateOf(blocks), undefined)).toBe(false);
  });
});

describe('changes of table rows', () => {
  const cell = (text: string) => ({ id: `c${text}`, content: [run(text)] });
  const table = (rows: Array<{ text: string; revision?: object }>): Block[] => [
    { id: 't', type: 'table', rows: rows.map((r, i) => ({ id: `r${i}`, cells: [cell(r.text)], ...(r.revision && { revision: r.revision }) })) } as Block,
    para([run('after')]),
  ];
  const rowTexts = (state: EditorState) => {
    const out: string[] = [];
    state.doc.firstChild!.forEach((row) => out.push(row.textContent));
    return out;
  };

  it('an inserted row stays when accepted and goes when rejected', () => {
    const blocks = table([{ text: 'a' }, { text: 'new', revision: { insertion: ADA } }]);
    expect(rowTexts(apply(stateOf(blocks), resolveAll('accept')))).toEqual(['a', 'new']);
    expect(readBlockRevision(apply(stateOf(blocks), resolveAll('accept')).doc.firstChild!.child(1).attrs)).toBeUndefined();
    expect(rowTexts(apply(stateOf(blocks), resolveAll('reject')))).toEqual(['a']);
  });

  it('a deleted row goes when accepted and stays when rejected', () => {
    const blocks = table([{ text: 'a' }, { text: 'old', revision: { deletion: BO } }]);
    expect(rowTexts(apply(stateOf(blocks), resolveAll('accept')))).toEqual(['a']);
    const kept = apply(stateOf(blocks), resolveAll('reject'));
    expect(rowTexts(kept)).toEqual(['a', 'old']);
    expect(readBlockRevision(kept.doc.firstChild!.child(1).attrs)).toBeUndefined();
  });

  it('removes the table when its only row goes', () => {
    const blocks = table([{ text: 'only', revision: { deletion: BO } }]);
    const next = apply(stateOf(blocks), resolveAll('accept'));
    expect(texts(next.doc)).toEqual(['after']);
  });

  it('settles the text in a removed row without disturbing the rest, and lists the row as a change', () => {
    const blocks = [
      ...table([
        { text: 'a' },
        { text: 'old', revision: { deletion: BO } },
        { text: 'c' },
      ]).slice(0, 1),
      para([run('x', { insertion: ADA }), run('y')]),
    ];
    const state = stateOf(blocks);
    expect(listChanges(state.doc).map((c) => [c.kind, c.author])).toEqual([['rowDeleted', 'Bo'], ['insertion', 'Ada']]);
    const next = apply(state, resolveAll('accept'));
    expect(rowTexts(next)).toEqual(['a', 'c']);
    expect(next.doc.child(1).textContent).toBe('xy');
    expect(changeUnits(next.doc)).toEqual([]);
  });
});

describe('listing changes', () => {
  it('joins the runs of one change that touch, keeps changes with the same id apart when they do not, and gives each change an id of its own', () => {
    const state = stateOf([
      para([run('a', { insertion: ADA }), run('b', { insertion: ADA }, { bold: true }), run('-'), run('c', { insertion: ADA })]),
      para([run('d', { deletion: { ...ADA } })]),
    ]);
    const changes = listChanges(state.doc);
    expect(changes.map((c) => [c.kind, c.text])).toEqual([['insertion', 'ab'], ['insertion', 'c'], ['deletion', 'd']]);
    expect(new Set(changes.map((c) => c.id)).size).toBe(3);
    expect(changes[0].date).toBe('2024-01-01T00:00:00Z');
    // Settling the first change leaves the second, which has the same revision id.
    const next = apply(state, resolveChange(changes[0].id, 'reject'));
    expect(texts(next.doc)).toEqual(['-c', 'd']);
  });

  it('lists every kind of block change with its kind', () => {
    const state = stateOf([
      para([run('p')], { revision: { insertion: ADA, deletion: BO, formatChange: { id: '3', author: 'Cy', previous: {} } } }),
      para([run('q')]),
    ]);
    expect(listChanges(state.doc).map((c) => c.kind)).toEqual(['paragraphInserted', 'paragraphDeleted', 'paragraphFormat']);
  });

  it('ignores a revision attribute that is not an object, and nodes that cannot carry one', () => {
    const doc = docsSchema.nodes.doc.create(null, [docsSchema.nodes.paragraph.create({ revision: 'nope' }, [docsSchema.text('x')])]);
    expect(changeUnits(doc)).toEqual([]);
  });
});

describe('a pasted copy of tracked text is a change of its own', () => {
  it('has other ids than the source, and accepting or rejecting one leaves the other', () => {
    const state = stateOf([para([run('A '), run('kept', { insertion: ADA }), run(' B '), run('cut', { deletion: BO })])]);
    const plugin = createPlugins(docsSchema).find((p) => p.props.transformPasted)!;
    const copy = plugin.props.transformPasted!.call(plugin, state.doc.slice(3, 7), undefined as never, false);
    const copyDeleted = plugin.props.transformPasted!.call(plugin, state.doc.slice(10, 13), undefined as never, false);
    const end = state.doc.content.size - 1;
    let current = state.apply(state.tr.replaceRange(end, end, copy));
    current = current.apply(current.tr.replaceRange(current.doc.content.size - 1, current.doc.content.size - 1, copyDeleted));
    const changes = listChanges(current.doc);
    const insertions = changes.filter((c) => c.kind === 'insertion');
    const deletions = changes.filter((c) => c.kind === 'deletion');
    expect(insertions).toHaveLength(2);
    expect(deletions).toHaveLength(2);
    expect(new Set(insertions.map((c) => c.units[0].id)).size).toBe(2);
    expect(new Set(deletions.map((c) => c.units[0].id)).size).toBe(2);
    // Accept the pasted insertion: the original one stays pending, with its own id.
    const pastedInsertion = insertions[1];
    const accepted = apply(current, resolveChange(pastedInsertion.id, 'accept'));
    expect(listChanges(accepted.doc).filter((c) => c.kind === 'insertion').map((c) => c.units[0].id)).toEqual([insertions[0].units[0].id]);
    // Reject the pasted deletion: its text comes back as live text, the original deletion stays.
    const rejected = apply(accepted, resolveChange(listChanges(accepted.doc).filter((c) => c.kind === 'deletion')[1].id, 'reject'));
    expect(listChanges(rejected.doc).filter((c) => c.kind === 'deletion')).toHaveLength(1);
    expect(rejected.doc.textContent).toBe('A kept B cutkeptcut');
  });
});

describe('author colours of the list of changes', () => {
  it('match the page when the other end of a move has an author who is in no other change', () => {
    // Bo and Di get the same first colour; Bo only has the deleted end of a move, which the list gives to Ada.
    const blocks = [para([run('moved', { deletion: { id: '1', author: 'Bo', moveId: 'm' } }), run('moved', { insertion: { id: '2', author: 'Ada', moveId: 'm' } }), run('new', { insertion: { id: '3', author: 'Di' } })])];
    const doc = stateOf(blocks).doc;
    const changes = listChanges(doc);
    expect(changes.map((c) => c.author)).toEqual(['Ada', 'Di']);
    const page = assignAuthorColors(documentRevisionAuthors(doc));
    expect(revisionAuthors(changes, documentRevisionAuthors(doc)).map((a) => [a.name, a.color])).toEqual([
      ['Ada', page.get('Ada')],
      ['Di', page.get('Di')],
    ]);
    // Without the palette of the document, Di would get another colour than on the page.
    expect(revisionAuthors(changes).find((a) => a.name === 'Di')!.color).not.toBe(page.get('Di'));
  });
});
