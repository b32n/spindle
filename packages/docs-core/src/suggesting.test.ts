// Suggesting mode, at the level of transactions: what the user does, and what the document looks like afterwards, with the
// result of accepting and rejecting everything checked against doing the edit directly and not doing it at all.

import { EditorState, TextSelection, type Transaction } from 'prosemirror-state';
import { history, undo } from 'prosemirror-history';
import { deleteSelection, joinBackward, joinForward, splitBlock } from 'prosemirror-commands';
import type { Node as PmNode } from 'prosemirror-model';
import { Step, StepMap, StepResult } from 'prosemirror-transform';
import { docsSchema } from './prosemirror/schema';
import { blocksToPmDoc } from './prosemirror/sync';
import { changeUnits, listChanges, resolveAll } from './prosemirror/review';
import { suggestingPlugin, type SuggestingOptions } from './prosemirror/suggesting';
import type { Block, InlineContent } from './types';

const T0 = '2024-06-01T10:00:00Z';
let ids = 0;
interface Setup {
  author: string;
  active: boolean;
  notices: string[][];
}
const setup: Setup = { author: 'Ada', active: true, notices: [] };
const options: SuggestingOptions = {
  isActive: () => setup.active,
  author: () => setup.author,
  now: () => T0,
  newId: () => `n${++ids}`,
  onNotice: (n) => setup.notices.push(n),
};

beforeEach(() => {
  setup.author = 'Ada';
  setup.active = true;
  setup.notices = [];
  ids = 0;
});

const run = (text: string, extra: object = {}): InlineContent => ({ type: 'text', text, ...extra }) as InlineContent;
const para = (...content: InlineContent[]): Block => ({ id: 'p', type: 'paragraph', content }) as Block;
const stateOf = (blocks: Block[], at?: number, to?: number): EditorState => {
  const doc = blocksToPmDoc(blocks);
  const state = EditorState.create({ doc, plugins: [suggestingPlugin(options), history()] });
  return at === undefined ? state : state.apply(state.tr.setSelection(TextSelection.create(doc, at, to ?? at)));
};
const texts = (doc: PmNode) => {
  const out: string[] = [];
  doc.forEach((b) => out.push(b.textContent));
  return out;
};
/** The runs of the first paragraph as "text:marks" (only the tracked-change marks). */
const runs = (doc: PmNode, index = 0): string[] => {
  const out: string[] = [];
  doc.child(index).forEach((n) => {
    out.push(`${n.text ?? '?'}:${n.marks.map((m) => m.type.name).filter((x) => ['insertion', 'deletion', 'formatChange'].includes(x)).join('+')}`);
  });
  return out;
};
const apply = (state: EditorState, tr: Transaction) => state.apply(tr);
const type = (state: EditorState, text: string) => apply(state, state.tr.insertText(text));
const doCommand = (state: EditorState, command: (s: EditorState, d?: (tr: Transaction) => void) => boolean) => {
  let next = state;
  expect(command(state, (tr) => (next = state.apply(tr)))).toBe(true);
  return next;
};
const settleAll = (state: EditorState, action: 'accept' | 'reject') => doCommand(state, resolveAll(action));
const head = (state: EditorState) => state.selection.head;
/** Notices go out in a microtask: let it run. */
const flush = () => Promise.resolve();

describe('typing', () => {
  it('inserts text marked as inserted by the author, with the time, and leaves the caret after it', () => {
    const next = type(stateOf([para(run('abc'))], 3), 'XY');
    expect(texts(next.doc)).toEqual(['abXYc']);
    expect(runs(next.doc)).toEqual(['ab:', 'XY:insertion', 'c:']);
    const mark = next.doc.child(0).child(1).marks.find((m) => m.type.name === 'insertion')!;
    expect(mark.attrs).toMatchObject({ author: 'Ada', date: T0, revisionId: 'n1' });
    expect(head(next)).toBe(5);
  });

  it('continues one change while the same author types next to it, and starts another for another author', () => {
    let state = stateOf([para(run('abc'))], 3);
    state = type(state, 'X');
    state = type(state, 'Y');
    expect(listChanges(state.doc).map((c) => [c.author, c.text])).toEqual([['Ada', 'XY']]);
    setup.author = 'Bo';
    state = type(state, 'Z');
    expect(listChanges(state.doc).map((c) => [c.author, c.text])).toEqual([['Ada', 'XY'], ['Bo', 'Z']]);
  });

  it('accepting what was typed keeps it, rejecting takes it away', () => {
    const state = type(stateOf([para(run('abc'))], 3), 'X');
    expect(texts(settleAll(state, 'accept').doc)).toEqual(['abXc']);
    expect(texts(settleAll(state, 'reject').doc)).toEqual(['abc']);
  });

  it('keeps the formatting of typed text next to the insertion mark', () => {
    let state = stateOf([para(run('abc'))], 3);
    state = state.apply(state.tr.addStoredMark(docsSchema.marks.bold.create()));
    state = type(state, 'X');
    expect(state.doc.child(0).child(1).marks.map((m) => m.type.name).sort()).toEqual(['bold', 'insertion']);
  });

  it('typing over a selection keeps the old text, marked deleted, before the new text', () => {
    const next = type(stateOf([para(run('abcdef'))], 3, 5), 'XY');
    expect(texts(next.doc)).toEqual(['abcdXYef']);
    expect(runs(next.doc)).toEqual(['ab:', 'cd:deletion', 'XY:insertion', 'ef:']);
    expect(head(next)).toBe(7);
    expect(texts(settleAll(next, 'accept').doc)).toEqual(['abXYef']);
    expect(texts(settleAll(next, 'reject').doc)).toEqual(['abcdef']);
  });
});

describe('deleting text', () => {
  it('keeps backspaced text, marked deleted, with the caret before it, and one change for a run of backspaces', () => {
    let state = stateOf([para(run('abcdef'))], 5);
    state = apply(state, state.tr.delete(4, 5));
    expect(runs(state.doc)).toEqual(['abc:', 'd:deletion', 'ef:']);
    expect(head(state)).toBe(4);
    state = apply(state, state.tr.delete(3, 4));
    expect(runs(state.doc)).toEqual(['ab:', 'cd:deletion', 'ef:']);
    expect(listChanges(state.doc).map((c) => c.text)).toEqual(['cd']);
    expect(texts(settleAll(state, 'accept').doc)).toEqual(['abef']);
    expect(texts(settleAll(state, 'reject').doc)).toEqual(['abcdef']);
  });

  it('moves the caret past text deleted forward, and past text that is already deleted', () => {
    let state = stateOf([para(run('abcdef'))], 4);
    state = apply(state, state.tr.delete(4, 5));
    expect(runs(state.doc)).toEqual(['abc:', 'd:deletion', 'ef:']);
    expect(head(state)).toBe(5);
    // Deleting again over the deleted letter changes nothing in the text and moves on.
    const again = stateOf([para(run('abc'), run('d', { revision: { deletion: { id: 'x', author: 'Bo' } } }), run('ef'))], 4);
    const next = apply(again, again.tr.delete(4, 5));
    expect(runs(next.doc)).toEqual(['abc:', 'd:deletion', 'ef:']);
    expect(head(next)).toBe(5);
  });

  it('removes text the author inserted outright, with no trace', () => {
    let state = type(stateOf([para(run('abc'))], 3), 'XY');
    state = apply(state, state.tr.delete(4, 5));
    expect(texts(state.doc)).toEqual(['abXc']);
    expect(runs(state.doc)).toEqual(['ab:', 'X:insertion', 'c:']);
    state = apply(state, state.tr.delete(3, 4));
    expect(runs(state.doc)).toEqual(['abc:']);
    expect(changeUnits(state.doc)).toEqual([]);
  });

  it('marks text another author inserted as deleted by this author, keeping their insertion', () => {
    setup.author = 'Bo';
    const state = stateOf([para(run('a'), run('xyz', { revision: { insertion: { id: 'i', author: 'Ada' } } }), run('b'))], 3, 4);
    const next = apply(state, state.tr.deleteSelection());
    expect(runs(next.doc)).toEqual(['a:', 'x:insertion', 'y:insertion+deletion', 'z:insertion', 'b:']);
    expect(texts(settleAll(next, 'accept').doc)).toEqual(['axzb']);
  });

  it('deletes across several runs with their formatting, then accepting equals deleting', () => {
    const blocks = [para(run('aa', { bold: true }), run('bb'), run('cc', { italic: true }))];
    const state = stateOf(blocks, 2, 6);
    const next = apply(state, state.tr.deleteSelection());
    expect(texts(next.doc)).toEqual(['aabbcc']);
    expect(texts(settleAll(next, 'accept').doc)).toEqual(['ac']);
    expect(texts(settleAll(next, 'reject').doc)).toEqual(['aabbcc']);
  });
});

describe('paragraph marks', () => {
  const two = () => stateOf([para(run('one')), para(run('two'))]);

  it('Enter inserts a paragraph mark on the paragraph that ends there; rejecting joins the paragraphs again', () => {
    let state = stateOf([para(run('abcd'))], 3);
    state = doCommand(state, splitBlock);
    expect(texts(state.doc)).toEqual(['ab', 'cd']);
    expect(state.doc.child(0).attrs.revision).toMatchObject({ insertion: { author: 'Ada', id: 'n1' } });
    expect(state.doc.child(1).attrs.revision).toBeNull();
    expect(texts(settleAll(state, 'accept').doc)).toEqual(['ab', 'cd']);
    expect(texts(settleAll(state, 'reject').doc)).toEqual(['abcd']);
  });

  it('Enter and then Backspace leaves no trace', () => {
    let state = stateOf([para(run('abcd'))], 3);
    state = doCommand(state, splitBlock);
    state = doCommand(state, joinBackward);
    expect(texts(state.doc)).toEqual(['abcd']);
    expect(state.doc.childCount).toBe(1);
    expect(changeUnits(state.doc)).toEqual([]);
  });

  it('Backspace at the start of a paragraph keeps the break and marks it deleted; accepting joins, rejecting keeps it', () => {
    let state = two();
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 6)));
    state = doCommand(state, joinBackward);
    expect(texts(state.doc)).toEqual(['one', 'two']);
    expect(state.doc.child(0).attrs.revision).toMatchObject({ deletion: { author: 'Ada' } });
    expect(texts(settleAll(state, 'accept').doc)).toEqual(['onetwo']);
    const rejected = settleAll(state, 'reject');
    expect(texts(rejected.doc)).toEqual(['one', 'two']);
    expect(rejected.doc.child(0).attrs.revision).toBeNull();
  });

  it('Delete at the end of a paragraph does the same, and moves the caret on', () => {
    let state = two();
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 4)));
    state = doCommand(state, joinForward);
    expect(texts(state.doc)).toEqual(['one', 'two']);
    expect(state.doc.child(0).attrs.revision).toMatchObject({ deletion: { author: 'Ada' } });
    expect(texts(settleAll(state, 'accept').doc)).toEqual(['onetwo']);
  });

  it('deleting a selection across paragraphs marks the text and the breaks deleted', () => {
    const state = stateOf([para(run('one')), para(run('two')), para(run('three'))], 3, 13);
    const next = doCommand(state, deleteSelection);
    expect(texts(next.doc)).toEqual(['one', 'two', 'three']);
    expect(runs(next.doc, 0)).toEqual(['on:', 'e:deletion']);
    expect(next.doc.child(0).attrs.revision).toMatchObject({ deletion: { author: 'Ada' } });
    expect(next.doc.child(1).attrs.revision).toMatchObject({ deletion: { author: 'Ada' } });
    expect(texts(settleAll(next, 'accept').doc)).toEqual(['onree']);
    expect(texts(settleAll(next, 'reject').doc)).toEqual(['one', 'two', 'three']);
  });

  it('pasting several paragraphs marks the text and the breaks inserted', () => {
    const state = stateOf([para(run('abcd'))], 3);
    const pasted = docsSchema.nodeFromJSON({
      type: 'doc',
      content: [
        { type: 'paragraph', content: [{ type: 'text', text: 'X' }] },
        { type: 'paragraph', content: [{ type: 'text', text: 'Y' }] },
      ],
    });
    const next = apply(state, state.tr.replaceSelection(pasted.slice(1, pasted.content.size - 1)));
    expect(texts(next.doc)).toEqual(['abX', 'Ycd']);
    expect(next.doc.child(0).attrs.revision).toMatchObject({ insertion: { author: 'Ada' } });
    expect(texts(settleAll(next, 'reject').doc)).toEqual(['abcd']);
    expect(texts(settleAll(next, 'accept').doc)).toEqual(['abX', 'Ycd']);
  });
});

describe('formatting', () => {
  const bold = docsSchema.marks.bold;

  it('records what the formatting was, and rejecting puts it back', () => {
    const state = stateOf([para(run('abcdef'))]);
    const next = apply(state, state.tr.addMark(3, 5, bold.create()));
    expect(runs(next.doc)).toEqual(['ab:', 'cd:formatChange', 'ef:']);
    const mark = next.doc.child(0).child(1).marks.find((m) => m.type.name === 'formatChange')!;
    expect(mark.attrs).toMatchObject({ author: 'Ada', previous: {} });
    expect(next.doc.child(0).child(1).marks.some((m) => m.type === bold)).toBe(true);
    const rejected = settleAll(next, 'reject');
    expect(rejected.doc.child(0).childCount).toBe(1);
    expect(rejected.doc.child(0).child(0).marks).toEqual([]);
    expect(settleAll(next, 'accept').doc.child(0).child(1).marks.map((m) => m.type.name)).toEqual(['bold']);
  });

  it('records the formatting before a removal, and no change at all when it is put back', () => {
    const state = stateOf([para(run('abc', { bold: true }))]);
    const off = apply(state, state.tr.removeMark(1, 3, bold));
    expect(off.doc.child(0).child(0).marks.find((m) => m.type.name === 'formatChange')!.attrs.previous).toMatchObject({ bold: true });
    const back = apply(off, off.tr.addMark(1, 3, bold.create()));
    expect(changeUnits(back.doc)).toEqual([]);
  });

  it('keeps the first record when the same text is formatted again', () => {
    let state = stateOf([para(run('abc'))]);
    state = apply(state, state.tr.addMark(1, 4, bold.create()));
    state = apply(state, state.tr.addMark(1, 4, docsSchema.marks.italic.create()));
    const mark = state.doc.child(0).child(0).marks.find((m) => m.type.name === 'formatChange')!;
    expect(mark.attrs.previous).toEqual({});
    expect(listChanges(state.doc)).toHaveLength(1);
  });

  it('does not record formatting of text the author inserted themselves, nor the marks of comments and links', () => {
    let state = type(stateOf([para(run('abc'))], 3), 'XY');
    state = apply(state, state.tr.addMark(3, 5, bold.create()));
    expect(listChanges(state.doc).map((c) => c.kind)).toEqual(['insertion']);
    const plain = stateOf([para(run('abc'))]);
    const linked = apply(plain, plain.tr.addMark(1, 3, docsSchema.marks.link.create({ href: 'https://example.com/' })));
    expect(changeUnits(linked.doc)).toEqual([]);
  });

  it('records the old formatting of a paragraph whose alignment or type was changed, and keeps the first record', () => {
    let state = stateOf([para(run('title'))], 2);
    state = apply(state, state.tr.setNodeMarkup(0, undefined, { ...state.doc.child(0).attrs, alignment: 'center' }));
    expect(state.doc.child(0).attrs.revision).toMatchObject({ formatChange: { author: 'Ada', previous: {} } });
    state = apply(state, state.tr.setNodeMarkup(0, docsSchema.nodes.heading, { ...docsSchema.nodes.heading.create().attrs, level: 2, alignment: 'center', revision: state.doc.child(0).attrs.revision }));
    expect(state.doc.child(0).type.name).toBe('heading');
    expect(state.doc.child(0).attrs.revision.formatChange.previous).toEqual({});
    const rejected = settleAll(state, 'reject');
    expect(rejected.doc.child(0).type.name).toBe('paragraph');
    expect(rejected.doc.child(0).attrs.alignment).toBe('left');
    // No record when nothing the format records changed.
    const same = stateOf([para(run('x'))]);
    const none = apply(same, same.tr.setNodeMarkup(0, undefined, { ...same.doc.child(0).attrs }));
    expect(none.doc.child(0).attrs.revision).toBeNull();
  });

  it('records a paragraph attribute set one at a time', () => {
    const state = stateOf([para(run('x'))]);
    const next = apply(state, state.tr.setNodeAttribute(0, 'alignment', 'right'));
    expect(next.doc.child(0).attrs.revision).toMatchObject({ formatChange: { author: 'Ada' } });
    const ignored = apply(state, state.tr.setNodeAttribute(0, 'revision', { insertion: { id: 'a', author: 'Bo' } }));
    expect(ignored.doc.child(0).attrs.revision).toEqual({ insertion: { id: 'a', author: 'Bo' } });
  });
});

describe('what is not recorded', () => {
  it('is left as the user made it, and the host is told', async () => {
    const state = stateOf([para(run('abc')), para(run('def'))], 2);
    const table = docsSchema.nodeFromJSON({
      type: 'table',
      content: [{ type: 'table_row', content: [{ type: 'table_cell', content: [{ type: 'paragraph', content: [{ type: 'text', text: 'cell' }] }] }] }],
    });
    const next = apply(state, state.tr.insert(state.doc.child(0).nodeSize, table));
    expect(next.doc.childCount).toBe(3);
    await flush();
    expect(setup.notices.flat().join(' ')).toMatch(/structure that was inserted could not be suggested/);
    // A list made from a paragraph (wrapping) is applied directly too.
    setup.notices = [];
    const list = docsSchema.nodes.bullet_list;
    const wrapped = apply(state, state.tr.wrap(state.doc.resolve(1).blockRange()!, [{ type: list }, { type: docsSchema.nodes.list_item }]));
    expect(wrapped.doc.child(0).type.name).toBe('bullet_list');
    await flush();
    expect(setup.notices.flat().join(' ')).toMatch(/cannot be suggested and was made directly/);
  });

  it('deleting a table or a list is made directly, with a notice', async () => {
    const blocks: Block[] = [
      { id: 'l', type: 'list-item', listType: 'bullet', level: 0, content: [run('a')] } as Block,
      para(run('after')),
    ];
    const state = stateOf(blocks);
    const next = apply(state, state.tr.delete(0, state.doc.child(0).nodeSize));
    expect(texts(next.doc)).toEqual(['after']);
    await flush();
    expect(setup.notices.flat().join(' ')).toMatch(/was deleted directly/);
  });
});

/** A step the plugin does not know: it changes nothing. */
class UnknownStep extends Step {
  apply(doc: PmNode) {
    return StepResult.ok(doc);
  }
  getMap() {
    return StepMap.empty;
  }
  invert() {
    return this;
  }
  map() {
    return this;
  }
  toJSON() {
    return { stepType: 'unknown' };
  }
}

describe('when it does nothing', () => {
  it('tells the host about a step it does not know, and changes nothing', async () => {
    const state = stateOf([para(run('abc'))], 2);
    const next = apply(state, state.tr.step(new UnknownStep()));
    expect(next.doc.eq(state.doc)).toBe(true);
    await flush();
    expect(setup.notices.flat().join(' ')).toMatch(/A change that cannot be suggested was made directly/);
  });

  it('is off outside Suggesting mode', () => {
    setup.active = false;
    const next = type(stateOf([para(run('abc'))], 3), 'X');
    expect(runs(next.doc)).toEqual(['abXc:']);
  });

  it('leaves changes from collaborators, undo and redo, and its own alone', () => {
    const state = stateOf([para(run('abc'))], 3);
    for (const key of ['y-sync$', 'history$', 'y-undo$']) {
      const next = apply(state, state.tr.insertText('X').setMeta(key, true));
      expect(runs(next.doc)).toEqual(['abXc:']);
    }
    const hidden = apply(state, state.tr.insertText('X').setMeta('addToHistory', false));
    expect(runs(hidden.doc)).toEqual(['abXc:']);
  });

  it('ignores a transaction that does not change the document', () => {
    const state = stateOf([para(run('abc'))], 2);
    const next = apply(state, state.tr.setSelection(TextSelection.create(state.doc, 3)));
    expect(next.doc.eq(state.doc)).toBe(true);
  });

  it('takes the typing and what it recorded back with one undo', () => {
    const original = stateOf([para(run('abc'))], 3);
    const typed = type(original, 'X');
    expect(runs(typed.doc)).toEqual(['ab:', 'X:insertion', 'c:']);
    const undone = doCommand(typed, undo);
    expect(texts(undone.doc)).toEqual(['abc']);
    expect(changeUnits(undone.doc)).toEqual([]);
  });

  it('handles a transaction with several steps', () => {
    const state = stateOf([para(run('abcdef'))], 1);
    const tr = state.tr.insertText('X', 1).delete(5, 6);
    const next = apply(state, tr);
    expect(texts(next.doc)).toEqual(['Xabcdef']);
    expect(runs(next.doc)).toEqual(['X:insertion', 'abc:', 'd:deletion', 'ef:']);
    expect(texts(settleAll(next, 'accept').doc)).toEqual(['Xabcef']);
    expect(texts(settleAll(next, 'reject').doc)).toEqual(['abcdef']);
  });

  it('strips the tracked changes pasted text came with and makes it a suggestion of its own', () => {
    const pasted = docsSchema.nodeFromJSON({
      type: 'paragraph',
      content: [{ type: 'text', text: 'P', marks: [{ type: 'deletion', attrs: { revisionId: 'old', author: 'Cy', date: null } }] }],
    });
    const state = stateOf([para(run('abc'))], 3);
    const next = apply(state, state.tr.replaceSelectionWith(pasted.firstChild!, false));
    expect(runs(next.doc)).toEqual(['ab:', 'P:insertion', 'c:']);
  });
});

describe('several steps in one transaction', () => {
  const formatOf = (state: EditorState) => {
    const found: Array<Record<string, unknown>> = [];
    state.doc.descendants((n) => {
      const mark = n.marks.find((m) => m.type.name === 'formatChange');
      if (mark) found.push(mark.attrs.previous as Record<string, unknown>);
    });
    return found;
  };
  const bold = docsSchema.marks.bold;
  const italic = docsSchema.marks.italic;

  it('records the formatting before the first step, so rejecting brings back every mark that was removed (clear formatting)', () => {
    const state = stateOf([para(run('abcd', { bold: true, italic: true }))], 1, 5);
    const next = apply(state, state.tr.removeMark(1, 5, bold).removeMark(1, 5, italic));
    expect(next.doc.child(0).child(0).marks.some((m) => m.type.name === 'bold' || m.type.name === 'italic')).toBe(false);
    expect(formatOf(next)).toEqual([{ bold: true, italic: true }]);
    const rejected = settleAll(next, 'reject');
    const marks = rejected.doc.child(0).child(0).marks.map((m) => m.type.name);
    expect(marks).toEqual(expect.arrayContaining(['bold', 'italic']));
    expect(marks).not.toContain('formatChange');
  });

  it('records the formatting before the first step when marks are added one after the other, and none when they cancel out', () => {
    const state = stateOf([para(run('abcd'))], 1, 5);
    const next = apply(state, state.tr.addMark(1, 5, bold.create()).addMark(1, 5, italic.create()));
    expect(formatOf(next)).toEqual([{}]);
    const rejected = settleAll(next, 'reject');
    expect(rejected.doc.child(0).child(0).marks.map((m) => m.type.name)).toEqual([]);
    const cancelled = apply(state, state.tr.addMark(1, 5, bold.create()).removeMark(1, 5, bold));
    expect(formatOf(cancelled)).toEqual([]);
  });

  it('overlapping marks added in one transaction record the text each covers with what it was', () => {
    const state = stateOf([para(run('abcdef'))], 1, 7);
    const next = apply(state, state.tr.addMark(1, 5, bold.create()).addMark(3, 7, italic.create()));
    expect(texts(settleAll(next, 'reject').doc)).toEqual(['abcdef']);
    const rejected = settleAll(next, 'reject');
    const names: string[] = [];
    rejected.doc.child(0).forEach((n) => names.push(n.marks.map((m) => m.type.name).join('+') || '-'));
    expect(names).toEqual(['-']);
  });
});

describe('the selection', () => {
  it('stays a range after a format, so a second format reaches the same text', () => {
    const state = stateOf([para(run('abcd'))], 2, 4);
    const bold = apply(state, state.tr.addMark(2, 4, docsSchema.marks.bold.create()));
    expect([bold.selection.from, bold.selection.to]).toEqual([2, 4]);
    const italic = apply(bold, bold.tr.addMark(2, 4, docsSchema.marks.italic.create()));
    expect([italic.selection.from, italic.selection.to]).toEqual([2, 4]);
    expect(runs(italic.doc).join(' ')).toContain('formatChange');
  });

  it('stays a range after a paragraph is aligned', () => {
    const state = stateOf([para(run('abcd')), para(run('efgh'))], 2, 4);
    const aligned = apply(state, state.tr.setNodeMarkup(0, undefined, { ...state.doc.child(0).attrs, alignment: 'center' }));
    expect([aligned.selection.from, aligned.selection.to]).toEqual([2, 4]);
    expect(aligned.doc.child(0).attrs.revision).toMatchObject({ formatChange: { author: 'Ada' } });
  });
});

describe('what is inserted that cannot be recorded', () => {
  it('is not marked as inserted, as the notice says, and the paragraphs around it are not touched', async () => {
    const state = stateOf([para(run('abc')), para(run('def'))], 2);
    const table = docsSchema.nodeFromJSON({
      type: 'table',
      content: [{ type: 'table_row', content: [{ type: 'table_cell', content: [{ type: 'paragraph', content: [{ type: 'text', text: 'cell' }] }] }] }],
    });
    const next = apply(state, state.tr.insert(state.doc.child(0).nodeSize, table));
    expect(changeUnits(next.doc)).toEqual([]);
    await flush();
    expect(setup.notices.flat().join(' ')).toMatch(/inserted directly/);
  });
});

describe('notices', () => {
  it('go out after the transaction was applied, once', async () => {
    const state = stateOf([para(run('abc')), para(run('def'))], 2);
    const table = docsSchema.nodeFromJSON({ type: 'table', content: [{ type: 'table_row', content: [{ type: 'table_cell', content: [{ type: 'paragraph' }] }] }] });
    apply(state, state.tr.insert(state.doc.child(0).nodeSize, table));
    expect(setup.notices).toEqual([]);
    await flush();
    expect(setup.notices).toHaveLength(1);
  });
});

describe('paragraph marks in more places', () => {
  it('Backspace in an empty paragraph the author did not insert marks its mark deleted, and keeps the caret', () => {
    const state = stateOf([para(run('abc')), para(), para(run('def'))], 6);
    const next = doCommand(state, joinBackward);
    expect(next.doc.childCount).toBe(3);
    expect(listChanges(next.doc).map((c) => c.kind)).toEqual(['paragraphDeleted']);
    expect(texts(settleAll(next, 'accept').doc)).toEqual(['abc', 'def']);
    expect(texts(settleAll(next, 'reject').doc)).toEqual(['abc', '', 'def']);
  });

  it('splitting a paragraph that carries a change of its mark gives the new mark to the first half and keeps the change on the second', () => {
    const state = stateOf([{ id: 'p', type: 'paragraph', content: [run('abcd')], revision: { insertion: { id: 'x', author: 'Bo' } } } as Block], 3);
    const next = doCommand(state, splitBlock);
    expect(texts(next.doc)).toEqual(['ab', 'cd']);
    const by = listChanges(next.doc).map((c) => [c.kind, c.author]);
    expect(by).toContainEqual(['paragraphInserted', 'Ada']);
    expect(by).toContainEqual(['paragraphInserted', 'Bo']);
    expect(texts(settleAll(next, 'reject').doc)).toEqual(['abcd']);
  });
});

describe('line breaks', () => {
  const br = docsSchema.nodes.hard_break;

  it('a break typed is kept in tracked text, so it is part of the insertion and goes with it when rejected', () => {
    const state = stateOf([para(run('abcd'))], 3);
    const next = apply(state, state.tr.replaceSelectionWith(br.create()));
    expect(runs(next.doc)).toEqual(['ab:', '\n:insertion', 'cd:']);
    expect(listChanges(next.doc).map((c) => [c.kind, c.text])).toEqual([['insertion', '\n']]);
    expect(texts(settleAll(next, 'reject').doc)).toEqual(['abcd']);
    expect(texts(settleAll(next, 'accept').doc)).toEqual(['ab\ncd']);
  });

  it('a break deleted stays, marked deleted, and rejecting gets it back', () => {
    const state = stateOf([para(run('ab'), run('\n'), run('cd'))], 4);
    expect(state.doc.child(0).child(1).type).toBe(br);
    const next = apply(state, state.tr.delete(3, 4));
    expect(runs(next.doc)).toEqual(['ab:', '\n:deletion', 'cd:']);
    expect(texts(settleAll(next, 'accept').doc)).toEqual(['abcd']);
    expect(texts(settleAll(next, 'reject').doc)).toEqual(['ab\ncd']);
  });
});
