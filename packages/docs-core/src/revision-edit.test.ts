// Pasting tracked text, typing inside a deleted run and line breaks inside tracked text.

import { EditorState, TextSelection } from 'prosemirror-state';
import { Fragment, Slice } from 'prosemirror-model';
import { docsSchema } from './prosemirror/schema';
import { createPlugins } from './prosemirror/plugins';
import { freshRevisionIds, insertTrackedLineBreak, newRevisionId, pasteOutsideDeletion, typeOutsideDeletion } from './prosemirror/revision-edit';
import { collectRevisions } from './revisions';

const { insertion, deletion } = docsSchema.marks;
const ADA = { revisionId: '1', author: 'Ada', date: '2024-01-01T00:00:00Z' };
const BO = { revisionId: '2', author: 'Bo', date: null };

const text = (s: string, ...marks: ReturnType<typeof insertion.create>[]) => docsSchema.text(s, marks);
const paragraph = (...children: ReturnType<typeof text>[]) => docsSchema.nodes.paragraph.create(null, children);
const docOf = (...children: ReturnType<typeof text>[]) => docsSchema.node('doc', null, [paragraph(...children)]);
const stateOf = (doc = docOf(text('keep '), text('gone away', deletion.create(BO)), text(' new', insertion.create(ADA)))) =>
  EditorState.create({ doc, plugins: createPlugins(docsSchema) });

const markIds = (doc: ReturnType<typeof docOf>, kind: 'insertion' | 'deletion') => {
  const out: string[] = [];
  doc.descendants((n) => {
    for (const m of n.marks) if (m.type.name === kind) out.push(String(m.attrs.revisionId));
  });
  return out;
};

describe('pasting tracked HTML', () => {
  // The text style rule matches every `span`; a rule of the same priority listed after it never sees the pasted change
  // (the marks of copied tracked text were lost in the browser before the rules had a priority of their own).
  const rulesOf = (name: string) => docsSchema.marks[name].spec.parseDOM ?? [];
  const textStyleRule = rulesOf('textStyle').find((rule) => rule.tag === 'span')!;

  it.each(['insertion', 'deletion'])('parses a %s span before the plain span rule of the text style', (name) => {
    expect(textStyleRule).toBeDefined();
    const [rule] = rulesOf(name);
    expect(rule.priority ?? 50).toBeGreaterThan(textStyleRule.priority ?? 50);
  });
});

describe('freshRevisionIds', () => {
  const sliceOf = (...children: ReturnType<typeof text>[]) => new Slice(Fragment.from(paragraph(...children)), 1, 1);

  it('gives each distinct change a new id, keeps runs of one change together, and keeps author and date', () => {
    const slice = sliceOf(text('a', insertion.create(ADA)), text('b', insertion.create(ADA), docsSchema.marks.bold.create()), text('c', insertion.create({ ...ADA, revisionId: '5' })), text('d', deletion.create(BO)));
    let n = 0;
    const out = freshRevisionIds(slice, () => `f${++n}`);
    const kids: ReturnType<typeof text>[] = [];
    out.content.firstChild!.forEach((c) => kids.push(c as never));
    const ids = kids.map((c) => c.marks.find((m) => m.type.name === 'insertion' || m.type.name === 'deletion')!.attrs);
    expect(ids.map((a) => a.revisionId)).toEqual(['f1', 'f1', 'f2', 'f3']);
    expect(ids[0]).toMatchObject({ author: 'Ada', date: '2024-01-01T00:00:00Z' });
    expect(ids[3]).toMatchObject({ author: 'Bo', date: null });
    // Other marks stay, the source slice is untouched, the open depth is kept.
    expect(kids[1].marks.map((m) => m.type.name).sort()).toEqual(['bold', 'insertion']);
    expect(slice.content.firstChild!.firstChild!.marks[0].attrs.revisionId).toBe('1');
    expect(out.openStart).toBe(1);
  });

  it('treats an insertion and a deletion that share an id as two changes, and leaves plain text alone', () => {
    const plain = text('plain');
    const slice = sliceOf(plain, text('x', insertion.create({ ...ADA, revisionId: '9' })), text('y', deletion.create({ ...BO, revisionId: '9' })));
    const out = freshRevisionIds(slice);
    expect(out.content.firstChild!.firstChild).toBe(plain);
    const [i, d] = [1, 2].map((k) => out.content.firstChild!.child(k).marks[0].attrs.revisionId);
    expect(i).not.toBe(d);
    expect(i).not.toBe('9');
  });

  it('reaches tracked text nested in lists and tables', () => {
    const li = docsSchema.nodes.list_item.create(null, paragraph(text('n', insertion.create(ADA))));
    const list = docsSchema.nodes.bullet_list.create(null, li);
    const out = freshRevisionIds(new Slice(Fragment.from(list), 0, 0), () => 'fresh');
    expect(markIds(docsSchema.node('doc', null, [out.content.firstChild!]) as never, 'insertion')).toEqual(['fresh']);
  });

  it('makes ids that differ from call to call', () => {
    expect(newRevisionId()).not.toBe(newRevisionId());
  });

  it('gives pasted tracked text change ids of its own (distinct from the copied change)', () => {
    const state = stateOf(docOf(text('A '), text('kept', insertion.create(ADA)), text(' B ')));
    // Copy the inserted run, paste it at the end through the plugin's transformPasted.
    const run = state.doc.slice(3, 7);
    const plugin = state.plugins.find((p) => p.props.transformPasted)!;
    const pasted = plugin.props.transformPasted!.call(plugin, run, undefined as never, false);
    const end = state.doc.content.size - 1;
    const next = state.apply(state.tr.replaceRange(end, end, pasted));
    const ids = markIds(next.doc, 'insertion');
    expect(ids).toHaveLength(2);
    expect(new Set(ids).size).toBe(2);
    const changes = collectRevisions(next.doc);
    expect(changes.map((c) => c.id).sort()).toEqual([...ids].sort());
    expect(changes.every((c) => c.author === 'Ada')).toBe(true);
  });
});

describe('typing inside a deletion', () => {
  const at = (state: EditorState, pos: number) => state.apply(state.tr.setSelection(TextSelection.create(state.doc, pos)));

  it('is live text: the text typed in the middle of a deletion does not carry it', () => {
    const state = at(stateOf(), 1 + 5 + 3); // "gon|e away"
    const tr = typeOutsideDeletion(state, state.selection.from, state.selection.to, 'XY')!;
    expect(tr).not.toBeNull();
    const next = state.apply(tr);
    const kids: { text: string; marks: string[] }[] = [];
    next.doc.firstChild!.forEach((c) => kids.push({ text: c.text!, marks: c.marks.map((m) => m.type.name) }));
    expect(kids).toEqual([
      { text: 'keep ', marks: [] },
      { text: 'gon', marks: ['deletion'] },
      { text: 'XY', marks: [] },
      { text: 'e away', marks: ['deletion'] },
      { text: ' new', marks: ['insertion'] },
    ]);
    // The next typed character does not pick the deletion up again.
    expect(next.storedMarks?.some((m) => m.type.name === 'deletion') ?? false).toBe(false);
  });

  it('is tagged as ordinary input, like the default text input', () => {
    const state = at(stateOf(), 1 + 5 + 3);
    const tr = typeOutsideDeletion(state, state.selection.from, state.selection.to, 'Q')!;
    expect(tr.getMeta('uiEvent')).toBe('input');
  });

  it('replaces a selection inside a deletion with live text', () => {
    const state = stateOf();
    const next = state.apply(typeOutsideDeletion(state, 7, 9, 'Z')!);
    expect(markIds(next.doc, 'deletion')).toHaveLength(2);
    expect(next.doc.textBetween(0, next.doc.content.size, '|', '|')).toContain('Z');
  });

  it('leaves the edges and plain text to the default insertion', () => {
    const state = stateOf();
    expect(typeOutsideDeletion(state, 1, 1, 'x')).toBeNull();
    expect(typeOutsideDeletion(state, 6, 6, 'x')).toBeNull(); // before the deletion
    expect(typeOutsideDeletion(state, 15, 15, 'x')).toBeNull(); // after the deletion, before the insertion
  });

  it('follows the stored marks over the marks around the caret', () => {
    const state = stateOf();
    // Stored marks carrying a deletion at a plain spot: the text is still live, and the stored marks lose the deletion.
    const stored = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 3)).setStoredMarks([deletion.create(BO), docsSchema.marks.bold.create()]));
    const tr = typeOutsideDeletion(stored, 3, 3, 'q')!;
    expect(tr).not.toBeNull();
    const next = stored.apply(tr);
    expect(next.doc.nodeAt(3)!.marks.map((m) => m.type.name)).not.toContain('deletion');
    expect(next.storedMarks!.map((m) => m.type.name)).toEqual(['bold']);
    // Stored marks without it inside a deletion: nothing to do, the default insertion is already live text.
    const plain = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 9)).setStoredMarks([]));
    expect(typeOutsideDeletion(plain, 9, 9, 'q')).toBeNull();
  });

  it('pastes plain text inside a deletion as live text, and leaves other spots alone', () => {
    const state = at(stateOf(), 1 + 5 + 3);
    const plain = new Slice(Fragment.from(text('PQ', deletion.create(BO))), 0, 0); // what the clipboard parser makes of plain text: the marks around the caret
    const next = state.apply(pasteOutsideDeletion(state, plain)!);
    const kids: { text: string; marks: string[] }[] = [];
    next.doc.firstChild!.forEach((c) => kids.push({ text: c.text!, marks: c.marks.map((m) => m.type.name) }));
    expect(kids.slice(0, 4)).toEqual([
      { text: 'keep ', marks: [] },
      { text: 'gon', marks: ['deletion'] },
      { text: 'PQ', marks: [] },
      { text: 'e away', marks: ['deletion'] },
    ]);
    expect(next.storedMarks?.some((m) => m.type.name === 'deletion') ?? false).toBe(false);
    const outside = at(stateOf(), 1);
    expect(pasteOutsideDeletion(outside, plain)).toBeNull();
  });

  it('is installed as a paste handler that leaves clipboard HTML alone', () => {
    const state = at(stateOf(), 9);
    const plugin = state.plugins.find((p) => p.props.handlePaste)!;
    let dispatched = 0;
    const view = { state, dispatch: () => dispatched++ } as never;
    const slice = new Slice(Fragment.from(text('PQ')), 0, 0);
    const clip = (html: string) => ({ clipboardData: { getData: (type: string) => (type === 'text/html' ? html : '') } }) as unknown as ClipboardEvent;
    expect(plugin.props.handlePaste!.call(plugin, view, clip('<b>x</b>'), slice)).toBe(false);
    expect(plugin.props.handlePaste!.call(plugin, view, clip(''), slice)).toBe(true);
    expect(dispatched).toBe(1);
  });

  it('keeps text typed inside an insertion as part of it', () => {
    const state = stateOf();
    expect(typeOutsideDeletion(state, 17, 17, 'x')).toBeNull();
    const typed = state.apply(state.tr.insertText('x', 17));
    expect(typed.doc.nodeAt(17)!.marks.map((m) => m.type.name)).toContain('insertion');
  });

  it('is installed in the editor plugins (handleTextInput)', () => {
    const state = at(stateOf(), 9);
    const plugin = state.plugins.find((p) => p.props.transformPasted)!;
    let dispatched: ReturnType<typeof state.tr.setMeta> | undefined;
    const view = { state, dispatch: (tr: typeof dispatched) => (dispatched = tr) } as never;
    expect(plugin.props.handleTextInput!.call(plugin, view, 9, 9, 'q', () => ({}) as never)).toBe(true);
    expect(dispatched).toBeDefined();
    const plain = at(stateOf(docOf(text('plain'))), 2);
    const view2 = { state: plain, dispatch: () => undefined } as never;
    expect(plugin.props.handleTextInput!.call(plugin, view2, 2, 2, 'q', () => ({}) as never)).toBe(false);
  });
});

describe('a line break typed inside tracked text', () => {
  const at = (state: EditorState, pos: number) => state.apply(state.tr.setSelection(TextSelection.create(state.doc, pos)));

  it('is a newline in the tracked text, not a hard_break node', () => {
    const state = at(stateOf(docOf(text('ab', insertion.create(ADA)))), 2);
    let tr;
    expect(insertTrackedLineBreak(state, (t) => (tr = t))).toBe(true);
    const next = state.apply(tr!);
    expect(next.doc.textContent).toBe('a\nb');
    let hardBreaks = 0;
    next.doc.descendants((n) => {
      if (n.type.name === 'hard_break') hardBreaks++;
    });
    expect(hardBreaks).toBe(0);
    expect(markIds(next.doc as never, 'insertion')).toEqual(['1']); // one text node: the break is inside the run
  });

  it('inside a deletion the break is live; elsewhere the command passes', () => {
    const state = at(stateOf(), 9);
    let tr;
    expect(insertTrackedLineBreak(state, (t) => (tr = t))).toBe(true);
    const next = state.apply(tr!);
    let live = true;
    next.doc.descendants((n) => {
      if (n.isText && n.text!.includes('\n') && n.marks.length) live = false;
    });
    expect(live).toBe(true);
    expect(insertTrackedLineBreak(at(stateOf(docOf(text('plain'))), 2))).toBe(false);
    // A selection is replaced by the default break, not by this command.
    const tracked = stateOf(docOf(text('ab', insertion.create(ADA))));
    expect(insertTrackedLineBreak(tracked.apply(tracked.tr.setSelection(TextSelection.create(tracked.doc, 1, 2))))).toBe(false);
  });
});
