import { AllSelection, EditorState, TextSelection } from 'prosemirror-state';
import type { Node as PmNode } from 'prosemirror-model';
import { docsSchema } from './schema';
import { insertPageBreak, insertSectionBreak, removeSectionBreak, toggleBulletList, toggleOrderedList } from './commands';

const { nodes } = docsSchema;
const p = (text: string) => nodes.paragraph.create(null, docsSchema.text(text));
const item = (text: string) => nodes.list_item.create(null, p(text));

function stateWith(doc: PmNode, selectAll = false): EditorState {
  const base = EditorState.create({ doc, schema: docsSchema });
  const sel = selectAll
    ? TextSelection.between(TextSelection.atStart(doc).$from, TextSelection.atEnd(doc).$to)
    : TextSelection.atEnd(doc);
  return base.apply(base.tr.setSelection(sel));
}

function run(state: EditorState, cmd: ReturnType<typeof toggleBulletList>): EditorState {
  let next = state;
  const ok = cmd(state, (tr) => {
    next = state.apply(tr);
  });
  expect(ok).toBe(true);
  return next;
}

const topTypes = (s: EditorState) => s.doc.children.map((c) => c.type.name);

describe('list toggle commands', () => {
  it('wraps a paragraph in a bullet list', () => {
    const s = run(stateWith(nodes.doc.create(null, p('one'))), toggleBulletList(docsSchema));
    expect(topTypes(s)).toEqual(['bullet_list']);
  });

  it('toggles a bullet list off again', () => {
    const doc = nodes.doc.create(null, nodes.bullet_list.create(null, [item('one'), item('two')]));
    const s = run(stateWith(doc, true), toggleBulletList(docsSchema));
    expect(topTypes(s)).toEqual(['paragraph', 'paragraph']);
  });

  it('converts a bullet list to a numbered list (regression: was a no-op)', () => {
    const doc = nodes.doc.create(null, nodes.bullet_list.create(null, [item('one'), item('two')]));
    const s = run(stateWith(doc, true), toggleOrderedList(docsSchema));
    expect(topTypes(s)).toEqual(['ordered_list']);
    expect(s.doc.firstChild!.childCount).toBe(2);
  });

  it('converts a list under a select-all selection (regression: no-op)', () => {
    const doc = nodes.doc.create(null, nodes.bullet_list.create(null, [item('one'), item('two')]));
    const base = EditorState.create({ doc, schema: docsSchema });
    const state = base.apply(base.tr.setSelection(new AllSelection(doc)));
    const s = run(state, toggleOrderedList(docsSchema));
    expect(topTypes(s)).toEqual(['ordered_list']);
  });

  it('converts a numbered list to a bullet list', () => {
    const doc = nodes.doc.create(null, nodes.ordered_list.create(null, [item('one'), item('two')]));
    const s = run(stateWith(doc), toggleBulletList(docsSchema));
    expect(topTypes(s)).toEqual(['bullet_list']);
  });
});

describe('insertPageBreak', () => {
  it('leaves the caret in a paragraph after the break so typing does not replace it', () => {
    
    const doc = nodes.doc.create(null, p('first'));
    const state = stateWith(doc);
    let next = state;
    expect(insertPageBreak(docsSchema)(state, (tr) => { next = state.apply(tr); })).toBe(true);
    expect(topTypes(next)).toEqual(['paragraph', 'page_break', 'paragraph']);
    const typed = next.apply(next.tr.insertText('second'));
    expect(topTypes(typed)).toEqual(['paragraph', 'page_break', 'paragraph']);
    expect(typed.doc.lastChild!.textContent).toBe('second');
  });
});

describe('insertSectionBreak', () => {
  const at = (doc: PmNode, text: string, offset: number): EditorState => {
    let pos = -1;
    doc.descendants((n, p0) => {
      if (n.isText && n.text!.includes(text) && pos < 0) pos = p0 + n.text!.indexOf(text) + offset;
    });
    const base = EditorState.create({ doc, schema: docsSchema });
    return base.apply(base.tr.setSelection(TextSelection.create(doc, pos)));
  };

  it('splits the paragraph at the caret, names the first part and leaves the caret after the break', () => {
    const state = at(nodes.doc.create(null, [p('alphabeta'), p('tail')]), 'alphabeta', 5);
    const next = run(state, insertSectionBreak(docsSchema, 'sec_new'));
    expect(topTypes(next)).toEqual(['paragraph', 'section_break', 'paragraph', 'paragraph']);
    expect(next.doc.child(0).textContent).toBe('alpha');
    expect(next.doc.child(1).attrs.sectionId).toBe('sec_new');
    expect(next.doc.child(2).textContent).toBe('beta');
    expect(next.selection.$from.parent.textContent).toBe('beta');
    const typed = next.apply(next.tr.insertText('X'));
    expect(typed.doc.child(1).type.name).toBe('section_break');
  });

  it('adds a paragraph when the break would be last, and replaces a selection', () => {
    const doc = nodes.doc.create(null, p('only'));
    const next = run(stateWith(doc), insertSectionBreak(docsSchema, 's'));
    expect(topTypes(next)).toEqual(['paragraph', 'section_break', 'paragraph']);
    expect(next.doc.lastChild!.textContent).toBe('');
    const selected = stateWith(nodes.doc.create(null, p('hello world')), true);
    const out = run(selected, insertSectionBreak(docsSchema, 's'));
    expect(out.doc.textContent).toBe('');
    expect(out.doc.children.some((c) => c.type.name === 'section_break')).toBe(true);
  });

  it('puts the break after a table or list the caret is in', () => {
    const cell = nodes.table_cell.create(null, p('cell'));
    const table = nodes.table.create(null, nodes.table_row.create(null, cell));
    const doc = nodes.doc.create(null, [table, p('after')]);
    const next = run(at(doc, 'cell', 1), insertSectionBreak(docsSchema, 't'));
    expect(topTypes(next)).toEqual(['table', 'section_break', 'paragraph']);
    const listDoc = nodes.doc.create(null, [nodes.bullet_list.create(null, item('one')), p('after')]);
    expect(topTypes(run(at(listDoc, 'one', 1), insertSectionBreak(docsSchema, 'l')))).toEqual(['bullet_list', 'section_break', 'paragraph']);
  });

  it('does nothing without an id', () => {
    expect(insertSectionBreak(docsSchema, '')(stateWith(nodes.doc.create(null, p('x'))), () => {})).toBe(false);
  });
});

describe('removeSectionBreak', () => {
  const brk = (id: string) => nodes.section_break.create({ sectionId: id });
  const doc = nodes.doc.create(null, [p('one'), brk('a'), p('two'), brk('b'), p('three')]);
  const caretIn = (text: string): EditorState => {
    let pos = 0;
    doc.descendants((n, p0) => {
      if (n.isText && n.text === text) pos = p0 + 1;
    });
    const base = EditorState.create({ doc, schema: docsSchema });
    return base.apply(base.tr.setSelection(TextSelection.create(doc, pos)));
  };
  const ids = (s: EditorState) => s.doc.children.filter((c) => c.type.name === 'section_break').map((c) => c.attrs.sectionId);

  it('removes the break that ends the caret’s section', () => {
    expect(ids(run(caretIn('one'), removeSectionBreak()))).toEqual(['b']);
    expect(ids(run(caretIn('two'), removeSectionBreak()))).toEqual(['a']);
  });

  it('in the last section removes the break that starts it', () => {
    expect(ids(run(caretIn('three'), removeSectionBreak()))).toEqual(['a']);
  });

  it('is not available without a break', () => {
    expect(removeSectionBreak()(stateWith(nodes.doc.create(null, p('x'))), () => {})).toBe(false);
  });
});
