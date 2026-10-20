import { AllSelection, EditorState, TextSelection } from 'prosemirror-state';
import type { Node as PmNode } from 'prosemirror-model';
import { docsSchema } from './schema';
import { toggleBulletList, toggleOrderedList } from './commands';

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
