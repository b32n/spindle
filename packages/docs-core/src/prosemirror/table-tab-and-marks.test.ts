import { EditorState, TextSelection } from 'prosemirror-state';
import { docsSchema } from './schema';
import { goToNextCell, toggleMarkCommand } from './commands';
import { createPlugins } from './plugins';

const s = docsSchema;
const p = (text?: string) => s.node('paragraph', null, text ? [s.text(text)] : []);
const cell = (text?: string) => s.node('table_cell', null, [p(text)]);
const row = (...texts: string[]) => s.node('table_row', null, texts.map((t) => cell(t)));

function stateWith(doc: ReturnType<typeof s.node>, pos: number): EditorState {
  const state = EditorState.create({ doc, plugins: createPlugins(s) });
  return state.apply(state.tr.setSelection(TextSelection.create(state.doc, pos)));
}

/** Position just inside the text of the n-th textblock (0-based). */
function textPos(doc: ReturnType<typeof s.node>, n: number, offset = 0): number {
  let found = -1;
  let i = 0;
  doc.descendants((node, pos) => {
    if (node.isTextblock) {
      if (i === n) found = pos + 1 + offset;
      i++;
      return false;
    }
    return found < 0;
  });
  return found;
}

function run(state: EditorState, cmd: ReturnType<typeof goToNextCell>): EditorState {
  let next = state;
  cmd(state, (tr) => { next = state.apply(tr); });
  return next;
}

const cellTextAt = (state: EditorState) => state.selection.$from.parent.textContent;

describe('goToNextCell (Tab in tables)', () => {
  const table = s.node('table', null, [row('A1', 'B1', 'C1'), row('A2', 'B2', 'C2')]);
  const doc = s.node('doc', null, [p('intro'), table, p('outro')]);

  it('moves to the next cell, wrapping to the next row', () => {
    let st = stateWith(doc, textPos(doc, 1, 1)); // inside A1
    st = run(st, goToNextCell(1));
    expect(cellTextAt(st)).toBe('B1');
    st = run(st, goToNextCell(1));
    expect(cellTextAt(st)).toBe('C1');
    st = run(st, goToNextCell(1));
    expect(cellTextAt(st)).toBe('A2');
  });

  it('Shift-Tab moves back, and stays in the first cell', () => {
    let st = stateWith(doc, textPos(doc, 2)); // B1
    st = run(st, goToNextCell(-1));
    expect(cellTextAt(st)).toBe('A1');
    const same = run(st, goToNextCell(-1));
    expect(cellTextAt(same)).toBe('A1');
  });

  it('Tab in the last cell appends a row and moves into it', () => {
    let st = stateWith(doc, textPos(doc, 6)); // C2
    expect(cellTextAt(st)).toBe('C2');
    st = run(st, goToNextCell(1));
    const t = st.doc.child(1);
    expect(t.childCount).toBe(3);
    expect(t.lastChild!.childCount).toBe(3);
    expect(cellTextAt(st)).toBe('');
    expect(st.selection.$from.node(st.selection.$from.depth - 2)).toBe(t.lastChild);
    expect(st.doc.lastChild!.textContent).toBe('outro');
  });

  it('does nothing outside a table', () => {
    const st = stateWith(doc, textPos(doc, 0, 2));
    expect(goToNextCell(1)(st, () => undefined)).toBe(false);
  });

});

describe('toggleMarkCommand on a partly-marked selection (#71)', () => {
  const em = s.marks.italic;
  const mixed = s.node('doc', null, [s.node('paragraph', null, [s.text('aa '), s.text('bb', [em.create()])])]);

  function select(doc: typeof mixed, from: number, to: number): EditorState {
    const st = EditorState.create({ doc });
    return st.apply(st.tr.setSelection(TextSelection.create(st.doc, from, to)));
  }

  it('adds the mark to the whole selection when only part has it', () => {
    let st = select(mixed, 1, 6);
    toggleMarkCommand(em)(st, (tr) => { st = st.apply(tr); });
    expect(st.doc.rangeHasMark(1, 6, em)).toBe(true);
    let all = true;
    st.doc.nodesBetween(1, 6, (n) => { if (n.isText && !em.isInSet(n.marks)) all = false; });
    expect(all).toBe(true);
  });

  it('removes the mark when the whole selection has it', () => {
    let st = select(mixed, 4, 6);
    toggleMarkCommand(em)(st, (tr) => { st = st.apply(tr); });
    expect(st.doc.rangeHasMark(4, 6, em)).toBe(false);
  });
});
