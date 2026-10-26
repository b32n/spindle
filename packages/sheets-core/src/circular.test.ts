import { WorkbookImpl } from './workbook';
import type { CircularReferenceMode } from './types';

function wb(mode?: CircularReferenceMode): WorkbookImpl {
  const w = new WorkbookImpl('wb', 'WB');
  if (mode) w.setCircularReferenceMode(mode);
  return w;
}
const v = (w: WorkbookImpl, r: number, c: number) => w.getCellValue(undefined, r, c);

describe('circular references (default: Excel — 0 plus a warning event)', () => {
  it('a self-reference evaluates to 0, even when it would otherwise count up', () => {
    const w = wb();
    w.setFormula(undefined, 0, 0, '=A1+1');
    expect(v(w, 0, 0)).toBe(0);
  });

  it('every cell on a cycle is 0; a cell merely downstream of it is computed from that 0', () => {
    const w = wb();
    w.setFormula(undefined, 0, 0, '=B1+1'); // A1
    w.setFormula(undefined, 0, 1, '=A1+1'); // B1 closes the cycle
    w.setFormula(undefined, 0, 2, '=A1*2+5'); // C1 reads the cycle
    w.setFormula(undefined, 0, 3, '=C1+1'); // D1 reads C1
    expect([v(w, 0, 0), v(w, 0, 1)]).toEqual([0, 0]);
    expect(v(w, 0, 2)).toBe(5);
    expect(v(w, 0, 3)).toBe(6);
  });

  it('a three-cell cycle entered in any order settles to 0', () => {
    const w = wb();
    w.setFormula(undefined, 0, 1, '=C1'); // B1
    w.setFormula(undefined, 0, 2, '=A1'); // C1
    w.setFormula(undefined, 0, 0, '=B1'); // A1 -> cycle
    expect([v(w, 0, 0), v(w, 0, 1), v(w, 0, 2)]).toEqual([0, 0, 0]);
  });

  it('recovers when the cycle is broken', () => {
    const w = wb();
    w.setFormula(undefined, 0, 0, '=B1+1');
    w.setFormula(undefined, 0, 1, '=A1+1');
    w.setFormula(undefined, 0, 2, '=A1*2');
    w.setCellValue(undefined, 0, 1, 5); // B1 becomes a constant
    expect([v(w, 0, 0), v(w, 0, 1), v(w, 0, 2)]).toEqual([6, 5, 12]);
  });

  it('SUM over a range containing its own cell is circular (0), neighbours are unaffected', () => {
    const w = wb();
    w.setCellValue(undefined, 0, 0, 1);
    w.setFormula(undefined, 2, 0, '=SUM(A1:A5)');
    w.setFormula(undefined, 0, 2, '=A1+1');
    expect(v(w, 2, 0)).toBe(0);
    expect(v(w, 0, 2)).toBe(2);
  });

  it('is settled after a structural edit re-evaluates everything', () => {
    const w = wb();
    w.setFormula(undefined, 1, 0, '=A3'); // A2
    w.setFormula(undefined, 2, 0, '=A2+1'); // A3 closes the cycle
    w.getSheet().insertRows(0, 1); // shifts the cycle down; recalculates all
    expect([v(w, 2, 0), v(w, 3, 0)]).toEqual([0, 0]);
  });

  it("emits a 'circularReference' event naming the cells on the cycle", () => {
    const w = wb();
    const events: Array<{ row: number; col: number }[]> = [];
    w.on('circularReference', (e) => events.push((e.payload as { cells: { row: number; col: number }[] }).cells));
    w.setFormula(undefined, 0, 0, '=B1');
    expect(events).toHaveLength(0);
    w.setFormula(undefined, 0, 1, '=A1');
    expect(events.length).toBeGreaterThan(0);
    const cells = events[events.length - 1].map((c) => `${c.row}:${c.col}`).sort();
    expect(cells).toEqual(['0:0', '0:1']);
  });
});

describe("circular references with mode 'ref'", () => {
  it('cells on the cycle are #REF!, downstream cells see the error', () => {
    const w = wb('ref');
    w.setFormula(undefined, 0, 0, '=B1');
    w.setFormula(undefined, 0, 1, '=A1');
    w.setFormula(undefined, 0, 2, '=IFERROR(A1,"bad")');
    expect([v(w, 0, 0), v(w, 0, 1)]).toEqual(['#REF!', '#REF!']);
    expect(v(w, 0, 2)).toBe('bad');
  });

  it('switching the mode re-evaluates existing cycles', () => {
    const w = wb();
    w.setFormula(undefined, 0, 0, '=A1');
    expect(v(w, 0, 0)).toBe(0);
    w.setCircularReferenceMode('ref');
    expect(v(w, 0, 0)).toBe('#REF!');
    w.setCircularReferenceMode('zero');
    expect(v(w, 0, 0)).toBe(0);
  });
});

describe("circular references with mode 'error' (diagnostic token)", () => {
  it('cells on and downstream of the cycle show #CIRCULAR!', () => {
    const w = wb('error');
    w.setFormula(undefined, 0, 0, '=B1+1');
    w.setFormula(undefined, 0, 1, '=A1+1');
    expect([v(w, 0, 0), v(w, 0, 1)]).toEqual(['#CIRCULAR!', '#CIRCULAR!']);
  });
});
