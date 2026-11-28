import { InMemoryProvider, __resetInMemoryRooms, type CollabChannel, type CollabIdentity } from '@b32nio/spindle-shared';
import { WorkbookImpl } from './workbook';
import type { Cell } from './types';

// S4b: insert/delete of rows and columns rewrite the references in the text of cached-only formulas (see formula-parser/ref-shift.ts).

const identity = (name: string): CollabIdentity => ({ userId: name, displayName: name, color: '#123456' });
const LET = '=_xlfn.LET(_xlpm.x,A1,_xlpm.y,B2:C3,_xlpm.x+_xlpm.y)';

/** A workbook with a cached-only cell at (r, c) of the sheet. */
function cachedOnly(wb: WorkbookImpl, sheetId: string | undefined, r: number, c: number, formula: string, value: unknown = 99): void {
  wb.setCell(sheetId, r, c, { formula, formulaAst: undefined, value: value as Cell['value'], cachedOnly: true });
}

/** A workbook whose first sheet is named Main. */
function mainWorkbook(id = 'wb'): WorkbookImpl {
  const wb = new WorkbookImpl(id, 'WB');
  wb.renameSheet(wb.getSheet().id, 'Main');
  return wb;
}

const cellAt = (wb: WorkbookImpl, r: number, c: number, sheetId?: string): Cell | undefined => wb.getCell(sheetId, r, c);

describe('a cached-only formula follows structural edits', () => {
  it('rewrites its references on insert and delete of rows and columns, and keeps its saved value and mark', () => {
    const wb = mainWorkbook();
    cachedOnly(wb, undefined, 9, 5, LET);
    const sheet = wb.getSheet();
    sheet.insertRows(0, 2);
    expect(cellAt(wb, 11, 5)).toMatchObject({ formula: '=_xlfn.LET(_xlpm.x,A3,_xlpm.y,B4:C5,_xlpm.x+_xlpm.y)', value: 99, cachedOnly: true });
    sheet.insertCols(1, 1);
    expect(cellAt(wb, 11, 6)?.formula).toBe('=_xlfn.LET(_xlpm.x,A3,_xlpm.y,C4:D5,_xlpm.x+_xlpm.y)');
    sheet.deleteRows(0, 1);
    expect(cellAt(wb, 10, 6)?.formula).toBe('=_xlfn.LET(_xlpm.x,A2,_xlpm.y,C3:D4,_xlpm.x+_xlpm.y)');
    sheet.deleteCols(2, 1); // C is a corner of the range: the range keeps D, now C
    expect(cellAt(wb, 10, 5)?.formula).toBe('=_xlfn.LET(_xlpm.x,A2,_xlpm.y,C3:C4,_xlpm.x+_xlpm.y)');
    sheet.deleteCols(2, 1); // its last column goes: the range is gone
    expect(cellAt(wb, 10, 4)?.formula).toBe('=_xlfn.LET(_xlpm.x,A2,_xlpm.y,#REF!,_xlpm.x+_xlpm.y)');
    expect(cellAt(wb, 10, 4)).toMatchObject({ value: 99, cachedOnly: true });
  });

  it('shifts the references of every sheet that point at the edited one, and only those', () => {
    const wb = mainWorkbook();
    const other = wb.addSheet('Other');
    cachedOnly(wb, undefined, 0, 0, '=_xlfn.F(A5,Other!A5,\'Other\'!B1:B9)');
    cachedOnly(wb, other.id, 0, 0, '=_xlfn.F(A5,Main!A5)');
    wb.getSheet().insertRows(0, 1);
    expect(cellAt(wb, 1, 0)?.formula).toBe("=_xlfn.F(A6,Other!A5,'Other'!B1:B9)");
    expect(cellAt(wb, 0, 0, other.id)?.formula).toBe('=_xlfn.F(A5,Main!A6)');
    other.insertRows(0, 1);
    expect(cellAt(wb, 1, 0)?.formula).toBe("=_xlfn.F(A6,Other!A6,'Other'!B2:B10)");
    expect(cellAt(wb, 1, 0, other.id)?.formula).toBe('=_xlfn.F(A6,Main!A6)');
  });

  it('leaves alone a cell in the deleted rows (it is gone), a cached-only cell without text, and a text without references', () => {
    const wb = mainWorkbook();
    cachedOnly(wb, undefined, 0, 0, '=_xlfn.F(A5)');
    cachedOnly(wb, undefined, 1, 0, '=_xlfn.F(1)');
    wb.setCell(undefined, 2, 0, { value: 5, cachedOnly: true });
    wb.getSheet().deleteRows(0, 1);
    expect(cellAt(wb, 0, 0)).toMatchObject({ formula: '=_xlfn.F(1)' });
    expect(cellAt(wb, 1, 0)).toMatchObject({ value: 5, cachedOnly: true });
    expect(cellAt(wb, 1, 0)?.formula).toBeUndefined();
  });

  it('does nothing for an edit of zero rows, a change of layout, or a sort', () => {
    const wb = mainWorkbook();
    cachedOnly(wb, undefined, 0, 2, '=_xlfn.F(A5,B1:B3)');
    wb.setCellValue(undefined, 2, 0, 3);
    wb.setCellValue(undefined, 1, 0, 9);
    const sheet = wb.getSheet();
    sheet.insertRows(0, 0);
    sheet.setColWidth(0, 150);
    sheet.setRowHeight(0, 30);
    expect(cellAt(wb, 0, 2)?.formula).toBe('=_xlfn.F(A5,B1:B3)');
    sheet.setSortOrder([{ column: 0, direction: 'asc' }] as never);
    wb.sortSheet();
    expect(cellAt(wb, 0, 2)?.formula).toBe('=_xlfn.F(A5,B1:B3)');
  });

  it('does not touch an ordinary formula (its stable references move it) or recalculate the saved value', () => {
    const wb = mainWorkbook();
    wb.setCellValue(undefined, 0, 0, 4);
    wb.setFormula(undefined, 0, 1, '=A1*2');
    cachedOnly(wb, undefined, 0, 2, '=_xlfn.F(A1)', 'saved');
    wb.getSheet().insertRows(0, 1);
    expect(cellAt(wb, 1, 1)).toMatchObject({ formula: '=A2*2', value: 8 });
    expect(cellAt(wb, 1, 2)).toMatchObject({ formula: '=_xlfn.F(A2)', value: 'saved', cachedOnly: true });
  });

  it('is undone and redone with the structural edit', () => {
    const wb = mainWorkbook();
    cachedOnly(wb, undefined, 9, 5, LET);
    wb.recordHistory();
    wb.getSheet().insertRows(0, 2);
    expect(cellAt(wb, 11, 5)?.formula).not.toBe(LET);
    expect(wb.undo()).toBe(true);
    expect(cellAt(wb, 9, 5)).toMatchObject({ formula: LET, value: 99, cachedOnly: true });
    expect(wb.redo()).toBe(true);
    expect(cellAt(wb, 11, 5)?.formula).toBe('=_xlfn.LET(_xlpm.x,A3,_xlpm.y,B4:C5,_xlpm.x+_xlpm.y)');
    wb.recordHistory();
    wb.getSheet().deleteRows(0, 2);
    expect(cellAt(wb, 9, 5)?.formula).toBe(LET);
    wb.undo();
    expect(cellAt(wb, 11, 5)?.formula).toBe('=_xlfn.LET(_xlpm.x,A3,_xlpm.y,B4:C5,_xlpm.x+_xlpm.y)');
  });

  it('after a deleted reference, undo brings the original text back', () => {
    const wb = mainWorkbook();
    cachedOnly(wb, undefined, 9, 0, '=_xlfn.F(A1,B2:B3)');
    wb.recordHistory();
    wb.getSheet().deleteRows(0, 1);
    expect(cellAt(wb, 8, 0)?.formula).toBe('=_xlfn.F(#REF!,B1:B2)');
    wb.undo();
    expect(cellAt(wb, 9, 0)?.formula).toBe('=_xlfn.F(A1,B2:B3)');
  });
});

describe('collaboration: the shifted text is the cell’s new value (last writer wins per cell)', () => {
  afterEach(() => __resetInMemoryRooms());

  async function pair(): Promise<{ a: WorkbookImpl; b: WorkbookImpl; detach: () => void }> {
    const a = mainWorkbook('room');
    cachedOnly(a, undefined, 9, 5, LET);
    a.setCellValue(undefined, 0, 0, 1);
    const ha = await a.attachCollab(new InMemoryProvider(), identity('A'), { roomId: 'shift' });
    const b = new WorkbookImpl('b', 'B');
    const hb = await b.attachCollab(new InMemoryProvider(), identity('B'), { roomId: 'shift' });
    return { a, b, detach: () => (ha.detach(), hb.detach()) };
  }

  const snapshot = (wb: WorkbookImpl) =>
    [...wb.getSheet().entries()].map(([r, c, cell]) => [r, c, cell.formula, cell.value, cell.cachedOnly]).sort((x, y) => String(x).localeCompare(String(y)));

  it('a structural edit by one client reaches the other with the shifted text, and both converge', async () => {
    const { a, b, detach } = await pair();
    expect(cellAt(b, 9, 5)).toMatchObject({ formula: LET, cachedOnly: true });
    a.getSheet().insertRows(0, 2);
    for (const wb of [a, b]) expect(cellAt(wb, 11, 5)).toMatchObject({ formula: '=_xlfn.LET(_xlpm.x,A3,_xlpm.y,B4:C5,_xlpm.x+_xlpm.y)', value: 99, cachedOnly: true });
    b.getSheet().deleteCols(0, 1);
    for (const wb of [a, b]) expect(cellAt(wb, 11, 4)).toMatchObject({ formula: '=_xlfn.LET(_xlpm.x,#REF!,_xlpm.y,A4:B5,_xlpm.x+_xlpm.y)', cachedOnly: true });
    expect(snapshot(a)).toEqual(snapshot(b));
    detach();
  });

  // Both peers edit before either one's change reaches the other: each provider holds its outgoing doc messages until released.
  class GatedProvider extends InMemoryProvider {
    private held: Uint8Array[] | undefined;
    hold(): void {
      this.held = [];
    }
    send(channel: CollabChannel, payload: Uint8Array): void {
      if (this.held && channel === 'doc') this.held.push(payload);
      else super.send(channel, payload);
    }
    release(): void {
      const held = this.held ?? [];
      this.held = undefined;
      for (const payload of held) super.send('doc', payload);
    }
  }

  async function gatedPair(): Promise<{ a: WorkbookImpl; b: WorkbookImpl; ga: GatedProvider; gb: GatedProvider; detach: () => void }> {
    const a = mainWorkbook('room');
    cachedOnly(a, undefined, 9, 5, LET);
    const [ga, gb] = [new GatedProvider(), new GatedProvider()];
    const ha = await a.attachCollab(ga, identity('A'), { roomId: 'gated' });
    const b = new WorkbookImpl('b', 'B');
    const hb = await b.attachCollab(gb, identity('B'), { roomId: 'gated' });
    return { a, b, ga, gb, detach: () => (ha.detach(), hb.detach()) };
  }

  // What the LET cell (row 9, column F) is after one structural edit alone: [row, column, text]. Structure sync is last writer wins
  // (a concurrent structural edit is lost: pre-existing), and the cell's text is rewritten by the edit that survives only, so the
  // outcome is always one of the two single-edit results: never both edits and never a second shift of the text.
  const LET_AT = (x: string, range: string): string => `=_xlfn.LET(_xlpm.x,${x},_xlpm.y,${range},_xlpm.x+_xlpm.y)`;
  const alone = {
    insertRow1: [10, 5, LET_AT('A2', 'B3:C4')],
    insertRow2: [11, 5, LET_AT('A3', 'B4:C5')],
    deleteRow1: [8, 5, LET_AT('#REF!', 'B1:C2')],
    insertCol1: [9, 6, LET_AT('B1', 'C2:D3')],
  } as const;
  const concurrent: Array<[string, (w: WorkbookImpl) => void, (w: WorkbookImpl) => void, ReadonlyArray<readonly [number, number, string]>]> = [
    ['the same row inserted by both', (w) => w.getSheet().insertRows(0, 1), (w) => w.getSheet().insertRows(0, 1), [alone.insertRow1]],
    ['different inserts', (w) => w.getSheet().insertRows(0, 2), (w) => w.getSheet().insertRows(0, 1), [alone.insertRow2, alone.insertRow1]],
    ['an insert and a delete', (w) => w.getSheet().insertRows(0, 2), (w) => w.getSheet().deleteRows(0, 1), [alone.insertRow2, alone.deleteRow1]],
    ['a row insert and a column insert', (w) => w.getSheet().insertRows(0, 2), (w) => w.getSheet().insertCols(0, 1), [alone.insertRow2, alone.insertCol1]],
  ];

  it.each(concurrent)('concurrent edits (%s): both orders converge on one single-edit result, no double shift', async (_name, editA, editB, allowed) => {
    for (const order of ['a then b', 'b then a']) {
      __resetInMemoryRooms();
      const { a, b, ga, gb, detach } = await gatedPair();
      ga.hold();
      gb.hold();
      editA(a);
      editB(b);
      const [first, second] = order === 'a then b' ? [ga, gb] : [gb, ga];
      first.release();
      second.release();
      const cells = (wb: WorkbookImpl) => [...wb.getSheet().entries()].map(([r, c, cell]) => [r, c, cell.formula] as const);
      expect(cells(a)).toEqual(cells(b));
      expect(cells(a)).toHaveLength(1);
      expect(allowed.map((x) => JSON.stringify(x))).toContain(JSON.stringify(cells(a)[0]));
      expect(cellAt(a, cells(a)[0][0], cells(a)[0][1])).toMatchObject({ value: 99, cachedOnly: true });
      detach();
    }
  });

  it('undo of the structural edit restores the text on both clients', async () => {
    const { a, b, detach } = await pair();
    a.getSheet().insertRows(0, 2);
    expect(a.undo()).toBe(true);
    for (const wb of [a, b]) expect(cellAt(wb, 9, 5)).toMatchObject({ formula: LET, value: 99, cachedOnly: true });
    expect(snapshot(a)).toEqual(snapshot(b));
    detach();
  });
});
