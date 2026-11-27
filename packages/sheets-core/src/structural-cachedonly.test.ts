import { InMemoryProvider, __resetInMemoryRooms, type CollabIdentity } from '@b32nio/spindle-shared';
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
    sheet.deleteCols(2, 1); // C is a corner of the range: the range is gone
    expect(cellAt(wb, 10, 5)?.formula).toBe('=_xlfn.LET(_xlpm.x,A2,_xlpm.y,#REF!,_xlpm.x+_xlpm.y)');
    expect(cellAt(wb, 10, 5)).toMatchObject({ value: 99, cachedOnly: true });
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

  it('both clients inserting the same row at once end identical (each shifted the text: the cell is one last write)', async () => {
    const { a, b, detach } = await pair();
    a.getSheet().insertRows(0, 1);
    b.getSheet().insertRows(0, 1);
    expect(snapshot(a)).toEqual(snapshot(b));
    expect(cellAt(a, 11, 5)?.cachedOnly).toBe(true);
    detach();
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
