import { InMemoryProvider, __resetInMemoryRooms, type CollabIdentity } from '@b32nio/spindle-shared';
import { WorkbookImpl } from './workbook';

// A range keeps what survives of it when rows/columns are deleted and is #REF! only when all of it goes (Excel): the engine's
// stable-id path, with values, dependency tracking, undo, spills and collaboration (the text rule itself: ref-shift-excel-oracle.test.ts).

const identity = (name: string): CollabIdentity => ({ userId: name, displayName: name, color: '#123456' });

/** The sheet Main with A1:A5 = 1..5 and B10 = SUM(A2:A5). */
function book(id = 'wb'): WorkbookImpl {
  const wb = new WorkbookImpl(id, 'WB');
  wb.renameSheet(wb.getSheet().id, 'Main');
  [1, 2, 3, 4, 5].forEach((v, r) => wb.setCellValue(undefined, r, 0, v));
  wb.setFormula(undefined, 9, 1, '=SUM(A2:A5)');
  return wb;
}
/** The SUM formula of `book` (column B, wherever the deletes moved its row): its text and value. */
const sumOf = (wb: WorkbookImpl) => {
  for (const [r, c, cell] of wb.getSheet().entries()) if (c === 1 && cell.formula) return formulaOf(wb, r, 1);
  throw new Error('no formula');
};
const formulaOf = (wb: WorkbookImpl, r: number, c: number) => {
  const cell = wb.getCell(undefined, r, c);
  return { formula: cell?.formula, value: cell?.value };
};

describe('deleting rows under a range formula', () => {
  it.each([
    ['its last row', 4, 1, '=SUM(A2:A4)', 9],
    ['its first row', 1, 1, '=SUM(A2:A4)', 12],
    ['an inner row', 2, 1, '=SUM(A2:A4)', 11],
    ['rows 4 to 6', 3, 3, '=SUM(A2:A3)', 5],
    ['a row above it', 0, 1, '=SUM(A1:A4)', 14],
    ['a row below it', 5, 1, '=SUM(A2:A5)', 14],
  ])('%s: keeps what survives of it, with its new value', (_name, start, count, text, value) => {
    const wb = book();
    wb.getSheet().deleteRows(start, count);
    expect(sumOf(wb)).toEqual({ formula: text, value });
  });

  it('is #REF! only when every row of the range is deleted', () => {
    const wb = book();
    wb.getSheet().deleteRows(1, 4);
    expect(sumOf(wb)).toEqual({ formula: '=SUM(#REF!)', value: '#REF!' });
  });

  it('keeps tracking the rows that are left: editing one of them recalculates the formula', () => {
    const wb = book();
    wb.getSheet().deleteRows(1, 1); // A2:A5 -> A2:A4 (3, 4, 5)
    wb.setCellValue(undefined, 3, 0, 100);
    expect(sumOf(wb)).toEqual({ formula: '=SUM(A2:A4)', value: 3 + 4 + 100 });
    wb.getSheet().deleteRows(3, 1); // the last row of it, again
    expect(sumOf(wb)).toEqual({ formula: '=SUM(A2:A3)', value: 7 });
    wb.setCellValue(undefined, 2, 0, 10);
    expect(sumOf(wb).value).toBe(13);
  });

  it('an insert after a delete keeps extending the retargeted range', () => {
    const wb = book();
    wb.getSheet().deleteRows(4, 1);
    wb.getSheet().insertRows(2, 1);
    expect(sumOf(wb).formula).toBe('=SUM(A2:A5)');
    wb.setCellValue(undefined, 2, 0, 1000);
    expect(sumOf(wb).value).toBe(2 + 1000 + 3 + 4);
  });

  it('columns too: a whole-column range and a 2D range', () => {
    const wb = new WorkbookImpl('wb', 'WB');
    wb.setCellValue(undefined, 0, 0, 1);
    wb.setCellValue(undefined, 0, 1, 10);
    wb.setCellValue(undefined, 0, 2, 100);
    wb.setFormula(undefined, 2, 4, '=SUM(A:C)');
    wb.setFormula(undefined, 2, 5, '=SUM(B1:C1)');
    wb.getSheet().deleteCols(1, 1);
    expect(formulaOf(wb, 2, 3)).toEqual({ formula: '=SUM(A:B)', value: 101 });
    expect(formulaOf(wb, 2, 4)).toEqual({ formula: '=SUM(B1:B1)', value: 100 }); // C1 is B1 now
    wb.getSheet().deleteCols(1, 1);
    expect(formulaOf(wb, 2, 2)).toEqual({ formula: '=SUM(A:A)', value: 1 });
    expect(formulaOf(wb, 2, 3)).toEqual({ formula: '=SUM(#REF!)', value: '#REF!' }); // all of its columns are gone
  });

  it('a range on another sheet follows a delete of that sheet only', () => {
    const wb = book();
    const other = wb.addSheet('Other');
    wb.setFormula(other.id, 0, 0, '=SUM(Main!A2:A5)');
    wb.setFormula(other.id, 1, 0, '=SUM(A1:A5)');
    wb.getSheet(other.id).deleteRows(2, 1); // the other sheet's own rows: Main's range is not hit
    expect(wb.getCell(other.id, 0, 0)).toMatchObject({ formula: '=SUM(Main!A2:A5)', value: 14 });
    wb.getSheet().deleteRows(4, 1);
    expect(wb.getCell(other.id, 0, 0)).toMatchObject({ formula: '=SUM(Main!A2:A4)', value: 9 });
  });

  it('a reversed range and a range written with a 3D prefix', () => {
    const wb = book();
    wb.setFormula(undefined, 1, 1, '=SUM(A5:A2)');
    wb.setFormula(undefined, 2, 1, '=SUM(Sheet1:Sheet1!A2:A5)');
    wb.getSheet().deleteRows(4, 1);
    expect(formulaOf(wb, 1, 1)).toEqual({ formula: '=SUM(A4:A2)', value: 9 });
    expect(wb.getCell(undefined, 2, 1)?.formula).toBe('=SUM(Sheet1:Sheet1!A2:A5)'); // positional, as before
  });

  it('a spill whose source range loses a row shrinks with it', () => {
    const wb = book();
    wb.setFormula(undefined, 0, 3, '=TRANSPOSE(A2:A5)');
    expect(wb.getCellCalculatedValue(undefined, 0, 6)).toBe(5);
    wb.getSheet().deleteRows(4, 1);
    expect(formulaOf(wb, 0, 3).formula).toBe('=TRANSPOSE(A2:A4)');
    expect(wb.getCellCalculatedValue(undefined, 0, 5)).toBe(4);
    expect(wb.getCellCalculatedValue(undefined, 0, 6)).toBeFalsy();
  });

  it('undo restores the range and its formula, and a second delete behaves the same again', () => {
    const wb = book();
    wb.recordHistory();
    wb.getSheet().deleteRows(1, 1);
    expect(sumOf(wb).formula).toBe('=SUM(A2:A4)');
    expect(wb.undo()).toBe(true);
    expect(sumOf(wb)).toEqual({ formula: '=SUM(A2:A5)', value: 14 });
    wb.recordHistory();
    wb.getSheet().deleteRows(4, 1);
    expect(sumOf(wb)).toEqual({ formula: '=SUM(A2:A4)', value: 9 });
    expect(wb.undo()).toBe(true);
    expect(sumOf(wb)).toEqual({ formula: '=SUM(A2:A5)', value: 14 });
    expect(wb.redo()).toBe(true);
    expect(sumOf(wb)).toEqual({ formula: '=SUM(A2:A4)', value: 9 });
  });
});

describe('collaboration: a delete under a range formula', () => {
  afterEach(() => __resetInMemoryRooms());

  it('reaches the other client with the surviving range and the same values', async () => {
    const a = book('room');
    const ha = await a.attachCollab(new InMemoryProvider(), identity('A'), { roomId: 'range-delete' });
    const b = new WorkbookImpl('b', 'B');
    const hb = await b.attachCollab(new InMemoryProvider(), identity('B'), { roomId: 'range-delete' });
    a.getSheet().deleteRows(1, 1);
    for (const wb of [a, b]) expect(sumOf(wb)).toEqual({ formula: '=SUM(A2:A4)', value: 12 });
    b.getSheet().deleteRows(3, 1);
    for (const wb of [a, b]) expect(sumOf(wb)).toEqual({ formula: '=SUM(A2:A3)', value: 7 });
    b.setCellValue(undefined, 2, 0, 20);
    for (const wb of [a, b]) expect(sumOf(wb).value).toBe(23);
    ha.detach();
    hb.detach();
  });
});
