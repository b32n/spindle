import { InMemoryProvider, __resetInMemoryRooms, type CollabIdentity } from '@b32nio/spindle-shared';
import { WorkbookImpl } from './workbook';
import type { Cell } from './types';

// A structural edit scans a sheet's cells for formulas / cached-only formulas only when its cell store may hold one (`CellMap` flags:
// set by every write that carries one, cleared only by a scan that saw none). These tests make sure the skip never loses needed work.

const identity = (name: string): CollabIdentity => ({ userId: name, displayName: name, color: '#123456' });
const F = '=_xlfn.F(A5,B1:B3)';

function book(id = 'wb'): WorkbookImpl {
  const wb = new WorkbookImpl(id, 'WB');
  wb.renameSheet(wb.getSheet().id, 'Main');
  return wb;
}
const cachedOnly = (wb: WorkbookImpl, r: number, c: number, formula = F): void =>
  wb.setCell(undefined, r, c, { formula, formulaAst: undefined, value: 99 as Cell['value'], cachedOnly: true });
const flags = (wb: WorkbookImpl) => {
  const { cells } = wb.getSheet() as unknown as { cells: { maybeFormulas: boolean; maybeCachedOnly: boolean } };
  return [cells.maybeFormulas, cells.maybeCachedOnly];
};
const at = (wb: WorkbookImpl, r: number, c: number): Cell | undefined => wb.getCell(undefined, r, c);

describe('the scan skip of structural edits', () => {
  it('a value-only sheet has no flags and a structural edit leaves them clear', () => {
    const wb = book();
    wb.setCellValue(undefined, 0, 0, 1);
    expect(flags(wb)).toEqual([false, false]);
    wb.getSheet().deleteRows(0, 1);
    wb.getSheet().insertCols(0, 1);
    expect(flags(wb)).toEqual([false, false]);
  });

  it('a scan that saw none clears the flags, and a later formula or cached-only cell sets them again and is handled', () => {
    const wb = book();
    wb.setFormula(undefined, 0, 1, '=A1');
    cachedOnly(wb, 3, 3);
    expect(flags(wb)).toEqual([true, true]);
    wb.setCellValue(undefined, 0, 1, 5); // the formula is overwritten by a value
    wb.setCell(undefined, 3, 3, { formula: undefined, cachedOnly: undefined });
    wb.getSheet().deleteRows(9, 1);
    expect(flags(wb)).toEqual([false, false]);
    wb.setFormula(undefined, 5, 0, '=SUM(B1:B5)');
    cachedOnly(wb, 6, 0);
    expect(flags(wb)).toEqual([true, true]);
    wb.getSheet().deleteRows(0, 1);
    expect(at(wb, 4, 0)?.formula).toBe('=SUM(B1:B4)');
    expect(at(wb, 5, 0)?.formula).toBe('=_xlfn.F(A4,B1:B2)');
  });

  it('keeps the flags while formulas remain, per sheet', () => {
    const wb = book();
    const other = wb.addSheet('Other');
    wb.setFormula(undefined, 0, 0, '=1+1');
    wb.getSheet().deleteRows(5, 1);
    expect(flags(wb)).toEqual([true, false]);
    wb.setCell(other.id, 0, 0, { formula: "=_xlfn.F(Main!A5)", formulaAst: undefined, value: 1, cachedOnly: true });
    wb.getSheet().deleteRows(0, 1);
    expect(wb.getCell(other.id, 0, 0)?.formula).toBe('=_xlfn.F(Main!A4)');
  });

  it('undo and redo that bring a formula back set the flags', () => {
    const wb = book();
    wb.setFormula(undefined, 0, 1, '=SUM(A2:A5)');
    cachedOnly(wb, 7, 0);
    wb.recordHistory();
    wb.setCellValue(undefined, 0, 1, 1);
    wb.setCell(undefined, 7, 0, { formula: undefined, cachedOnly: undefined });
    wb.getSheet().deleteRows(20, 1); // the scans see none and clear the flags
    expect(flags(wb)).toEqual([false, false]);
    expect(wb.undo()).toBe(true);
    wb.getSheet().deleteRows(1, 1);
    expect(at(wb, 0, 1)?.formula).toBe('=SUM(A2:A4)');
    expect(at(wb, 6, 0)?.formula).toBe('=_xlfn.F(A4,B1:B2)');
    expect(wb.redo()).toBe(true);
    wb.getSheet().insertRows(0, 1);
    wb.setFormula(undefined, 2, 2, '=A3');
    expect(flags(wb)[0]).toBe(true);
  });

  it('data loaded with setData (an import) sets the flags', () => {
    const src = book();
    src.setCellValue(undefined, 0, 0, 1);
    src.setFormula(undefined, 0, 1, '=SUM(A2:A5)');
    cachedOnly(src, 7, 0);
    const wb = book('other');
    wb.setCellValue(undefined, 0, 0, 1);
    wb.getSheet().deleteRows(30, 1);
    expect(flags(wb)).toEqual([false, false]);
    wb.setData(src.getData());
    expect(flags(wb)).toEqual([true, true]);
    wb.getSheet().deleteRows(1, 1);
    expect(at(wb, 0, 1)?.formula).toBe('=SUM(A2:A4)');
    expect(at(wb, 6, 0)?.formula).toBe('=_xlfn.F(A4,B1:B2)');
  });

  describe('collaboration', () => {
    afterEach(() => __resetInMemoryRooms());

    it('a cached-only cell and a formula that arrive from the other client set the flags, and a structural edit then shifts them', async () => {
      const a = book('room');
      for (let r = 0; r < 5; r++) a.setCellValue(undefined, r, 0, r + 1);
      const b = new WorkbookImpl('b', 'B');
      const ha = await a.attachCollab(new InMemoryProvider(), identity('A'), { roomId: 'skip' });
      const hb = await b.attachCollab(new InMemoryProvider(), identity('B'), { roomId: 'skip' });
      expect(flags(b)).toEqual([false, false]);
      b.getSheet().deleteRows(20, 1); // a scan that sees none
      cachedOnly(a, 7, 0);
      a.setFormula(undefined, 0, 1, '=SUM(A2:A5)');
      expect(at(b, 7, 0)).toMatchObject({ formula: F, cachedOnly: true });
      expect(flags(b)).toEqual([true, true]);
      b.getSheet().deleteRows(1, 1);
      for (const wb of [a, b]) {
        expect(at(wb, 6, 0)?.formula).toBe('=_xlfn.F(A4,B1:B2)');
        expect(at(wb, 0, 1)?.formula).toBe('=SUM(A2:A4)');
      }
      ha.detach();
      hb.detach();
    });
  });

  describe('edits that do nothing, and cells without text', () => {
    it('deleting zero rows or columns leaves every formula and cached-only text as it is', () => {
      const wb = book();
      wb.setCellValue(undefined, 0, 0, 1);
      wb.setFormula(undefined, 0, 1, '=SUM(A1:A5)');
      cachedOnly(wb, 7, 0);
      wb.getSheet().deleteRows(1, 0);
      wb.getSheet().deleteCols(0, 0);
      expect(at(wb, 0, 1)).toMatchObject({ formula: '=SUM(A1:A5)', value: 1 });
      expect(at(wb, 7, 0)?.formula).toBe(F);
    });

    it('a formula cell that holds an AST but no formula text is retargeted too (the dependency graph follows)', () => {
      const wb = book();
      for (let r = 0; r < 5; r++) wb.setCellValue(undefined, r, 0, r + 1);
      wb.setFormula(undefined, 9, 1, '=SUM(A2:A5)');
      const { formulaAst } = at(wb, 9, 1) as Cell;
      wb.getSheet().setCell(9, 1, { formula: undefined });
      expect(at(wb, 9, 1)?.formula).toBeUndefined();
      expect(at(wb, 9, 1)?.formulaAst).toBe(formulaAst);
      wb.getSheet().deleteRows(4, 1);
      expect(at(wb, 8, 1)?.formulaAst).not.toBe(formulaAst);
    });

    it('an AST-only cell written onto a sheet that never held a formula raises the flag on its own', () => {
      const donor = book('donor');
      for (let r = 0; r < 5; r++) donor.setCellValue(undefined, r, 0, r + 1);
      donor.setFormula(undefined, 9, 1, '=SUM(A2:A5)');
      const { formulaAst } = at(donor, 9, 1) as Cell;
      // the same sheet, flags reset: no formula write has been seen since, so only the AST can raise the flag
      const store = donor.getSheet() as unknown as { cells: { maybeFormulas: boolean } };
      donor.getSheet().setCell(9, 1, { formula: undefined, formulaAst: undefined, value: 0 as Cell['value'] });
      store.cells.maybeFormulas = false;
      donor.getSheet().setCell(9, 2, { value: 10 as Cell['value'], formulaAst });
      expect(at(donor, 9, 2)?.formula).toBeUndefined();
      expect(flags(donor)[0]).toBe(true);
      donor.getSheet().deleteRows(4, 1);
      expect(at(donor, 8, 2)?.formulaAst).not.toBe(formulaAst);
    });
  });
});
