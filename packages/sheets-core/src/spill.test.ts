import { SpillIndex } from './spill';
import { WorkbookImpl } from './workbook';

describe('SpillIndex', () => {
  it('exposes spilled values and coverage for a registered region', () => {
    const index = new SpillIndex();
    index.register({ anchorRow: 0, anchorCol: 0, rows: 3, cols: 1, values: [[1], [2], [3]] });

    expect(index.spilledValueAt(1, 0)).toBe(2);
    expect(index.spilledValueAt(2, 0)).toBe(3);
    expect(index.spilledValueAt(0, 0)).toBeUndefined(); // the anchor is not "covered"
    expect(index.isCovered(1, 0)).toBe(true);
    expect(index.isCovered(0, 0)).toBe(false);
    expect(index.anchorOf(2, 0)).toEqual({ row: 0, col: 0 });
  });

  it('unregister removes the region and its coverage', () => {
    const index = new SpillIndex();
    index.register({ anchorRow: 0, anchorCol: 0, rows: 2, cols: 1, values: [[1], [2]] });
    const previous = index.unregister(0, 0);

    expect(previous?.rows).toBe(2);
    expect(index.isCovered(1, 0)).toBe(false);
    expect(index.spilledValueAt(1, 0)).toBeUndefined();
  });
});

describe('spill — workbook integration', () => {
  it('a 2D array formula fills a block of cells', () => {
    const wb = new WorkbookImpl('wb', 'Test');
    wb.setFormula(undefined, 0, 0, '=SEQUENCE(3)');
    expect(wb.getCellCalculatedValue(undefined, 0, 0)).toBe(1);
    expect(wb.getCellCalculatedValue(undefined, 1, 0)).toBe(2);
    expect(wb.getCellCalculatedValue(undefined, 2, 0)).toBe(3);

    wb.setFormula(undefined, 0, 0, '=SEQUENCE(2, 2)');
    expect(wb.getCellCalculatedValue(undefined, 1, 1)).toBe(4);
    // The old footprint shrank — row 3 is no longer part of the spill.
    expect(wb.getCellCalculatedValue(undefined, 2, 0)).toBe(0);
  });

  it('raises #SPILL! when the target block is blocked', () => {
    const wb = new WorkbookImpl('wb', 'Test');
    wb.setCellValue(undefined, 0, 1, 'blocker');
    wb.setFormula(undefined, 0, 0, '=SEQUENCE(1, 3)');
    expect(wb.getCellCalculatedValue(undefined, 0, 0)).toBe('#SPILL!');
  });

  it('a formula reading a spilled cell recalculates when the spill changes', () => {
    const wb = new WorkbookImpl('wb', 'Test');
    wb.setFormula(undefined, 0, 0, '=SEQUENCE(3)'); // A1:A3 = 1, 2, 3
    wb.setFormula(undefined, 0, 2, '=A2'); // C1 reads the spilled cell A2
    expect(wb.getCellCalculatedValue(undefined, 0, 2)).toBe(2);

    wb.setFormula(undefined, 0, 0, '=SEQUENCE(3, 1, 10)'); // A1:A3 = 10, 11, 12
    expect(wb.getCellCalculatedValue(undefined, 0, 2)).toBe(11);
  });

  it('SUM over a spilled range stays in sync', () => {
    const wb = new WorkbookImpl('wb', 'Test');
    wb.setFormula(undefined, 0, 0, '=SEQUENCE(4)'); // A1:A4 = 1, 2, 3, 4
    wb.setFormula(undefined, 0, 2, '=SUM(A1:A4)');
    expect(wb.getCellCalculatedValue(undefined, 0, 2)).toBe(10);
  });

  it('deleting the anchor releases the spill and frees its cells', () => {
    const wb = new WorkbookImpl('wb', 'Test');
    wb.setFormula(undefined, 0, 0, '=SEQUENCE(3)');
    expect(wb.isSpilledCell(undefined, 1, 0)).toBe(true);

    wb.setCellValue(undefined, 0, 0, null); // delete the anchor formula
    expect(wb.isSpilledCell(undefined, 1, 0)).toBe(false);
    expect(wb.getSpilledValue(undefined, 1, 0)).toBeUndefined();
    expect(wb.getCellCalculatedValue(undefined, 1, 0)).toBe(0);
  });

  it('writing a value over the anchor clears the formula and the spill', () => {
    const wb = new WorkbookImpl('wb', 'Test');
    wb.setFormula(undefined, 0, 0, '=SEQUENCE(3)');
    wb.setCellValue(undefined, 0, 0, 'text');

    const anchor = wb.getCell(undefined, 0, 0);
    expect(anchor?.value).toBe('text');
    expect(anchor?.formula).toBeUndefined();
    expect(anchor?.formulaAst).toBeUndefined();
    expect(wb.isSpilledCell(undefined, 1, 0)).toBe(false);
  });

  it('re-evaluating a blocked anchor spills once the blocker is cleared', () => {
    const wb = new WorkbookImpl('wb', 'Test');
    wb.setCellValue(undefined, 1, 0, 'blocker');
    wb.setFormula(undefined, 0, 0, '=SEQUENCE(3)');
    expect(wb.getCellCalculatedValue(undefined, 0, 0)).toBe('#SPILL!');

    wb.setCellValue(undefined, 1, 0, null); // clear the blocker
    wb.setFormula(undefined, 0, 0, '=SEQUENCE(3)'); // re-enter the formula
    expect(wb.getCellCalculatedValue(undefined, 0, 0)).toBe(1);
    expect(wb.getSpilledValue(undefined, 1, 0)).toBe(2);
  });
});

describe('#SPILL! recovery', () => {
  const a1Spill = (wb: WorkbookImpl) => wb.getCellCalculatedValue(undefined, 0, 0);

  it('spills by itself once the blocker is cleared (no re-entering the formula)', () => {
    const wb = new WorkbookImpl('wb', 'Test');
    wb.setCellValue(undefined, 1, 0, 'blocker');
    wb.setFormula(undefined, 0, 0, '=SEQUENCE(3)');
    expect(a1Spill(wb)).toBe('#SPILL!');

    wb.setCellValue(undefined, 1, 0, null);
    expect(a1Spill(wb)).toBe(1);
    expect(wb.getSpilledValue(undefined, 1, 0)).toBe(2);
    expect(wb.getSpilledValue(undefined, 2, 0)).toBe(3);
  });

  it('recovers for a 2-D footprint when a blocker in the far corner is cleared', () => {
    const wb = new WorkbookImpl('wb', 'Test');
    wb.setCellValue(undefined, 2, 1, 'x');
    wb.setFormula(undefined, 0, 0, '=SEQUENCE(3,2)');
    expect(a1Spill(wb)).toBe('#SPILL!');
    wb.setCellValue(undefined, 2, 1, null);
    expect(a1Spill(wb)).toBe(1);
    expect(wb.getSpilledValue(undefined, 2, 1)).toBe(6);
  });

  it('a dependent of the blocked anchor updates when it recovers', () => {
    const wb = new WorkbookImpl('wb', 'Test');
    wb.setCellValue(undefined, 1, 0, 'blocker');
    wb.setFormula(undefined, 0, 0, '=SEQUENCE(3)');
    wb.setFormula(undefined, 0, 2, '=A1');
    expect(wb.getCellCalculatedValue(undefined, 0, 2)).toBe('#SPILL!');
    wb.setCellValue(undefined, 1, 0, null);
    expect(wb.getCellCalculatedValue(undefined, 0, 2)).toBe(1);
  });

  it('stays #SPILL! while another cell still blocks, and recovers when the last is cleared', () => {
    const wb = new WorkbookImpl('wb', 'Test');
    wb.setCellValue(undefined, 1, 0, 'a');
    wb.setCellValue(undefined, 2, 0, 'b');
    wb.setFormula(undefined, 0, 0, '=SEQUENCE(3)');
    wb.setCellValue(undefined, 1, 0, null);
    expect(a1Spill(wb)).toBe('#SPILL!');
    wb.setCellValue(undefined, 2, 0, null);
    expect(a1Spill(wb)).toBe(1);
  });

  it('recovers when the blocker is another spill that goes away', () => {
    const wb = new WorkbookImpl('wb', 'Test');
    wb.setFormula(undefined, 0, 1, '=SEQUENCE(1,3)'); // B1:D1
    wb.setFormula(undefined, 0, 0, '=SEQUENCE(1,3)'); // A1:C1 collides with B1's spill
    expect(a1Spill(wb)).toBe('#SPILL!');
    wb.setCellValue(undefined, 0, 1, null);
    expect(a1Spill(wb)).toBe(1);
    expect(wb.getSpilledValue(undefined, 0, 2)).toBe(3);
  });

  it('recovers when the blocker is cleared by deleting its row', () => {
    const wb = new WorkbookImpl('wb', 'Test');
    wb.setCellValue(undefined, 1, 0, 'blocker');
    wb.setFormula(undefined, 0, 0, '=SEQUENCE(3)');
    expect(a1Spill(wb)).toBe('#SPILL!');
    wb.getSheet().deleteRows(1, 1);
    expect(a1Spill(wb)).toBe(1);
  });
});
