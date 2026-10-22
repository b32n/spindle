import { WorkbookImpl } from './workbook';

const nextTick = () => Promise.resolve();

describe('undo history for unbatched edit loops', () => {
  it('a synchronous loop of edits is one undo step', () => {
    const w = new WorkbookImpl('wb', 'WB');
    w.setCellValue(undefined, 0, 0, 'before');
    return nextTick().then(() => {
      for (let i = 0; i < 100; i++) w.setCellValue(undefined, i, 0, i);
      w.setFormula(undefined, 0, 1, '=A1+1');
      expect(w.getCellValue(undefined, 99, 0)).toBe(99);
      expect(w.undo()).toBe(true);
      expect(w.getCellValue(undefined, 0, 0)).toBe('before');
      expect(w.getCellValue(undefined, 99, 0)).toBeNull();
      expect(w.getCell(undefined, 0, 1)).toBeUndefined();
      w.redo();
      expect(w.getCellValue(undefined, 99, 0)).toBe(99);
      expect(w.getCellValue(undefined, 0, 1)).toBe(1);
    });
  });

  it('edits in separate tasks stay separate undo steps', async () => {
    const w = new WorkbookImpl('wb', 'WB');
    w.setCellValue(undefined, 0, 0, 1);
    await nextTick();
    w.setCellValue(undefined, 0, 0, 2);
    await nextTick();
    w.setCellValue(undefined, 0, 0, 3);
    w.undo();
    expect(w.getCellValue(undefined, 0, 0)).toBe(2);
    w.undo();
    expect(w.getCellValue(undefined, 0, 0)).toBe(1);
    w.undo();
    expect(w.getCellValue(undefined, 0, 0)).toBeNull();
    expect(w.canUndo()).toBe(false);
  });

  it('an edit right after an undo starts a new step and clears redo', async () => {
    const w = new WorkbookImpl('wb', 'WB');
    w.setCellValue(undefined, 0, 0, 1);
    await nextTick();
    w.setCellValue(undefined, 0, 0, 2);
    w.undo();
    w.setCellValue(undefined, 0, 0, 5); // same task as the undo
    expect(w.canRedo()).toBe(false);
    w.undo();
    expect(w.getCellValue(undefined, 0, 0)).toBe(1);
  });

  it('structural operations keep their own undo step', () => {
    const w = new WorkbookImpl('wb', 'WB');
    w.setCellValue(undefined, 0, 1, 'covered');
    w.mergeCells({ startRow: 0, startCol: 0, endRow: 0, endCol: 1 });
    w.undo();
    expect(w.getSheet().getMergedRegions()).toEqual([]);
    expect(w.getCellValue(undefined, 0, 1)).toBe('covered');
  });

  it('keeps the history size cap', async () => {
    const w = new WorkbookImpl('wb', 'WB');
    for (let i = 0; i < 80; i++) {
      w.setCellValue(undefined, 0, 0, i);
      await nextTick();
    }
    let steps = 0;
    while (w.undo()) steps++;
    expect(steps).toBe(50);
  });
});

describe('bulk edit performance budget', () => {
  // Before coalescing every edit copied the whole workbook for undo, so N edits cost O(N²):
  // 4,000 edits took ~2 s and 20,000 would take close to a minute. Linear is a few tens of ms.
  it('20,000 unbatched value edits and 5,000 formulas finish within the budget', () => {
    const w = new WorkbookImpl('wb', 'WB');
    const start = performance.now();
    for (let i = 0; i < 20_000; i++) w.setCellValue(undefined, i % 5000, Math.floor(i / 5000), i);
    for (let i = 0; i < 5_000; i++) w.setFormula(undefined, i, 5, `=A${i + 1}*2`);
    const elapsed = performance.now() - start;
    expect(w.getCellValue(undefined, 4999, 5)).toBe(2 * 4999);
    expect(elapsed).toBeLessThan(4000);
  });
});
