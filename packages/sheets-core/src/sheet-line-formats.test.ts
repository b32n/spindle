import { WorkbookImpl } from './workbook';

function setup() {
  const wb = new WorkbookImpl('wb', 'WB');
  const sheet = wb.getSheet(undefined)!;
  const s1 = wb.getStylePool().getOrCreate({ bold: true });
  const s2 = wb.getStylePool().getOrCreate({ italic: true });
  const f1 = wb.getFormatPool().getOrCreate({ type: 'percentage', decimalPlaces: 0 });
  return { wb, sheet, s1, s2, f1 };
}

describe('column and row default formats', () => {
  it('a cell typed into a formatted column starts with its style and format, a row\'s wins, and an explicit style is kept', () => {
    const { wb, sheet, s1, s2, f1 } = setup();
    sheet.setColFormat(1, { styleId: s1, formatId: f1 });
    sheet.setRowFormat(3, { styleId: s2 });
    wb.setCellValue(undefined, 0, 1, 5);
    expect(wb.getCell(undefined, 0, 1)).toMatchObject({ styleId: s1, formatId: f1 });
    wb.setCellValue(undefined, 3, 1, 5);
    expect(wb.getCell(undefined, 3, 1)).toMatchObject({ styleId: s2 });
    expect(wb.getCell(undefined, 3, 1)?.formatId).toBeUndefined();
    wb.setCell(undefined, 6, 1, { value: 1, styleId: s2 });
    expect(wb.getCell(undefined, 6, 1)?.styleId).toBe(s2);
    expect(wb.getCell(undefined, 6, 1)?.formatId).toBeUndefined();
    wb.setCell(undefined, 7, 1, { value: 1, styleId: undefined });
    expect(wb.getCell(undefined, 7, 1)?.styleId).toBeUndefined();
    // an existing cell is not touched, a cell outside both is plain
    wb.setCellValue(undefined, 0, 1, 6);
    expect(wb.getCell(undefined, 0, 1)?.styleId).toBe(s1);
    wb.setCellValue(undefined, 0, 5, 1);
    expect(wb.getCell(undefined, 0, 5)?.styleId).toBeUndefined();
  });

  it('is cleared by an empty format, and a clearing of an absent one does nothing', () => {
    const { sheet, s1 } = setup();
    sheet.setColFormat(1, { styleId: s1 });
    sheet.setColFormat(1, {});
    expect(sheet.getColFormat(1)).toBeUndefined();
    sheet.setColFormat(9, undefined);
    sheet.setRowFormat(2, { styleId: s1 });
    sheet.setRowFormat(2, undefined);
    sheet.setRowFormat(30, undefined);
    expect(sheet.getRowFormat(2)).toBeUndefined();
    expect(sheet.getColFormat(400)).toBeUndefined();
    expect(sheet.getRowFormat(400)).toBeUndefined();
  });

  it('follows its column and row through insert and delete, and goes with a deleted one', () => {
    const { sheet, s1, s2 } = setup();
    sheet.setColFormat(2, { styleId: s1 });
    sheet.setRowFormat(4, { styleId: s2 });
    sheet.insertCols(0, 1);
    sheet.insertRows(0, 2);
    expect(sheet.getColFormat(3)?.styleId).toBe(s1);
    expect(sheet.getRowFormat(6)?.styleId).toBe(s2);
    expect(sheet.getColFormat(2)).toBeUndefined();
    sheet.deleteCols(3, 1);
    sheet.deleteRows(6, 1);
    expect(sheet.config.colStyles?.size).toBe(0);
    expect(sheet.config.rowStyles?.size).toBe(0);
  });

  it('is part of undo and redo, getData and setData (numeric legacy keys too)', () => {
    const { wb, sheet, s1 } = setup();
    wb.setCellValue(undefined, 0, 0, 1);
    sheet.setColFormat(2, { styleId: s1 });
    sheet.setRowFormat(8, { styleId: s1 });
    wb.recordHistory();
    wb.setCellValue(undefined, 2, 0, 3);
    wb.undo();
    expect(wb.getSheet(undefined)!.getColFormat(2)?.styleId).toBe(s1);
    wb.redo();
    expect(wb.getSheet(undefined)!.getColFormat(2)?.styleId).toBe(s1);

    const data = wb.getData();
    const copy = new WorkbookImpl('c', 'C');
    copy.setData(data);
    expect(copy.getSheet(undefined)!.getColFormat(2)?.styleId).toBe(s1);
    expect(copy.getSheet(undefined)!.getRowFormat(8)?.styleId).toBe(s1);

    const legacy = new WorkbookImpl('l', 'L');
    legacy.setData({
      ...data,
      sheets: [{ id: 's', name: 'S', cells: [{ key: '0:0', cell: { value: 1 } }], config: { colStyles: [[3, { styleId: s1 }]], rowStyles: [[4, { styleId: s1 }]] }, rowCount: 10, colCount: 10 }],
    });
    expect(legacy.getSheet('s')!.getColFormat(3)?.styleId).toBe(s1);
    expect(legacy.getSheet('s')!.getRowFormat(4)?.styleId).toBe(s1);
  });
});
