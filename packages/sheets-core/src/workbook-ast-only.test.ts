import { WorkbookImpl } from './workbook';

// A cell that carries a parsed formula (`formulaAst`) and no text reaches `setData` only from a hand-made or damaged document
// (an editor's own cells always have both). The formula graph registers it under an empty text; with no text to evaluate, the
// cell keeps its value, whatever it was set to, and a change to its precedents does not move it.
describe('setData with a cell that has a formula AST and no formula text', () => {
  it.each([[true], [false]])('registers it and keeps its value (keepCachedValues %s)', (keepCachedValues) => {
    const source = new WorkbookImpl('src', 'Src');
    source.setCellValue(undefined, 0, 0, 2);
    source.setFormula(undefined, 0, 1, '=A1*10');
    const data = source.getData();
    const cell = data.sheets[0].cells.find(({ cell: c }) => c.formulaAst)!.cell;
    delete cell.formula;
    cell.value = 20;

    const wb = new WorkbookImpl('wb', 'WB');
    wb.setData(data, { keepCachedValues });
    expect(wb.getCell(undefined, 0, 1)).toMatchObject({ value: 20 });
    expect(wb.getCell(undefined, 0, 1)?.formula).toBeUndefined();
    wb.setCellValue(undefined, 0, 0, 3);
    expect(wb.getCell(undefined, 0, 1)?.value).toBe(20);
  });
});
