import { WorkbookImpl } from './workbook';

// A formula whose result is an error keeps its text in step with structural edits, like any other: the references in the file written
// from it must be the ones the cell has (found with `=@A1:A3*2` outside its range, which is #VALUE!).

describe('the text of a formula with an error result follows insert and delete', () => {
  const book = (formula: string, row: number): WorkbookImpl => {
    const wb = new WorkbookImpl('wb', 'WB');
    wb.setCellValue(undefined, 0, 0, 1);
    wb.setFormula(undefined, row, 5, formula);
    return wb;
  };

  it.each([
    ['an error thrown while evaluating (#VALUE!, outside the range of an `@`)', '=@A1:A3*2', '#VALUE!', '=@A3:A5*2'],
    ['a division by zero', '=1/0+A1', '#DIV/0!', '=1/0+A3'],
    ['a result past the number range', '=A1*10^400', '#DIV/0!', '=A3*10^400'],
  ])('%s', (_name, formula, error, shifted) => {
    const wb = book(formula, 4);
    expect(wb.getCell(undefined, 4, 5)?.value).toBe(error);
    wb.getSheet().insertRows(0, 2);
    expect(wb.getCell(undefined, 6, 5)).toMatchObject({ formula: shifted, value: error });
  });

  it('a delete too', () => {
    const wb = book('=1/0+B2', 6);
    wb.setCellValue(undefined, 1, 1, 5);
    wb.getSheet().deleteRows(0, 1);
    expect(wb.getCell(undefined, 5, 5)).toMatchObject({ formula: '=1/0+B1', value: '#DIV/0!' });
  });
});
