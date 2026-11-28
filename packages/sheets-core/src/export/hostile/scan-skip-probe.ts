import { WorkbookImpl } from '../../workbook';

// Run by `scan-skip-cost.test.ts` in a child process: a sheet of `count` cells (argv[3]) that are all of one kind (argv[2]: values, formulas or cachedOnly), then one delete of a row
// and one insert of a column. Prints what each cost as one JSON line: the cost of a structural edit on a sheet with no formula
// (`values`) must not grow with the formula scans, and a sheet with formulas costs what it did.
const [kind, count] = [process.argv[2], Number(process.argv[3])];

const wb = new WorkbookImpl('probe', 'Probe');
const sheet = wb.getSheet();
const rows = Math.ceil(count / 1000);
for (let i = 0; i < count; i++) {
  const [row, col] = [Math.floor(i / 1000), i % 1000];
  if (kind === 'formulas') sheet.setCell(row, col, { value: 1, formula: `=SUM(A${row + 1}:B${row + 1})+${i}` });
  else if (kind === 'cachedOnly') sheet.setCell(row, col, { value: 1, formula: `=_xlfn.F(B${row + 1}:C${row + 3})`, cachedOnly: true });
  else sheet.setCell(row, col, { value: i });
}
// the formulas get their stable ASTs and dependency graph as a loaded file's do
if (kind === 'formulas') wb.setData(wb.getData());
const t0 = performance.now();
sheet.deleteRows(Math.floor(rows / 2), 1);
const deleteRowMs = performance.now() - t0;
const t1 = performance.now();
sheet.insertCols(500, 1);
const insertColMs = performance.now() - t1;
console.log(JSON.stringify({ kind, count, deleteRowMs: Math.round(deleteRowMs), insertColMs: Math.round(insertColMs) }));
