import { WorkbookImpl } from '../../workbook';

// Run by `ref-shift-hostile.test.ts` in a child process: a workbook of `count` cached-only cells whose text has a hostile shape
// (argv[2], argv[3]), then one insert of a row and one delete of the same row on the sheet they all point at. Prints what it cost and gave
// as one JSON line. The test's heap and time limits on the child are the assertions about memory and time; the figures here are for
// the PR's table.
const [shape, count] = [process.argv[2], Number(process.argv[3])];

const LIMIT = 8192; // the longest formula the importer keeps

/** The text of cached-only cell `i`, by shape. */
function text(i: number): string {
  switch (shape) {
    case 'typical':
      return `=_xlfn.F(A${i + 1},B2:C3,$D$4,Sheet1!E5,'Sheet1'!F:G,1:2)`;
    // thousands of short words: a sheet prefix must not be searched for through all of them at each word
    case 'words':
      return `=${'ab '.repeat(2700)}A1`;
    case 'words-bang':
      return `=${'ab '.repeat(2700)}!A1`;
    case 'words-3d':
      return `=${'ab '.repeat(1300)}:cd ${'ab '.repeat(1300)}!A1`;
    // the most references a formula can hold
    case 'refs':
      return `=${'A1+'.repeat(2700)}B2`;
    case 'ranges':
      return `=SUM(${'A1:B2,'.repeat(1300)}C3)`;
    case 'prefixed':
      return `=${'Sheet1!A1+'.repeat(800)}1`;
    // text the grammar rejects, in bulk
    case 'brackets':
      return `=A1+${'['.repeat(LIMIT - 10)}`;
    case 'brackets-nested':
      return `=A1+T${'[a'.repeat(4000)}`;
    case 'quotes':
      return `=A1+${"'".repeat(LIMIT - 10)}`;
    case 'strings':
      return `=A1&"${'"" B2 '.repeat(1300)}"&B2`;
    case 'garbage':
      return `=A1${'€'.repeat(LIMIT - 10)}`;
    case 'colons':
      return `=A1${':A1'.repeat(2700)}`;
    case 'spills':
      return `=${'A1#+'.repeat(2000)}1`;
    case 'dollars':
      return `=${'$'.repeat(LIMIT - 10)}A1`;
    default:
      throw new Error(`unknown shape ${shape}`);
  }
}

const wb = new WorkbookImpl('probe', 'Probe');
const sheet = wb.getSheet();
for (let i = 0; i < count; i++) wb.setCell(undefined, i + 5, 0, { formula: text(i), formulaAst: undefined, value: 1, cachedOnly: true });
const t0 = Date.now();
sheet.insertRows(0, 1);
const insertMs = Date.now() - t0;
const t1 = Date.now();
sheet.deleteRows(0, 1);
const deleteMs = Date.now() - t1;
// insert + delete of the same row put every text back as it was
let kept = 0;
let restored = 0;
for (const [row, , cell] of sheet.entries()) {
  if (cell.cachedOnly) kept++;
  if (cell.formula === text(row - 5)) restored++;
}
console.log(JSON.stringify({ shape, count, insertMs, deleteMs, kept, restored, length: text(0).length }));
