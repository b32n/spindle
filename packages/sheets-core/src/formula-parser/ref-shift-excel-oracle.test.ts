// An oracle table for a structural edit of rows/columns, written from what Excel documents and does, NOT derived from this engine:
// - a reference to a deleted cell is #REF!;
// - a range keeps the rows/columns of it that survive (a deleted boundary moves to the nearest survivor) and is #REF! only when every
//   row/column of it is deleted;
// - inserting before the first row/column of a range shifts it, inserting after its last one leaves it alone, inserting inside it
//   (at any of its rows/columns after the first) extends it;
// - a whole-column reference only reacts to column edits and a whole-row reference to row edits;
// - references to another sheet only react to edits of that sheet; sheet names are case-insensitive.
// Each row runs on both paths and must give the same literal: a normal formula (stable ids, re-rendered) and a cached-only one (text).

import { WorkbookImpl } from '../workbook';
import { renderStableAst } from './stable-ast';

type Edit = { sheet: string; what: 'insRows' | 'delRows' | 'insCols' | 'delCols'; at: number; n: number };

// Excel numbering: rows from 1, columns by letter (`at` is the first row/column inserted before or deleted)
const COL = (letter: string): number => letter.charCodeAt(0) - 65 + 1;
const insRows = (at: number, n = 1, sheet = 'Main'): Edit => ({ sheet, what: 'insRows', at, n });
const delRows = (at: number, n = 1, sheet = 'Main'): Edit => ({ sheet, what: 'delRows', at, n });
const insCols = (at: string, n = 1, sheet = 'Main'): Edit => ({ sheet, what: 'insCols', at: COL(at), n });
const delCols = (at: string, n = 1, sheet = 'Main'): Edit => ({ sheet, what: 'delCols', at: COL(at), n });

// [formula, edit, what Excel shows afterwards]
const table: Array<[string, Edit, string]> = [
  // --- a range of rows, A2:A5 (rows 2 to 5)
  ['=SUM(A2:A5)', delRows(5), '=SUM(A2:A4)'], // the last row goes
  ['=SUM(A2:A5)', delRows(2), '=SUM(A2:A4)'], // the first row goes: A3:A5 moves up
  ['=SUM(A2:A5)', delRows(3), '=SUM(A2:A4)'], // an inner row
  ['=SUM(A2:A5)', delRows(4), '=SUM(A2:A4)'],
  ['=SUM(A2:A5)', delRows(2, 2), '=SUM(A2:A3)'], // from the start
  ['=SUM(A2:A5)', delRows(4, 3), '=SUM(A2:A3)'], // to past the end
  ['=SUM(A2:A5)', delRows(1, 2), '=SUM(A1:A3)'], // from before the start
  ['=SUM(A2:A5)', delRows(5, 2), '=SUM(A2:A4)'],
  ['=SUM(A2:A5)', delRows(3, 2), '=SUM(A2:A3)'], // the interior
  ['=SUM(A2:A5)', delRows(1), '=SUM(A1:A4)'], // above it
  ['=SUM(A2:A5)', delRows(6), '=SUM(A2:A5)'], // below it
  ['=SUM(A2:A5)', delRows(2, 4), '=SUM(#REF!)'], // all of it
  ['=SUM(A2:A5)', delRows(1, 6), '=SUM(#REF!)'], // all of it and more
  ['=SUM(A2:A5)', delRows(1, 5), '=SUM(#REF!)'],
  ['=SUM(A2:A5)', delRows(2, 5), '=SUM(#REF!)'],
  ['=SUM(A2:A2)', delRows(2), '=SUM(#REF!)'], // a one-cell range
  ['=SUM(A2:A5)', insRows(2), '=SUM(A3:A6)'], // before its first row: it moves down
  ['=SUM(A2:A5)', insRows(1), '=SUM(A3:A6)'],
  ['=SUM(A2:A5)', insRows(3), '=SUM(A2:A6)'], // inside: it grows
  ['=SUM(A2:A5)', insRows(5), '=SUM(A2:A6)'], // before its last row: it grows
  ['=SUM(A2:A5)', insRows(6), '=SUM(A2:A5)'], // after its last row: it stays
  ['=SUM(A2:A5)', insRows(9), '=SUM(A2:A5)'],
  ['=SUM(A2:A5)', insRows(4, 3), '=SUM(A2:A8)'],
  ['=SUM(A2:A5)', insRows(1, 3), '=SUM(A5:A8)'],
  // --- a single cell
  ['=A5', delRows(5), '=#REF!'],
  ['=A5', delRows(4), '=A4'],
  ['=A5', delRows(3, 3), '=#REF!'],
  ['=A5', delRows(6), '=A5'],
  ['=A5', insRows(5), '=A6'],
  ['=A5', insRows(6), '=A5'],
  ['=C5', delCols('C'), '=#REF!'],
  ['=C5', delCols('B'), '=B5'],
  ['=C5', insCols('C'), '=D5'],
  ['=C5', insCols('D'), '=C5'],
  // --- a two-dimensional range, B2:D5; each axis on its own
  ['=SUM(B2:D5)', delRows(5), '=SUM(B2:D4)'],
  ['=SUM(B2:D5)', delRows(2), '=SUM(B2:D4)'],
  ['=SUM(B2:D5)', delCols('D'), '=SUM(B2:C5)'],
  ['=SUM(B2:D5)', delCols('B'), '=SUM(B2:C5)'],
  ['=SUM(B2:D5)', delCols('C'), '=SUM(B2:C5)'],
  ['=SUM(B2:D5)', delCols('C', 2), '=SUM(B2:B5)'],
  ['=SUM(B2:D5)', delCols('B', 3), '=SUM(#REF!)'],
  ['=SUM(B2:D5)', delRows(2, 4), '=SUM(#REF!)'],
  ['=SUM(B2:D5)', insCols('B'), '=SUM(C2:E5)'],
  ['=SUM(B2:D5)', insCols('D'), '=SUM(B2:E5)'],
  ['=SUM(B2:D5)', insCols('E'), '=SUM(B2:D5)'],
  ['=SUM(B2:D5)', insRows(2), '=SUM(B3:D6)'],
  // --- absolute markers do not matter to an edit and stay
  ['=SUM($A$2:$A$5)', delRows(5), '=SUM($A$2:$A$4)'],
  ['=SUM($A$2:$A$5)', delRows(2), '=SUM($A$2:$A$4)'],
  ['=SUM($A2:A$5)', insRows(2), '=SUM($A3:A$6)'],
  ['=SUM($B$2:$D$5)', delCols('B'), '=SUM($B$2:$C$5)'],
  ['=$A$5', insRows(1), '=$A$6'],
  // --- the corners written the other way round
  ['=SUM(A5:A2)', delRows(5), '=SUM(A4:A2)'],
  ['=SUM(A5:A2)', delRows(2), '=SUM(A4:A2)'],
  ['=SUM(A5:A2)', delRows(2, 4), '=SUM(#REF!)'],
  ['=SUM(A5:A2)', insRows(2), '=SUM(A6:A3)'],
  ['=SUM(A5:A2)', insRows(6), '=SUM(A5:A2)'],
  ['=SUM(D5:B2)', delCols('B'), '=SUM(C5:B2)'],
  ['=SUM(D5:B2)', delCols('D'), '=SUM(C5:B2)'],
  ['=SUM(D5:B2)', insCols('C'), '=SUM(E5:B2)'],
  // --- whole columns
  ['=SUM(B:D)', delCols('B'), '=SUM(B:C)'],
  ['=SUM(B:D)', delCols('D'), '=SUM(B:C)'],
  ['=SUM(B:D)', delCols('C'), '=SUM(B:C)'],
  ['=SUM(B:D)', delCols('B', 3), '=SUM(#REF!)'],
  ['=SUM(B:D)', delCols('A', 2), '=SUM(A:B)'],
  ['=SUM(B:B)', delCols('B'), '=SUM(#REF!)'],
  ['=SUM(B:D)', insCols('B'), '=SUM(C:E)'],
  ['=SUM(B:D)', insCols('D'), '=SUM(B:E)'],
  ['=SUM(B:D)', insCols('E'), '=SUM(B:D)'],
  ['=SUM($B:$D)', delCols('C'), '=SUM($B:$C)'],
  ['=SUM(D:B)', delCols('B'), '=SUM(C:B)'],
  ['=SUM(B:D)', insRows(1), '=SUM(B:D)'], // rows do not touch a whole column
  ['=SUM(B:D)', delRows(2, 3), '=SUM(B:D)'],
  // --- whole rows
  ['=SUM(2:4)', delRows(2), '=SUM(2:3)'],
  ['=SUM(2:4)', delRows(4), '=SUM(2:3)'],
  ['=SUM(2:4)', delRows(3), '=SUM(2:3)'],
  ['=SUM(2:4)', delRows(2, 3), '=SUM(#REF!)'],
  ['=SUM(2:4)', delRows(1, 5), '=SUM(#REF!)'],
  ['=SUM(3:3)', delRows(3), '=SUM(#REF!)'],
  ['=SUM(2:4)', insRows(2), '=SUM(3:5)'],
  ['=SUM(2:4)', insRows(4), '=SUM(2:5)'],
  ['=SUM(2:4)', insRows(5), '=SUM(2:4)'],
  ['=SUM($2:$4)', delRows(2), '=SUM($2:$3)'],
  ['=SUM(4:2)', delRows(4), '=SUM(3:2)'],
  ['=SUM(2:4)', insCols('A'), '=SUM(2:4)'], // columns do not touch a whole row
  ['=SUM(2:4)', delCols('A', 3), '=SUM(2:4)'],
  // --- other sheets: only an edit of the referenced sheet reaches a reference
  ['=SUM(Sheet2!A2:A5)', delRows(5, 1, 'Sheet2'), '=SUM(Sheet2!A2:A4)'],
  ['=SUM(Sheet2!A2:A5)', delRows(5, 1, 'Main'), '=SUM(Sheet2!A2:A5)'],
  ['=SUM(Sheet2!A2:A5)', insRows(2, 1, 'Sheet2'), '=SUM(Sheet2!A3:A6)'],
  ['=SUM(A2:A5)', delRows(5, 1, 'Sheet2'), '=SUM(A2:A5)'], // unprefixed: the formula's own sheet
  ["=SUM('My Sheet'!A2:A5)", delRows(2, 1, 'My Sheet'), "=SUM('My Sheet'!A2:A4)"],
  ['=Sheet2!A5', delRows(5, 1, 'Sheet2'), '=#REF!'],
  ['=SUM(Sheet2!B:D)', delCols('B', 1, 'Sheet2'), '=SUM(Sheet2!B:C)'],
  ['=SUM(Sheet2!2:4)', delRows(2, 1, 'Sheet2'), '=SUM(Sheet2!2:3)'],
  // --- sheet names are case-insensitive
  ['=SUM(sheet2!A2:A5)', delRows(5, 1, 'Sheet2'), '=SUM(sheet2!A2:A4)'],
  ['=SUM(SHEET2!A2:A5)', insRows(2, 1, 'Sheet2'), '=SUM(SHEET2!A3:A6)'],
  ['=sheet2!A1', insRows(1, 1, 'Sheet2'), '=sheet2!A2'],
  ["=SUM('my sheet'!A2:A5)", delRows(5, 1, 'My Sheet'), "=SUM('my sheet'!A2:A4)"],
  ['=SUM(sheet2!A2:A5)', delRows(5, 1, 'Main'), '=SUM(sheet2!A2:A5)'],
  // --- several references in a formula
  ['=SUM(A2:A5)+A5+B:B+3:3', delRows(5), '=SUM(A2:A4)+#REF!+B:B+3:3'],
  ['=A1:A3+SUM(A2:A5)', insRows(3), '=A1:A4+SUM(A2:A6)'],
];

const NAMES = ['Main', 'Sheet2', 'My Sheet'];

function build(text: string, cachedOnly: boolean): WorkbookImpl {
  const wb = new WorkbookImpl('wb', 'WB');
  wb.renameSheet(wb.getSheet().id, 'Main');
  wb.addSheet('Sheet2');
  wb.addSheet('My Sheet');
  const main = wb.getSheetIdByName('Main');
  if (cachedOnly) wb.setCell(main, 29, 29, { formula: text, formulaAst: undefined, value: 7, cachedOnly: true });
  else wb.setFormula(main, 29, 29, text);
  return wb;
}

function apply(wb: WorkbookImpl, e: Edit): void {
  const sheet = wb.getSheet(wb.getSheetIdByName(e.sheet));
  const at = e.at - 1; // 0-based
  if (e.what === 'insRows') sheet.insertRows(at, e.n);
  else if (e.what === 'delRows') sheet.deleteRows(at, e.n);
  else if (e.what === 'insCols') sheet.insertCols(at, e.n);
  else sheet.deleteCols(at, e.n);
}

/** The text of the one formula cell (a normal formula is rendered from its stable AST right now). */
function textOf(wb: WorkbookImpl): string {
  const resolver = { getSheet: (name?: string) => (name && wb.getSheetIdByName(name) ? wb.getSheet(wb.getSheetIdByName(name)) : undefined) };
  for (const name of NAMES) {
    const sheet = wb.getSheet(wb.getSheetIdByName(name));
    for (const [, , c] of sheet.entries()) {
      if (c.formulaAst) return renderStableAst(c.formulaAst, resolver, sheet);
      if (c.formula) return c.formula;
    }
  }
  throw new Error('no formula cell');
}

describe("a structural edit, by Excel's documented behaviour", () => {
  it.each(table)('%s under %j is %s on both paths', (formula, edit, expected) => {
    for (const cachedOnly of [false, true]) {
      const wb = build(formula, cachedOnly);
      apply(wb, edit);
      expect({ cachedOnly, text: textOf(wb) }).toEqual({ cachedOnly, text: expected });
    }
  });
});
