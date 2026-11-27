import { shiftFormula } from './shared-formula';

// [formula as the master wrote it, rows, cols, the formula as that many cells away reads]
const ORACLE: Array<[string, number, number, string]> = [
  ['A2*2', 1, 0, 'A3*2'],
  ['A2*2', 0, 1, 'B2*2'],
  ['A2*2', 3, 2, 'C5*2'],
  ['$A2*2', 1, 1, '$A3*2'],
  ['A$2*2', 1, 1, 'B$2*2'],
  ['$A$2*2', 5, 5, '$A$2*2'],
  ['A2+B2', -1, 0, 'A1+B1'],
  ['SUM(A1:B2)', 1, 1, 'SUM(B2:C3)'],
  ['SUM($A$1:B2)', 1, 1, 'SUM($A$1:C3)'],
  ['SUM(A:A)', 0, 2, 'SUM(C:C)'],
  ['SUM(A:A)', 4, 0, 'SUM(A:A)'],
  ['SUM($A:B)', 0, 1, 'SUM($A:C)'],
  ['SUM(1:1)', 2, 0, 'SUM(3:3)'],
  ['SUM($1:2)', 2, 3, 'SUM($1:4)'],
  ['SUM(A1:A3)/COUNT(A1:A3)', 1, 0, 'SUM(A2:A4)/COUNT(A2:A4)'],
  ['Sheet2!A1+1', 1, 0, 'Sheet2!A2+1'],
  ['Q1!A1+Q1', 1, 0, 'Q1!A2+Q2'], // the sheet is named Q1; the second Q1 is a cell
  ["'My Sheet'!B2+1", 1, 1, "'My Sheet'!C3+1"],
  ["'It''s'!B2", 1, 0, "'It''s'!B3"],
  ['Sheet1:Sheet3!A1', 1, 0, 'Sheet1:Sheet3!A2'],
  ['"A1"&A1', 1, 0, '"A1"&A2'],
  ['"say ""A1"""&A1', 1, 0, '"say ""A1"""&A2'],
  ['LOG10(A1)', 1, 0, 'LOG10(A2)'],
  ['LOG10(100)+ATAN2(A1,B1)', 1, 0, 'LOG10(100)+ATAN2(A2,B2)'],
  ['Rate_2024*A1', 1, 0, 'Rate_2024*A2'],
  ['A1B*A1', 1, 0, 'A1B*A2'],
  ['_xlfn.X1+A1', 1, 0, '_xlfn.X1+A2'],
  ['name.A1+A1', 1, 0, 'name.A1+A2'],
  ['XFD1+A1', 0, 1, '#REF!+B1'],
  ['A1', -1, 0, '#REF!'],
  ['A1', 0, -1, '#REF!'],
  ['A1048576', 1, 0, '#REF!'],
  ['SUM(A:A)', 0, -1, 'SUM(#REF!:#REF!)'],
  ['SUM(XFD:XFD)', 0, 1, 'SUM(#REF!:#REF!)'],
  ['SUM(2:2)', -2, 0, 'SUM(#REF!:#REF!)'],
  ['SUM(1048576:1048576)', 1, 0, 'SUM(#REF!:#REF!)'],
  ['Table1[[#This Row],[A1]]+A1', 1, 0, 'Table1[[#This Row],[A1]]+A2'],
  ['Table1[Col [x]]', 1, 0, 'Table1[Col [x]]'],
  ['[1]Sheet1!A1+A1', 1, 0, '[1]Sheet1!A2+A2'],
  ['IF(A1>0,#REF!,#N/A)+A1', 1, 0, 'IF(A2>0,#REF!,#N/A)+A2'],
  ['#DIV/0!+A1', 1, 0, '#DIV/0!+A2'],
  ['#+A1', 1, 0, '#+A2'],
  ['1.5+A1', 1, 0, '1.5+A2'],
  ['1E5+A1', 1, 0, '1E5+A2'],
  ['ZZZ1+A1', 1, 0, 'ZZZ1+A2'], // ZZZ is past XFD: a name
  ['SUM(ZZZ:ZZZ)+A1', 1, 0, 'SUM(ZZZ:ZZZ)+A2'],
  ['A1*B', 1, 1, 'B2*B'],
  ['TRUE+A1', 1, 0, 'TRUE+A2'],
  ['$+A1', 1, 0, '$+A2'],
  ['A1:B', 1, 0, 'A2:B'],
  ['a1+B2', 1, 1, 'B2+C3'],
  ['_xlfn.SEQUENCE(3)+A1#', 1, 0, '_xlfn.SEQUENCE(3)+A2#'],
  ['A1&"unterminated', 1, 0, 'A2&"unterminated'],
  ['A1&[unterminated', 1, 0, 'A2&[unterminated'],
  // names with letters and digits of any script are names, not cells
  ['Größe1+A1', 1, 0, 'Größe1+A2'],
  ['Größe1+A1', 0, 1, 'Größe1+B1'],
  ['Налог2*A1', 1, 0, 'Налог2*A2'],
  ['税率3+B2', 1, 1, '税率3+C3'],
  ['A1+Größe', 1, 0, 'A2+Größe'],
  ['ÄA1+A1', 1, 0, 'ÄA1+A2'], // a letter before a would-be cell glues to it
  ['A1Ä+A1', 1, 0, 'A1Ä+A2'], // and so does one after it
  ['SUM(Ünit1:Ünit9)+A1', 1, 0, 'SUM(Ünit1:Ünit9)+A2'],
  ['Größe1(A1)+A1', 1, 0, 'Größe1(A2)+A2'],
  ['Größe1!A1+A1', 1, 0, 'Größe1!A2+A2'], // a sheet named Größe1
  ['x١2+A1', 1, 0, 'x١2+A2'], // a non-ASCII digit is a name character too
  ['𝒜1+A1', 1, 0, '𝒜1+A2'], // a letter outside the BMP (two UTF-16 units)
  ['A1+𝒜A1', 1, 0, 'A2+𝒜A1'],
  ['A1+"Größe1"', 1, 0, 'A2+"Größe1"'],
];

describe('shiftFormula', () => {
  it.each(ORACLE)('%s moved %i down and %i right reads %s', (formula, rows, cols, expected) => {
    expect(shiftFormula(formula, rows, cols)).toBe(expected);
  });

  it('is the identity for no move', () => {
    expect(shiftFormula('A1+"x"+Sheet!B2', 0, 0)).toBe('A1+"x"+Sheet!B2');
  });

  it('is its own inverse for in-range moves (a lower-case reference comes back upper-case)', () => {
    for (const [formula] of ORACLE.filter(([f, , , e]) => !e.includes('#REF!') && f !== 'a1+B2')) {
      expect(shiftFormula(shiftFormula(formula, 2, 3), -2, -3)).toBe(formula);
    }
  });
});
