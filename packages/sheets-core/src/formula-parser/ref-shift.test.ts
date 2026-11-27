import { shiftFormulaRefs, shiftIndex, type StructuralEdit } from './ref-shift';
import { MAX_SHEET_COLS, MAX_SHEET_ROWS } from './types';

const rows = (kind: 'insert' | 'delete', start: number, count = 1): StructuralEdit => ({ axis: 'row', kind, start, count });
const cols = (kind: 'insert' | 'delete', start: number, count = 1): StructuralEdit => ({ axis: 'col', kind, start, count });

/** The formula sits on the edited sheet "Main"; "Other" is another sheet. */
const onMain = (formula: string, edit: StructuralEdit): string => shiftFormulaRefs(formula, edit, true, (n) => n === 'Main');
/** The formula sits on another sheet than the edited "Main". */
const onOther = (formula: string, edit: StructuralEdit): string => shiftFormulaRefs(formula, edit, false, (n) => n === 'Main');

describe('shiftIndex', () => {
  it('moves what is at or after an insert, and what is after a delete; deleted rows have no position', () => {
    expect(shiftIndex(2, rows('insert', 3, 2))).toBe(2);
    expect(shiftIndex(3, rows('insert', 3, 2))).toBe(5);
    expect(shiftIndex(9, rows('insert', 3, 2))).toBe(11);
    expect(shiftIndex(2, rows('delete', 3, 2))).toBe(2);
    expect(shiftIndex(3, rows('delete', 3, 2))).toBeUndefined();
    expect(shiftIndex(4, rows('delete', 3, 2))).toBeUndefined();
    expect(shiftIndex(5, rows('delete', 3, 2))).toBe(3);
  });

  it('pushes what an insert moves past the last row or column off the sheet', () => {
    expect(shiftIndex(MAX_SHEET_ROWS - 1, rows('insert', 0, 1))).toBeUndefined();
    expect(shiftIndex(MAX_SHEET_ROWS - 2, rows('insert', 0, 1))).toBe(MAX_SHEET_ROWS - 1);
    expect(shiftIndex(MAX_SHEET_COLS - 1, cols('insert', 0, 1))).toBeUndefined();
    expect(shiftIndex(MAX_SHEET_COLS - 2, cols('insert', 0, 1))).toBe(MAX_SHEET_COLS - 1);
  });
});

describe('shiftFormulaRefs: oracle table (the outcome of a normal formula on the edited sheet)', () => {
  // [formula, edit, expected]
  const table: Array<[string, StructuralEdit, string]> = [
    // single cells
    ['=A5', rows('insert', 0, 2), '=A7'],
    ['=A5', rows('insert', 4), '=A6'],
    ['=A5', rows('insert', 5), '=A5'],
    ['=A5', rows('delete', 1), '=A4'],
    ['=A5', rows('delete', 4), '=#REF!'],
    ['=A5', rows('delete', 3, 3), '=#REF!'],
    ['=A5', rows('delete', 5), '=A5'],
    ['=A5', cols('insert', 0), '=B5'],
    ['=C5', cols('delete', 0, 2), '=A5'],
    ['=C5', cols('delete', 2), '=#REF!'],
    ['=C5', cols('insert', 3), '=C5'],
    // absolute markers and case are kept
    ['=$A$5', rows('insert', 0), '=$A$6'],
    ['=$A5', cols('insert', 0), '=$B5'],
    ['=A$5', rows('insert', 0), '=A$6'],
    ['=a5', cols('insert', 0), '=b5'],
    ['=aB5', cols('insert', 0, 3), '=AE5'],
    ['=ab5+Ab5', cols('insert', 0, 3), '=ae5+AE5'],
    ['=$a$5', rows('delete', 0), '=$a$4'],
    ['=A05', rows('insert', 0), '=A6'],
    ['=A05', rows('insert', 9), '=A05'],
    ['=XFD1', cols('insert', 0), '=#REF!'],
    ['=A1048576', rows('insert', 0), '=#REF!'],
    // ranges: both corners move; a deleted corner makes the whole range #REF! (as the engine does)
    ['=SUM(A1:A5)', rows('insert', 2), '=SUM(A1:A6)'],
    ['=SUM(A1:A5)', rows('insert', 0), '=SUM(A2:A6)'],
    ['=SUM(A1:A5)', rows('insert', 5), '=SUM(A1:A5)'],
    ['=SUM(A1:A5)', rows('delete', 1), '=SUM(A1:A4)'],
    ['=SUM(A1:A5)', rows('delete', 4), '=SUM(#REF!)'],
    ['=SUM(A1:A5)', rows('delete', 0), '=SUM(#REF!)'],
    ['=SUM(B2:C3)', cols('delete', 0), '=SUM(A2:B3)'],
    ['=SUM(B2:C3)', cols('delete', 1, 2), '=SUM(#REF!)'],
    ['=SUM($B$2:$C$3)', cols('insert', 1), '=SUM($C$2:$D$3)'],
    ['=SUM(B2 : C3)', rows('insert', 0), '=SUM(B3 : C4)'],
    ['=SUM(C3:B2)', rows('insert', 2), '=SUM(C4:B2)'],
    // whole columns and rows react to their own axis only
    ['=SUM(A:A)', rows('insert', 0), '=SUM(A:A)'],
    ['=SUM(A:A)', rows('delete', 0), '=SUM(A:A)'],
    ['=SUM(B:C)', cols('insert', 0), '=SUM(C:D)'],
    ['=SUM(B:C)', cols('delete', 0), '=SUM(A:B)'],
    ['=SUM($B:$C)', cols('insert', 2), '=SUM($B:$D)'],
    ['=SUM(b:c)', cols('insert', 0), '=SUM(c:d)'],
    ['=SUM(B:C)', cols('delete', 1), '=SUM(#REF!)'],
    ['=SUM(B:C)', cols('delete', 2), '=SUM(#REF!)'],
    ['=SUM(1:1)', rows('insert', 0), '=SUM(2:2)'],
    ['=SUM(2:3)', rows('insert', 2), '=SUM(2:4)'],
    ['=SUM($2:$3)', rows('delete', 0), '=SUM($1:$2)'],
    ['=SUM(2:3)', rows('delete', 1), '=SUM(#REF!)'],
    ['=SUM(2:3)', cols('insert', 0), '=SUM(2:3)'],
    ['=SUM(2:3)', cols('delete', 0), '=SUM(2:3)'],
    // spill references
    ['=A5#', rows('insert', 0), '=A6#'],
    ['=A5#', rows('delete', 4), '=#REF!#'],
    ['=SUM(A5#:B5)', rows('insert', 0), '=SUM(A6#:B6)'],
    // prefixes of the edited sheet
    ['=Main!A5', rows('insert', 0), '=Main!A6'],
    ["='Main'!A5", rows('insert', 0), "='Main'!A6"],
    ['=Main!A5', rows('delete', 4), '=#REF!'],
    ['=Main!A1:B5', rows('insert', 1), '=Main!A1:B6'],
    ['=Main!A1:B5', rows('delete', 4), '=#REF!'],
    ['=Main!A1:Main!B5', rows('insert', 1), '=Main!A1:Main!B6'],
    ['=Main!A1:Main!B5', rows('delete', 0), '=#REF!'],
    ['=Main!A:B', cols('insert', 0), '=Main!B:C'],
    ['=Main!A:B', cols('delete', 0), '=#REF!'],
    ['=Main!1:2', rows('insert', 0), '=Main!2:3'],
    ['=Main!A5#', rows('insert', 0), '=Main!A6#'],
    ['=Main!A1:B2 + A3', rows('insert', 0), '=Main!A2:B3 + A4'],
    // other sheets, the 3D spans and external workbooks are not the edited sheet
    ['=Other!A5', rows('insert', 0), '=Other!A5'],
    ["='My Sheet'!A5:B6", rows('insert', 0), "='My Sheet'!A5:B6"],
    ['=Other!A:A', cols('insert', 0), '=Other!A:A'],
    ['=Other!A5 + A5', rows('insert', 0), '=Other!A5 + A6'],
    ['=A5 + Other!A5', rows('insert', 0), '=A6 + Other!A5'],
    ['=SUM(Main:Other!A5)', rows('insert', 0), '=SUM(Main:Other!A5)'],
    ["=SUM('Main:Other'!A1:B2)", rows('insert', 0), "=SUM('Main:Other'!A1:B2)"],
    ['=SUM(Main:Other!A1:A5, A1)', rows('insert', 0), '=SUM(Main:Other!A1:A5, A2)'],
    ['=SUM(Main:Other!A1:Other!A5)', rows('insert', 0), '=SUM(Main:Other!A1:Other!A5)'],
    ['=[1]Main!A5', rows('insert', 0), '=[1]Main!A5'],
    ['=[Book.xlsx]Main!A1:B2', rows('insert', 0), '=[Book.xlsx]Main!A1:B2'],
    ["='[1]Main'!A5", rows('insert', 0), "='[1]Main'!A5"],
    ['=[1]Main!A5+A5', rows('insert', 0), '=[1]Main!A5+A6'],
    ['=Ghost!A5', rows('insert', 0), '=Ghost!A5'],
    // A1:B2 pairing: a different right sheet is two references; a chain pairs only its first two
    ['=A1:Other!B2', rows('insert', 0), '=A2:Other!B2'],
    ['=Other!A1:B2:C3', rows('insert', 0), '=Other!A1:B2:C4'],
    ['=A1:B2:C3', rows('insert', 0), '=A2:B3:C4'],
    ['=A1:B2:C3', rows('delete', 0), '=#REF!:C2'],
    ['=A1:B2:C3', rows('delete', 2), '=A1:B2:#REF!'],
    ['=INDEX(A1:A5,1):B2', rows('insert', 0), '=INDEX(A2:A6,1):B3'],
    ['=INDEX(A1:A5,1):B2:C3', rows('insert', 0), '=INDEX(A2:A6,1):B3:C4'],
    ['=A1:INDEX(A1:A5,1)', rows('delete', 0), '=#REF!:INDEX(#REF!,1)'],
    ['=A1:B2#', rows('delete', 1), '=A1:#REF!#'],
    ['=A1 B1:C1', rows('insert', 0), '=A2 B2:C2'],
    ['=@A1:B2', rows('insert', 0), '=@A2:B3'],
    ['=(A1,B2)', rows('delete', 0), '=(#REF!,B1)'],
    // not a reference: strings, names, structured references, functions, numbers
    ['=A5&"A5 and B7"&\'A5\'', rows('insert', 0), '=A6&"A5 and B7"&\'A5\''],
    ['=LOG10(A5)+ATAN2(1,2)', rows('insert', 0), '=LOG10(A6)+ATAN2(1,2)'],
    ['=MyName+A5', rows('insert', 0), '=MyName+A6'],
    ['=Table1[Col A]+A5', rows('insert', 0), '=Table1[Col A]+A6'],
    ['=Table1[[#This Row],[B7]]+A5', rows('insert', 0), '=Table1[[#This Row],[B7]]+A6'],
    ["=Table1[Col'[A1']]+A5", rows('insert', 0), "=Table1[Col'[A1']]+A6"],
    ['=Sheet2!Table1[B7]+A5', rows('insert', 0), '=Sheet2!Table1[B7]+A6'],
    ['={1,2;3,4}+A5', rows('insert', 0), '={1,2;3,4}+A6'],
    ['=#REF!+A5', rows('insert', 0), '=#REF!+A6'],
    ['=A1B5+A5', rows('insert', 0), '=A1B5+A6'],
    ['=A5.5+A5', rows('insert', 0), '=A5.5+A6'],
    ['=1.5E3+A5', rows('insert', 0), '=1.5E3+A6'],
    // whitespace is kept exactly
    ['=  SUM( A1 ,\tB2 ) \n', rows('insert', 0), '=  SUM( A2 ,\tB3 ) \n'],
    ['=SUM(A1:B2)  +  Other!A1', rows('insert', 0), '=SUM(A2:B3)  +  Other!A1'],
    // text that is not a formula, or has no references
    ['A5', rows('insert', 0), 'A6'],
    ['', rows('insert', 0), ''],
    ['=', rows('insert', 0), '='],
    ['=1+2', rows('insert', 0), '=1+2'],
  ];

  it.each(table)('%s under %j gives %s', (formula, edit, expected) => {
    expect(onMain(formula, edit)).toBe(expected);
  });

  it('returns the very same string when nothing changes', () => {
    const formula = '=SUM(A1:A3)+Other!B2';
    expect(onMain(formula, rows('insert', 10))).toBe(formula);
  });
});

describe('shiftFormulaRefs: the formula is on another sheet than the edited one', () => {
  it('leaves unprefixed references alone and shifts the ones that name the edited sheet', () => {
    expect(onOther('=A5+B2:C3+1:2+A:B', rows('insert', 0))).toBe('=A5+B2:C3+1:2+A:B');
    expect(onOther('=Main!A5+A5+Main!B2:C3', rows('insert', 0))).toBe('=Main!A6+A5+Main!B3:C4');
    expect(onOther('=Main!A:B', cols('delete', 0))).toBe('=#REF!');
    expect(onOther("=SUM('Main'!A1:A5)", rows('delete', 1))).toBe("=SUM('Main'!A1:A4)");
    expect(onOther('=Main!A5#', rows('delete', 4))).toBe('=#REF!#');
  });
});

describe('shiftFormulaRefs: text the grammar does not accept', () => {
  it('still shifts the references it can recognise and keeps every other byte', () => {
    expect(onMain('=A1+++', rows('insert', 0))).toBe('=A2+++');
    expect(onMain('=SUM(A1:', rows('insert', 0))).toBe('=SUM(A2:');
    expect(onMain('=A1 ) ( B2 @@ C3', rows('insert', 0))).toBe('=A2 ) ( B3 @@ C4');
    expect(onMain('=A1&"unterminated B2', rows('insert', 0))).toBe('=A2&"unterminated B2');
    expect(onMain("=A1&'unterminated B2", rows('insert', 0))).toBe("=A2&'unterminated B2");
    expect(onMain('=A1+Table1[unterminated B2', rows('insert', 0))).toBe('=A2+Table1[unterminated B2');
    expect(onMain('=A1+#BOGUS+B2', rows('insert', 0))).toBe('=A2+#BOGUS+B3');
    expect(onMain("=A1+''!B2+C3", rows('insert', 0))).toBe("=A2+''!B3+C4");
    expect(onMain('=A1+1:99999999+B2', rows('insert', 0))).toBe('=A2+1:99999999+B3');
    expect(onMain('=A1+AAAA:AAAA+B2+A:XFE', rows('insert', 0))).toBe('=A2+AAAA:AAAA+B3+A:XFE');
    expect(onMain('=A1+1e999+B2+.+$+B3', rows('insert', 0))).toBe('=A2+1e999+B3+.+$+B4');
    expect(onMain('=A1+€+B2~~B3', rows('insert', 0))).toBe('=A2+€+B3~~B4');
    expect(onMain('=A1+$x+B2', rows('insert', 0))).toBe('=A2+$x+B3');
  });
});

describe('shiftFormulaRefs: hostile text', () => {
  it('handles a formula that is only brackets, quotes or spaces', () => {
    for (const text of ['=[[[[[[', '=]]]]]]', "=''''''''", '="""""""', '=\'[\'[\'[', '=! ! ! !', '=A1!B2!C3!', '= \t\n\r ']) {
      expect(() => onMain(text, rows('delete', 0))).not.toThrow();
    }
    expect(onMain('=A1!B2', rows('insert', 0))).toBe('=A1!B2'); // a sheet named A1
  });

  it('is not confused by the words of a long run (no sheet prefix is searched for through all of it)', () => {
    const words = `=${'ab '.repeat(2700)}A1`;
    expect(onMain(words, rows('insert', 0))).toBe(`=${'ab '.repeat(2700)}A2`);
    const spaced = `=${'ab '.repeat(300)}!A1`;
    expect(() => onMain(spaced, rows('insert', 0))).not.toThrow();
  });
});

describe('shiftFormulaRefs: a formula without a leading =', () => {
  it('is read as the formula it is', () => {
    expect(onMain('SUM(A1:B2)', rows('insert', 0))).toBe('SUM(A2:B3)');
  });
});
