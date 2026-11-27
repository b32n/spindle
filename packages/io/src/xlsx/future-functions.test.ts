import { FUTURE_FUNCTIONS, fromFileFormula, toFileFormula } from './future-functions';

// [as the file has it, as the editor shows it]
const PAIRS: Array<[string, string]> = [
  // the prefixed functions of the corpus probe (fixtures/xlsx formulas-core, formulas-modern, formula-syntax, xlsxwriter-rich)
  ['_xlfn.MAXIFS(B2:B6,A2:A6,"?o*")', 'MAXIFS(B2:B6,A2:A6,"?o*")'],
  ['_xlfn.AGGREGATE(9,6,B2:B6)', 'AGGREGATE(9,6,B2:B6)'],
  ['_xlfn.IFS(B2>5,"big",TRUE,"small")', 'IFS(B2>5,"big",TRUE,"small")'],
  ['_xlfn.SWITCH(A1,1,"one",2,"two")', 'SWITCH(A1,1,"one",2,"two")'],
  ['_xlfn.XOR(TRUE,FALSE)', 'XOR(TRUE,FALSE)'],
  ['_xlfn.CONCAT(A2,"-",A3)', 'CONCAT(A2,"-",A3)'],
  ['_xlfn.TEXTJOIN(",",TRUE,A2:A4)', 'TEXTJOIN(",",TRUE,A2:A4)'],
  ['_xlfn.NUMBERVALUE("1,5",",")', 'NUMBERVALUE("1,5",",")'],
  ['_xlfn.TEXTBEFORE("a-b","-")', 'TEXTBEFORE("a-b","-")'],
  ['_xlfn.TEXTSPLIT("a-b","-")', 'TEXTSPLIT("a-b","-")'],
  ['_xlfn.DAYS(B1,A1)', 'DAYS(B1,A1)'],
  ['_xlfn.ISOWEEKNUM(A1)', 'ISOWEEKNUM(A1)'],
  ['_xlfn.STDEV.S(B2:B6)', 'STDEV.S(B2:B6)'],
  ['_xlfn.VAR.S(B2:B6)', 'VAR.S(B2:B6)'],
  ['_xlfn.PERCENTILE.INC(B2:B6,0.9)', 'PERCENTILE.INC(B2:B6,0.9)'],
  ['_xlfn.RANK.EQ(B2,B2:B6)', 'RANK.EQ(B2,B2:B6)'],
  ['_xlfn.XLOOKUP("East",A2:A6,B2:B6)', 'XLOOKUP("East",A2:A6,B2:B6)'],
  ['_xlfn.XMATCH("East",A2:A6)', 'XMATCH("East",A2:A6)'],
  ['_xlfn.SEQUENCE(3)', 'SEQUENCE(3)'],
  ['_xlfn._xlws.FILTER(B2:B6,B2:B6>4)', 'FILTER(B2:B6,B2:B6>4)'],
  ['_xlfn._xlws.SORT(B2:B6)', 'SORT(B2:B6)'],
  ['_xlfn.SORTBY(B2:B6,C2:C6)', 'SORTBY(B2:B6,C2:C6)'],
  // nesting, and a prefixed function inside a plain one
  ['SUM(_xlfn.IFS(A1>1,2,TRUE,3),_xlfn.DAYS(B1,A1))', 'SUM(IFS(A1>1,2,TRUE,3),DAYS(B1,A1))'],
  // LET and LAMBDA parameters
  ['_xlfn.LET(_xlpm.x,B2,_xlpm.y,B3,_xlpm.x+_xlpm.y)', 'LET(x,B2,y,B3,x+y)'],
  ['_xlfn.LAMBDA(_xlpm.n,_xlpm.n*2)(B2)', 'LAMBDA(n,n*2)(B2)'],
  // implicit intersection and the spill reference
  ['_xlfn.SINGLE(A1:A3)', '@A1:A3'],
  ["_xlfn.SINGLE('My sheet'!A1:A3)+1", "@'My sheet'!A1:A3+1"],
  ['_xlfn.SINGLE(INDEX(A1:A3,2))', '@INDEX(A1:A3,2)'],
  ['SUM(_xlfn.ANCHORARRAY(A1))', 'SUM(A1#)'],
  ['_xlfn.ANCHORARRAY(Sheet2!$B$4)*2', 'Sheet2!$B$4#*2'],
  ["_xlfn.ANCHORARRAY('My sheet'!B4)", "'My sheet'!B4#"],
];

describe('fromFileFormula / toFileFormula', () => {
  it.each(PAIRS.filter(([, shown]) => !shown.includes('_xlfn')))('reads %s as %s and writes it back', (file, shown) => {
    expect(fromFileFormula(file)).toBe(shown);
    expect(toFileFormula(shown)).toBe(file);
  });

  it('leaves a prefix that has no meaning for the call forms alone (an operand that is not a reference)', () => {
    // SINGLE of an expression and ANCHORARRAY of a range have no short form: the prefix goes, the call stays
    expect(fromFileFormula('_xlfn.SINGLE(A1+B1)')).toBe('SINGLE(A1+B1)');
    expect(fromFileFormula('_xlfn.SINGLE(A1 B1)')).toBe('SINGLE(A1 B1)');
    expect(fromFileFormula('_xlfn.ANCHORARRAY(A1:A3)')).toBe('ANCHORARRAY(A1:A3)');
    expect(fromFileFormula('_xlfn.ANCHORARRAY(A1,B1)')).toBe('ANCHORARRAY(A1,B1)');
    expect(fromFileFormula('_xlfn.SINGLE()')).toBe('SINGLE()');
    expect(fromFileFormula('_xlfn.SINGLE(A1')).toBe('SINGLE(A1');
    expect(fromFileFormula('_xlfn.ANCHORARRAY(A1')).toBe('ANCHORARRAY(A1');
    expect(fromFileFormula('SUM(_xlfn.ANCHORARRAY(_xlfn.SINGLE(A1:A2)))')).toBe('SUM(ANCHORARRAY(@A1:A2))');
    // without the _xlfn. prefix a SINGLE call is not the call form
    expect(fromFileFormula('SINGLE(A1)')).toBe('SINGLE(A1)');
  });

  it('never rewrites string literals, quoted sheet names or bracket groups', () => {
    const text = '"_xlfn.IFS(@A1 B1#)"&\'_xlfn.IFS\'!A1&Table1[[#This Row],[_xlfn.X]]&"it""s _xlfn.A("';
    expect(fromFileFormula(text)).toBe(text);
    expect(toFileFormula('"IFS(@A1)"&\'IFS(\'!A1&Table1[@[IFS(]]')).toBe('"IFS(@A1)"&\'IFS(\'!A1&Table1[@[IFS(]]');
    // an unterminated string or bracket runs to the end and is left alone
    expect(fromFileFormula('A1&"_xlfn.IFS(')).toBe('A1&"_xlfn.IFS(');
    expect(fromFileFormula('A1&T[_xlfn.IFS(')).toBe('A1&T[_xlfn.IFS(');
  });

  it('keeps case, spaces, numbers, error literals and defined names as they are', () => {
    expect(fromFileFormula('_XLFN.ifs( A1>1 , 2 )')).toBe('ifs( A1>1 , 2 )');
    expect(toFileFormula('ifs( A1>1 , 2 )')).toBe('_xlfn.ifs( A1>1 , 2 )');
    expect(toFileFormula('1.5E+3+A1+#REF!+IFERROR(A1,#N/A)')).toBe('1.5E+3+A1+#REF!+IFERROR(A1,#N/A)');
    expect(toFileFormula('A1#REF!')).toBe('A1#REF!');
    expect(toFileFormula('_xlnm.Print_Area+Größe')).toBe('_xlnm.Print_Area+Größe');
    expect(fromFileFormula('_xlfn.')).toBe('_xlfn.');
    expect(fromFileFormula('_xlfn.(1)')).toBe('_xlfn.(1)');
    expect(fromFileFormula('_xlpm.')).toBe('_xlpm.');
    expect(fromFileFormula('')).toBe('');
  });

  it('keeps runs of whitespace between tokens', () => {
    expect(toFileFormula('IFS(A1,\t 1,\n\n 2)')).toBe('_xlfn.IFS(A1,\t 1,\n\n 2)');
    expect(fromFileFormula('_xlfn.IFS(A1,   1)')).toBe('IFS(A1,   1)');
  });

  it('does not touch a function the file format needs no prefix for, or one that is already prefixed', () => {
    expect(toFileFormula('SUM(A1:A3)+COUNTIFS(A:A,1)+IFERROR(1,2)')).toBe('SUM(A1:A3)+COUNTIFS(A:A,1)+IFERROR(1,2)');
    expect(toFileFormula('_xlfn.IFS(A1,1)')).toBe('_xlfn.IFS(A1,1)');
    expect(toFileFormula('_xlfn._xlws.SORT(A1:A3)')).toBe('_xlfn._xlws.SORT(A1:A3)');
  });

  it('writes the implicit-intersection operator as a call, whatever it applies to', () => {
    expect(toFileFormula('@A1:A3')).toBe('_xlfn.SINGLE(A1:A3)');
    expect(toFileFormula('@Table1[Col]+1')).toBe('_xlfn.SINGLE(Table1[Col])+1');
    expect(toFileFormula('@A1:INDEX(B:B,2)')).toBe('_xlfn.SINGLE(A1:INDEX(B:B,2))');
    expect(toFileFormula('@Sheet1!A:A')).toBe('_xlfn.SINGLE(Sheet1!A:A)');
    expect(toFileFormula('@IFS(@A1:A3,1,TRUE,2)')).toBe('_xlfn.SINGLE(_xlfn.IFS(_xlfn.SINGLE(A1:A3),1,TRUE,2))');
    // a stray @ with nothing to apply to stays
    expect(toFileFormula('@')).toBe('@');
    expect(toFileFormula('@+1')).toBe('@+1');
    expect(toFileFormula('@INDEX(A1')).toBe('_xlfn.SINGLE(INDEX(A1)');
  });

  it('writes the spill operator as a call only after a single cell reference', () => {
    expect(toFileFormula('A1#')).toBe('_xlfn.ANCHORARRAY(A1)');
    expect(toFileFormula('SUM($B$2#,A1#)')).toBe('SUM(_xlfn.ANCHORARRAY($B$2),_xlfn.ANCHORARRAY(A1))');
    expect(toFileFormula("'My sheet'!A1#")).toBe("_xlfn.ANCHORARRAY('My sheet'!A1)");
    expect(toFileFormula('#')).toBe('#');
    expect(toFileFormula('A1 #')).toBe('A1 #');
    expect(toFileFormula('SUM#')).toBe('SUM#');
    expect(toFileFormula('A1:A2#')).toBe('A1:_xlfn.ANCHORARRAY(A2)');
  });

  it('prefixes the parameters of LET and LAMBDA only where the call uses them, nested calls included', () => {
    expect(toFileFormula('LET(x,B2,y,x*2,x+y)')).toBe('_xlfn.LET(_xlpm.x,B2,_xlpm.y,_xlpm.x*2,_xlpm.x+_xlpm.y)');
    expect(toFileFormula('LAMBDA(a,b,a+b)(1,2)')).toBe('_xlfn.LAMBDA(_xlpm.a,_xlpm.b,_xlpm.a+_xlpm.b)(1,2)');
    expect(toFileFormula('LET(f,LAMBDA(n,n+1),f(1))')).toBe('_xlfn.LET(_xlpm.f,_xlfn.LAMBDA(_xlpm.n,_xlpm.n+1),_xlpm.f(1))');
    // outside the call the same name is a defined name, not a parameter
    expect(toFileFormula('x+LET(x,1,x)')).toBe('x+_xlfn.LET(_xlpm.x,1,_xlpm.x)');
    // an argument that is not a bare name declares nothing; an unclosed call still ends the formula
    expect(toFileFormula('LET(A1+1,2,3)')).toBe('_xlfn.LET(A1+1,2,3)');
    expect(toFileFormula('LET(x,1,x')).toBe('_xlfn.LET(_xlpm.x,1,_xlpm.x');
  });

  it('is exact on formulas a file can hold: from(to(x)) = x and to(from(f)) = f for the table rows', () => {
    for (const [file, shown] of PAIRS) {
      expect(fromFileFormula(file)).toBe(shown);
      if (!shown.includes('_xlfn')) expect(toFileFormula(shown)).toBe(file);
    }
  });

  it('has a prefix for every future function and a plain name only in the table', () => {
    expect(FUTURE_FUNCTIONS.get('IFS')).toBe('_xlfn.');
    expect(FUTURE_FUNCTIONS.get('SORT')).toBe('_xlfn._xlws.');
    expect(FUTURE_FUNCTIONS.get('FILTER')).toBe('_xlfn._xlws.');
    expect(FUTURE_FUNCTIONS.get('SUM')).toBeUndefined();
    for (const [name, prefix] of FUTURE_FUNCTIONS) {
      expect(name).toBe(name.toUpperCase());
      expect(prefix.startsWith('_xlfn.')).toBe(true);
    }
  });

  it('treats object-prototype names as ordinary names', () => {
    for (const name of ['__proto__', 'constructor', 'toString', 'hasOwnProperty', 'valueOf']) {
      expect(toFileFormula(`${name}(1)+${name}`)).toBe(`${name}(1)+${name}`);
      expect(fromFileFormula(`_xlfn.${name}(1)`)).toBe(`${name}(1)`);
      expect(toFileFormula(`LET(${name},1,${name})`)).toBe(`_xlfn.LET(_xlpm.${name},1,_xlpm.${name})`);
    }
    expect(FUTURE_FUNCTIONS.get('constructor')).toBeUndefined();
  });

  it('is linear-ish on a long formula: deep nesting at the 8192-character limit finishes at once and unchanged in meaning', () => {
    const deep = '_xlfn.SINGLE('.repeat(600) + 'A1' + ')'.repeat(600);
    expect(fromFileFormula(deep)).toBe('@'.repeat(600) + 'A1');
    const wide = '_xlfn.IFS(A1,1,'.repeat(500) + '0' + ')'.repeat(500);
    expect(toFileFormula(fromFileFormula(wide))).toBe(wide);
    const letters = 'A'.repeat(8192);
    expect(fromFileFormula(letters)).toBe(letters);
    expect(toFileFormula('@'.repeat(2000))).toBe('@'.repeat(2000));
  });
});
