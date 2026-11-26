import { parseFormulaSyntax } from './syntax';
import type { ParsedFormulaNode } from './types';

const parse = (text: string, row = 0, col = 0): ParsedFormulaNode => parseFormulaSyntax(text, row, col);

const num = (value: number): ParsedFormulaNode => ({ type: 'number', value });
const op = (operator: string, left: ParsedFormulaNode, right: ParsedFormulaNode): ParsedFormulaNode => ({
  type: 'operator',
  operator,
  left,
  right,
});
const neg = (right: ParsedFormulaNode): ParsedFormulaNode => op('-', { type: 'string', value: '' }, right);

/** Compact structural form of a tree, for readable precedence assertions. */
function shape(node: ParsedFormulaNode): string {
  switch (node.type) {
    case 'number':
      return String(node.value);
    case 'string':
      return JSON.stringify(node.value);
    case 'operator':
      return `(${shape(node.left as ParsedFormulaNode)}${node.operator}${shape(node.right as ParsedFormulaNode)})`;
    case 'percent':
      return `${shape(node.left as ParsedFormulaNode)}%`;
    case 'implicit':
      return `@${shape(node.left as ParsedFormulaNode)}`;
    case 'cell':
      return 'cell';
    case 'range':
      return 'range';
    case 'function':
      return `${node.functionName}(${(node.args ?? []).map(shape).join(',')})`;
    default:
      return node.type;
  }
}

describe('operator precedence and associativity', () => {
  it.each([
    ['1+2*3', '(1+(2*3))'],
    ['1*2+3', '((1*2)+3)'],
    ['1-2-3', '((1-2)-3)'],
    ['8/4/2', '((8/4)/2)'],
    ['2^3^2', '((2^3)^2)'],
    ['1+2&3+4', '((1+2)&(3+4))'],
    ['1&2=3&4', '((1&2)=(3&4))'],
    ['1=2<>3', '((1=2)<>3)'],
    ['1<2', '(1<2)'],
    ['1>=2', '(1>=2)'],
    ['1<=2', '(1<=2)'],
    ['2*3^2', '(2*(3^2))'],
    ['(1+2)*3', '((1+2)*3)'],
    ['-2^2', '((""-2)^2)'],
    ['2^-2', '(2^(""-2))'],
    ['1--2', '(1-(""-2))'],
    ['+3', '(""+3)'],
    ['-(1+2)', '(""-(1+2))'],
    ['10%^2', '(10%^2)'],
    ['2^50%', '(2^50%)'],
    ['-50%', '(""-50)%'],
    ['1+50%', '(1+50%)'],
    ['5%%', '5%%'],
    ['(1+2)%', '(1+2)%'],
  ])('%s parses as %s', (text, expected) => {
    expect(shape(parse(text))).toBe(expected);
  });

  it('binds a percent to its operand only, below negation and above power', () => {
    expect(parse('10%')).toEqual({ type: 'percent', left: num(10) });
    expect(parse('-10%')).toEqual({ type: 'percent', left: neg(num(10)) });
    expect(parse('-(10%)')).toEqual(neg({ type: 'percent', left: num(10) }));
  });

  it('keeps the leading sign as an operator over the empty string', () => {
    expect(parse('-1')).toEqual(neg(num(1)));
    expect(parse('--1')).toEqual(neg(neg(num(1))));
  });
});

describe('literals', () => {
  it('parses numbers, strings, booleans, errors and names', () => {
    expect(parse('1.5e2')).toEqual(num(150));
    expect(parse('"a""b"')).toEqual({ type: 'string', value: 'a"b' });
    expect(parse("'single'")).toEqual({ type: 'string', value: 'single' });
    expect(parse('TRUE')).toEqual({ type: 'function', functionName: 'TRUE', args: [] });
    expect(parse('false')).toEqual({ type: 'function', functionName: 'FALSE', args: [] });
    expect(parse('#NULL!')).toEqual({ type: 'error', value: '#NULL!' });
    expect(parse('#n/a')).toEqual({ type: 'error', value: '#N/A' });
    expect(parse('Quarter1_Total')).toEqual({ type: 'variable', value: 'Quarter1_Total' });
  });

  it('parses an empty formula to an empty string', () => {
    expect(parse('')).toEqual({ type: 'string', value: '' });
    expect(parse('   ')).toEqual({ type: 'string', value: '' });
  });
});

describe('function calls', () => {
  it('upper-cases the name and keeps dots and digits', () => {
    expect(parse('sum(1)')).toEqual({ type: 'function', functionName: 'SUM', args: [num(1)] });
    expect(parse('error.type(1)')).toMatchObject({ functionName: 'ERROR.TYPE' });
    expect(parse('LOG10(100)')).toMatchObject({ functionName: 'LOG10' });
    expect(parse('_xlfn.SEQUENCE(3)')).toMatchObject({ functionName: '_XLFN.SEQUENCE' });
  });

  it('allows whitespace before the argument list and inside it', () => {
    expect(parse('SUM (1 , 2)')).toEqual({ type: 'function', functionName: 'SUM', args: [num(1), num(2)] });
  });

  it('parses an empty argument list and skips empty arguments', () => {
    expect(parse('NOW()')).toEqual({ type: 'function', functionName: 'NOW', args: [] });
    expect(parse('IF(1,,2)')).toMatchObject({ args: [num(1), num(2)] });
    expect(parse('SUM(1,)')).toMatchObject({ args: [num(1)] });
    expect(parse('SUM(,)')).toMatchObject({ args: [] });
  });

  it('keeps an empty argument in place for the functions where position matters', () => {
    const blank = { type: 'number', value: undefined };
    expect(parse('NUMBERVALUE(A1,,",")')).toMatchObject({ args: [{ type: 'cell' }, blank, { type: 'string', value: ',' }] });
    expect(parse('workday.intl(1,2,,3)')).toMatchObject({ args: [{}, {}, blank, {}] });
    expect(parse('INDEX(A1:C3,,2)')).toMatchObject({ args: [{ type: 'range' }, blank, num(2)] });
    expect(parse('INDEX(A1:C3,2,)')).toMatchObject({ args: [{ type: 'range' }, num(2), blank] });
    expect(parse('FIXED(1,,)')).toMatchObject({ args: [num(1), blank, blank] });
    expect(parse('_xlfn.INDEX(A1:C3,,2)')).toMatchObject({ args: [{ type: 'range' }, blank, num(2)] }); // Excel's file spelling
    expect(parse('IF(1,,2)')).toMatchObject({ args: [num(1), num(2)] }); // every other function still skips it
  });

  it('nests calls', () => {
    expect(shape(parse('IF(A1>1,SUM(1,2),MAX(3))'))).toBe('IF((cell>1),SUM(1,2),MAX(3))');
  });

  it('rejects a missing separator or closing parenthesis', () => {
    expect(() => parse('SUM(1 2)')).toThrow('Unexpected "2" at position 6 (expected "," or ")")');
    expect(() => parse('SUM(1')).toThrow('#ERROR! Unexpected end of formula (expected "," or ")")');
    expect(() => parse('SUM(')).toThrow('#ERROR! Unexpected end of formula');
  });
});

describe('cell references', () => {
  it('stores absolute anchors as positions and relative ones as offsets from the formula cell', () => {
    expect(parse('$B$3', 5, 5)).toEqual({
      type: 'cell',
      cellRef: { row: 2, col: 1, rowAbsolute: true, colAbsolute: true, sheetName: undefined },
    });
    expect(parse('B3', 5, 5)).toEqual({
      type: 'cell',
      cellRef: { row: -3, col: -4, rowAbsolute: false, colAbsolute: false, sheetName: undefined },
    });
    expect(parse('$B3', 5, 5)).toMatchObject({ cellRef: { row: -3, col: 1, rowAbsolute: false, colAbsolute: true } });
    expect(parse('B$3', 5, 5)).toMatchObject({ cellRef: { row: 2, col: -4, rowAbsolute: true, colAbsolute: false } });
  });

  it('qualifies a reference with its sheet, quoted or not', () => {
    expect(parse('Sheet2!A1')).toMatchObject({ type: 'cell', cellRef: { sheetName: 'Sheet2' } });
    expect(parse("'My Sheet'!A1")).toMatchObject({ cellRef: { sheetName: 'My Sheet' } });
  });

  it('requires a reference right after a sheet prefix', () => {
    expect(() => parse('Sheet1!')).toThrow('#ERROR! Unexpected end of formula (a reference must follow "Sheet1!")');
    expect(() => parse('Sheet1! A1')).toThrow('(a reference must follow "Sheet1!")');
    expect(() => parse('Sheet1!5')).toThrow('(a reference must follow "Sheet1!")');
    expect(() => parse('Sheet1:!A1')).toThrow('Unexpected');
  });

  it('keeps an error literal after a sheet prefix as the error', () => {
    expect(parse('Sheet1!#REF!')).toEqual({ type: 'error', value: '#REF!' });
  });
});

describe('ranges', () => {
  it('joins two cells into a range node, sharing the sheet', () => {
    expect(parse('A1:B2')).toEqual({
      type: 'range',
      rangeRef: {
        start: { row: 0, col: 0, rowAbsolute: false, colAbsolute: false, sheetName: undefined },
        end: { row: 1, col: 1, rowAbsolute: false, colAbsolute: false, sheetName: undefined },
      },
    });
    const r = parse('Sheet2!A1:B2');
    expect(r.rangeRef?.start.sheetName).toBe('Sheet2');
    expect(r.rangeRef?.end.sheetName).toBe('Sheet2');
    expect(parse('Sheet2!A1:Sheet2!B2').rangeRef?.end.sheetName).toBe('Sheet2');
  });

  it('keeps a range between differently-qualified cells as the operator', () => {
    expect(shape(parse('A1:Sheet2!B2'))).toBe('(cell:cell)');
    expect(shape(parse('Sheet1!A1:Sheet2!B2'))).toBe('(cell:cell)');
  });

  it('builds a whole-column range with the nominal sheet extent', () => {
    expect(parse('$B:D', 3, 3)).toEqual({
      type: 'range',
      rangeRef: {
        start: { row: 0, col: 1, rowAbsolute: true, colAbsolute: true, sheetName: undefined },
        end: { row: 1048575, col: 0, rowAbsolute: true, colAbsolute: false, sheetName: undefined },
        whole: 'column',
      },
    });
  });

  it('builds a whole-row range with the nominal sheet extent', () => {
    expect(parse('2:$4', 3, 3)).toEqual({
      type: 'range',
      rangeRef: {
        start: { row: -2, col: 0, rowAbsolute: false, colAbsolute: true, sheetName: undefined },
        end: { row: 3, col: 16383, rowAbsolute: true, colAbsolute: true, sheetName: undefined },
        whole: 'row',
      },
    });
  });

  it('qualifies whole-column and whole-row ranges with a sheet', () => {
    expect(parse('Data!A:A').rangeRef?.start.sheetName).toBe('Data');
    expect(parse('Data!1:1').rangeRef?.end.sheetName).toBe('Data');
  });

  it('builds the range operator over computed references', () => {
    expect(shape(parse('A1:INDEX(A1:A5,3)'))).toBe('(cell:INDEX(range,3))');
    expect(shape(parse('INDEX(A1:A5,1):INDEX(A1:A5,3)'))).toBe('(INDEX(range,1):INDEX(range,3))');
    expect(shape(parse('A1:B2:C3'))).toBe('(range:cell)');
    expect(shape(parse('OFFSET(A1,1,1):C5'))).toBe('(OFFSET(cell,1,1):cell)');
  });

  it('rejects a range operator without references on both sides', () => {
    expect(() => parse('A:')).toThrow('#ERROR! Unexpected end of formula');
    expect(() => parse('1:A')).toThrow('The range operator ":" needs references on both sides (position 1)');
    expect(() => parse('A1:1')).toThrow('The range operator ":" needs references');
    expect(() => parse('"a":A1')).toThrow('The range operator ":" needs references');
    expect(() => parse('A1:{1}')).toThrow('The range operator ":" needs references');
    expect(() => parse(':A1')).toThrow('Unexpected ":" at position 0');
  });
});

describe('3D references', () => {
  it('parses a 3D cell and a 3D range, keeping the last sheet beside the first', () => {
    expect(parse('Sheet1:Sheet3!A1')).toEqual({
      type: 'ref3d',
      sheetEnd: 'Sheet3',
      cellRef: { row: 0, col: 0, rowAbsolute: false, colAbsolute: false, sheetName: 'Sheet1' },
    });
    const range = parse('Sheet1:Sheet3!A1:B2');
    expect(range.type).toBe('ref3d');
    expect(range.sheetEnd).toBe('Sheet3');
    expect(range.cellRef).toBeUndefined();
    expect(range.rangeRef?.start.sheetName).toBe('Sheet1');
    expect(range.rangeRef?.end).toMatchObject({ row: 1, col: 1, sheetName: 'Sheet1' });
  });

  it('parses quoted spans and 3D whole-column / whole-row ranges', () => {
    expect(parse("'Sheet 1:Sheet 3'!A1")).toMatchObject({ type: 'ref3d', sheetEnd: 'Sheet 3' });
    expect(parse('Sheet1:Sheet3!A:A')).toMatchObject({ type: 'ref3d', rangeRef: { whole: 'column' } });
    expect(parse('Sheet1:Sheet3!1:2')).toMatchObject({ type: 'ref3d', rangeRef: { whole: 'row' } });
  });

  it('does not extend a 3D cell with a differently-qualified cell', () => {
    expect(shape(parse('Sheet1:Sheet3!A1:Sheet4!B2'))).toBe('(ref3d:cell)');
  });
});

describe('union and intersection', () => {
  it('builds a union only inside parentheses', () => {
    expect(shape(parse('(A1:A2,C1)'))).toBe('(range,cell)');
    expect(shape(parse('(A1,B1,C1)'))).toBe('((cell,cell),cell)');
    expect(shape(parse('SUM((A1:A2,C1))'))).toBe('SUM((range,cell))');
    expect(shape(parse('SUM(A1:A2,C1)'))).toBe('SUM(range,cell)');
  });

  it('rejects a union of non-references', () => {
    expect(() => parse('(1,2)')).toThrow('The union operator "," needs references on both sides (position 2)');
    expect(() => parse('(A1,2)')).toThrow('The union operator ","');
    expect(() => parse('(A1,B1')).toThrow('Unexpected end of formula (expected ")")');
  });

  it('builds an intersection from a space between references', () => {
    expect(shape(parse('A1:C3 B2:D4'))).toBe('(range range)');
    expect(shape(parse('SUM(A1:A3 A2:A4)'))).toBe('SUM((range range))');
    expect(shape(parse('A1 B1 C1'))).toBe('((cell cell) cell)');
    expect(shape(parse('A:A 1:1'))).toBe('(range range)');
    expect(shape(parse('Sheet1!A1 Sheet1!B1'))).toBe('(cell cell)');
    expect(shape(parse('(A1:B2) (B1:C3)'))).toBe('(range range)');
    expect(shape(parse('A1 B1:C1 D1'))).toBe('((cell range) cell)');
  });

  it('gives the range operator priority over intersection and intersection over arithmetic', () => {
    expect(shape(parse('A1:B2 B1:C3'))).toBe('(range range)');
    expect(shape(parse('1+A1:B2 B1:C3'))).toBe('(1+(range range))');
  });

  it('does not read a space before an operator as an intersection', () => {
    expect(shape(parse('A1 + B1'))).toBe('(cell+cell)');
    expect(shape(parse('A1 -B1'))).toBe('(cell-cell)');
    expect(shape(parse('SUM(A1 , B1)'))).toBe('SUM(cell,cell)');
  });

  it('rejects an intersection of non-references', () => {
    expect(() => parse('2 A1')).toThrow('The intersection operator needs references on both sides (position 2)');
    expect(() => parse('A1 (2)')).toThrow('The intersection operator needs references');
    expect(() => parse('A1 "x"')).toThrow('Unexpected "x" at position 3');
  });

  it('accepts error literals as reference operands, so a formula whose reference was deleted still parses', () => {
    expect(shape(parse('SUM(#REF!:A1)'))).toBe('SUM((error:cell))');
    expect(shape(parse('#REF! #REF!'))).toBe('(error error)');
    expect(shape(parse('(#REF!,A1)'))).toBe('(error,cell)');
    expect(shape(parse('SUM(A1:INDEX(#REF!,3))'))).toBe('SUM((cell:INDEX(error,3)))');
  });

  it('keeps the formula-level parentheses of a lone reference group', () => {
    expect(shape(parse('(A1)'))).toBe('cell');
  });
});

describe('implicit intersection and the spill operator', () => {
  it('parses @ over a range-level operand', () => {
    expect(shape(parse('@A1:A3'))).toBe('@range');
    expect(shape(parse('@A1'))).toBe('@cell');
    expect(shape(parse('@INDEX(A1:A3,1)'))).toBe('@INDEX(range,1)');
    expect(shape(parse('1+@A1:A3'))).toBe('(1+@range)');
    expect(shape(parse('@A1:A3*2'))).toBe('(@range*2)');
  });

  it('parses # after a cell as the spill operator', () => {
    expect(parse('A1#')).toEqual({
      type: 'spillref',
      cellRef: { row: 0, col: 0, rowAbsolute: false, colAbsolute: false, sheetName: undefined },
    });
    expect(parse('Sheet2!$B$2#')).toMatchObject({ type: 'spillref', cellRef: { sheetName: 'Sheet2', row: 1, col: 1 } });
    expect(shape(parse('SUM(A1#)'))).toBe('SUM(spillref)');
    expect(shape(parse('A1#:B1'))).toBe('(spillref:cell)');
    expect(shape(parse('A1:B2#'))).toBe('(cell:spillref)');
  });

  it('rejects a spill operator that does not follow a cell', () => {
    expect(() => parse('(A1)#')).toThrow('Unknown error literal');
    expect(() => parse('A1 #')).toThrow('Unknown error literal');
  });
});

describe('array constants', () => {
  it('parses rows and columns', () => {
    expect(parse('{1,2,3}')).toEqual({ type: 'array', value: [[1, 2, 3]] });
    expect(parse('{1;2;3}')).toEqual({ type: 'array', value: [[1], [2], [3]] });
    expect(parse('{1,2;3,4}')).toEqual({ type: 'array', value: [[1, 2], [3, 4]] });
    expect(parse('{ 1 , 2 ; 3 , 4 }')).toEqual({ type: 'array', value: [[1, 2], [3, 4]] });
    expect(parse('{5}')).toEqual({ type: 'array', value: [[5]] });
  });

  it('holds numbers, signed numbers, text, booleans and errors', () => {
    expect(parse('{-1,+2,"a""b",TRUE,false,#N/A,#DIV/0!,1.5e1}')).toEqual({
      type: 'array',
      value: [[-1, 2, 'a"b', true, false, { error: '#N/A' }, { error: '#DIV/0!' }, 15]],
    });
  });

  it('works as a function argument and an operand', () => {
    expect(shape(parse('SUM({1,2;3,4})'))).toBe('SUM(array)');
    expect(shape(parse('{1,2}*2'))).toBe('(array*2)');
  });

  it('rejects ragged arrays with the offending row', () => {
    expect(() => parse('{1,2;3}')).toThrow(
      '#ERROR! Array constant at position 0 is ragged: row 2 has 1 elements, row 1 has 2'
    );
    expect(() => parse('{1;2,3}')).toThrow('ragged: row 2 has 2 elements, row 1 has 1');
    expect(() => parse('SUM({1,2;3,4;5})')).toThrow('Array constant at position 4 is ragged: row 3 has 1 elements');
  });

  it('rejects elements an array constant cannot hold', () => {
    expect(() => parse('{}')).toThrow('Unexpected "}" at position 1 (array constants hold only');
    expect(() => parse('{1,}')).toThrow('Unexpected "}" at position 3 (array constants hold only');
    expect(() => parse('{1,;2}')).toThrow('Unexpected ";"');
    expect(() => parse('{A1}')).toThrow('Unexpected "A1" at position 1 (array constants hold only');
    expect(() => parse('{foo}')).toThrow('(array constants hold only numbers, text, TRUE/FALSE and errors)');
    expect(() => parse('{SUM(1)}')).toThrow('Unexpected "SUM"');
    expect(() => parse('{{1}}')).toThrow('Unexpected "{"');
    expect(() => parse('{-}')).toThrow('(a sign must be followed by a number in an array constant)');
    expect(() => parse('{-"a"}')).toThrow('(a sign must be followed by a number in an array constant)');
    expect(() => parse('{1 2}')).toThrow('Unexpected "2" at position 3 (expected ",", ";" or "}" in the array constant)');
    expect(() => parse('{1,2')).toThrow('Unexpected end of formula (expected ",", ";" or "}" in the array constant)');
  });
});

describe('syntax errors', () => {
  it.each([
    ['1+', '#ERROR! Unexpected end of formula'],
    ['*2', '#ERROR! Unexpected "*" at position 0'],
    ['1 2', '#ERROR! Unexpected "2" at position 2'],
    ['(1+2', '#ERROR! Unexpected end of formula (expected ")")'],
    ['1+2)', '#ERROR! Unexpected ")" at position 3'],
    ['()', '#ERROR! Unexpected ")" at position 1'],
    ['}', '#ERROR! Unexpected "}" at position 0'],
    [',', '#ERROR! Unexpected "," at position 0'],
    ['SUM(1)(2)', '#ERROR! Unexpected "(" at position 6'],
    ['"a" "b"', '#ERROR! Unexpected "b" at position 4'],
    ['1A', '#ERROR! Unexpected "A" at position 1'],
    ['@', '#ERROR! Unexpected end of formula'],
    ['Sales[Q1]', '#ERROR! Unexpected "[" at position 5'],
    ['#REF!#', '#ERROR! Unknown error literal at position 5'],
  ])('%s -> %s', (text, message) => {
    expect(() => parse(text)).toThrow(message);
  });

  it('describes every kind of offending token', () => {
    expect(() => parse('1 "s"')).toThrow('Unexpected "s" at position 2');
    expect(() => parse('1 #N/A')).toThrow('The intersection operator needs references');
    expect(() => parse('"s" Sheet1!A1')).toThrow('The intersection operator needs references');
    expect(() => parse('1 1:1')).toThrow('The intersection operator needs references');
    expect(() => parse('1 A:B')).toThrow('The intersection operator needs references');
    expect(() => parse('A1 1')).toThrow('Unexpected "1" at position 3');
    expect(() => parse('1 foo')).toThrow('intersection');
    expect(shape(parse('A1 foo bar'))).toBe('((cell variable) variable)'); // names may denote references
  });

  it('limits the depth of a long flat chain', () => {
    expect(parse('1' + '+1'.repeat(400)).type).toBe('operator');
    for (const chain of ['1' + '%'.repeat(9000), '1' + '+1'.repeat(4900), '1' + '^1'.repeat(4900), 'A1' + ' A1'.repeat(3000), '(A1' + ',A1'.repeat(3000) + ')']) {
      expect(() => parse(chain)).toThrow('#ERROR! Formula is too complex');
    }
    expect(() => parse('SUM(' + '1,'.repeat(100) + '1)')).not.toThrow(); // wide is fine
  });

  it('limits nesting', () => {
    const deep = (n: number) => '('.repeat(n) + '1' + ')'.repeat(n);
    expect(parse(deep(200))).toEqual(num(1));
    expect(() => parse(deep(300))).toThrow('#ERROR! Formula is nested too deeply');
    expect(() => parse('-'.repeat(300) + '1')).toThrow('Formula is nested too deeply');
    expect(() => parse('SUM('.repeat(300) + '1' + ')'.repeat(300))).toThrow('nested too deeply');
    expect(() => parse('@'.repeat(300) + 'A1')).toThrow('nested too deeply');
  });
});
