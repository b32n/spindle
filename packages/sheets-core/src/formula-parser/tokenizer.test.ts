import { tokenize, syntaxError, type Token } from './tokenizer';

/** Tokens without their positions, for compact expectations. */
function kinds(text: string): Array<Record<string, unknown>> {
  return tokenize(text)
    .slice(0, -1)
    .map((t: Token) => {
      const { pos: _pos, space: _space, ...rest } = t as Token & { pos: number; space: boolean };
      return rest;
    });
}

describe('tokenize', () => {
  it('always ends with an eof token', () => {
    expect(tokenize('')).toEqual([{ kind: 'eof', pos: 0, space: false }]);
    expect(tokenize('  ').map((t) => t.kind)).toEqual(['eof']);
    expect(tokenize('1').map((t) => t.kind)).toEqual(['num', 'eof']);
  });

  it('reads numbers, including fractions and exponents', () => {
    expect(kinds('1 2.5 .5 5. 1e3 1E-2 2.5e+2')).toEqual(
      [1, 2.5, 0.5, 5, 1000, 0.01, 250].map((value) => ({ kind: 'num', value }))
    );
  });

  it('rejects a number that overflows', () => {
    expect(() => tokenize('1e999')).toThrow('#ERROR! Number out of range at position 0');
  });

  it('reads strings with doubled quotes, and single-quoted text as a string', () => {
    expect(kinds('"a""b" "" \'it\'\'s\'')).toEqual([
      { kind: 'str', value: 'a"b' },
      { kind: 'str', value: '' },
      { kind: 'str', value: "it's" },
    ]);
  });

  it('rejects unterminated strings and quotes', () => {
    expect(() => tokenize('"abc')).toThrow('#ERROR! Unterminated string starting at position 0');
    expect(() => tokenize("1+'abc")).toThrow('#ERROR! Unterminated quote starting at position 2');
  });

  it('reads cell references with every anchoring', () => {
    expect(kinds('A1 $A1 A$1 $A$1 xfd1048576')).toEqual([
      { kind: 'cell', row: 0, col: 0, rowAbs: false, colAbs: false },
      { kind: 'cell', row: 0, col: 0, rowAbs: false, colAbs: true },
      { kind: 'cell', row: 0, col: 0, rowAbs: true, colAbs: false },
      { kind: 'cell', row: 0, col: 0, rowAbs: true, colAbs: true },
      { kind: 'cell', row: 1048575, col: 16383, rowAbs: false, colAbs: false },
    ]);
  });

  it('treats an address past the last column or row, or row 0, as a name', () => {
    expect(kinds('XFE1 A1048577 A0')).toEqual([
      { kind: 'ident', text: 'XFE1' },
      { kind: 'ident', text: 'A1048577' },
      { kind: 'ident', text: 'A0' },
    ]);
  });

  it('treats a cell-shaped word followed by ( as a function name', () => {
    expect(kinds('LOG10(1)').slice(0, 2)).toEqual([
      { kind: 'ident', text: 'LOG10' },
      { kind: 'op', text: '(' },
    ]);
    // With a space before the parenthesis it is still a cell (an intersection with a group).
    expect(kinds('A1 (B1)')[0]).toMatchObject({ kind: 'cell' });
  });

  it('reads whole-column and whole-row ranges', () => {
    expect(kinds('A:A $B:$D 1:1 $2:$4')).toEqual([
      { kind: 'colrange', c1: 0, c2: 0, abs1: false, abs2: false },
      { kind: 'colrange', c1: 1, c2: 3, abs1: true, abs2: true },
      { kind: 'rowrange', r1: 0, r2: 0, abs1: false, abs2: false },
      { kind: 'rowrange', r1: 1, r2: 3, abs1: true, abs2: true },
    ]);
    expect(kinds('xfd:XFD')[0]).toMatchObject({ kind: 'colrange', c1: 16383, c2: 16383 });
    expect(kinds('1048576:1048576')[0]).toMatchObject({ kind: 'rowrange', r1: 1048575 });
  });

  it('rejects whole ranges outside the sheet', () => {
    expect(() => tokenize('XFE:XFE')).toThrow('#ERROR! Column out of range at position 0');
    expect(() => tokenize('1048577:1')).toThrow('#ERROR! Row out of range at position 0');
    expect(() => tokenize('0:1')).toThrow('#ERROR! Row out of range at position 0');
  });

  it('does not read a half-open range as a whole range', () => {
    expect(kinds('A:')).toEqual([
      { kind: 'ident', text: 'A' },
      { kind: 'op', text: ':' },
    ]);
    expect(kinds('A:A1').map((t) => t.kind)).toEqual(['ident', 'op', 'cell']);
    expect(kinds('1:A').map((t) => t.kind)).toEqual(['num', 'op', 'ident']);
  });

  it('reads sheet prefixes, quoted or not, and 3D spans', () => {
    expect(kinds("Sheet1!A1 'My Sheet'!A1 Sheet1:Sheet3!A1 'Sheet 1:Sheet 3'!B2 'it''s'!A1")).toEqual([
      { kind: 'sheet', start: 'Sheet1', end: undefined },
      { kind: 'cell', row: 0, col: 0, rowAbs: false, colAbs: false },
      { kind: 'sheet', start: 'My Sheet', end: undefined },
      { kind: 'cell', row: 0, col: 0, rowAbs: false, colAbs: false },
      { kind: 'sheet', start: 'Sheet1', end: 'Sheet3' },
      { kind: 'cell', row: 0, col: 0, rowAbs: false, colAbs: false },
      { kind: 'sheet', start: 'Sheet 1', end: 'Sheet 3' },
      { kind: 'cell', row: 1, col: 1, rowAbs: false, colAbs: false },
      { kind: 'sheet', start: "it's", end: undefined },
      { kind: 'cell', row: 0, col: 0, rowAbs: false, colAbs: false },
    ]);
  });

  it('accepts unquoted sheet names of words and digits, but not a cell-like start or a spaced span', () => {
    expect(kinds('Q1 Data!A1')[0]).toEqual({ kind: 'sheet', start: 'Q1 Data', end: undefined });
    expect(kinds('My Sheet!A1')[0]).toEqual({ kind: 'sheet', start: 'My Sheet', end: undefined });
    expect(kinds('2024!A1')[0]).toEqual({ kind: 'sheet', start: '2024', end: undefined });
    expect(kinds('1st Quarter!B2')[0]).toEqual({ kind: 'sheet', start: '1st Quarter', end: undefined });
    expect(kinds('A1 Sheet2!B2')[0]).toEqual({ kind: 'sheet', start: 'A1 Sheet2', end: undefined }); // an operand start: a sheet name
    expect(kinds('Sheet1!A1 Sheet2!B2').map((t) => t.kind)).toEqual(['sheet', 'cell', 'sheet', 'cell']); // an intersection
    expect(kinds('Sheet1!A1 My Sheet!B2').map((t) => t.kind)).toEqual(['sheet', 'cell', 'ident', 'sheet', 'cell']);
    expect(kinds('SUM(1,Q1 Data!A1)').map((t) => t.kind)).toContain('sheet');
    expect(kinds('A:D Sheet2!B2').map((t) => t.kind)).toEqual(['colrange', 'sheet', 'cell']);
    expect(kinds('Sheet1:My Sheet!A1')[0].kind).not.toBe('sheet'); // a span with a space needs quotes
    expect(kinds('1.5')[0]).toEqual({ kind: 'num', value: 1.5 });
  });

  it('accepts unicode sheet names and rejects an empty half of a quoted span', () => {
    expect(kinds('Übersicht!A1')[0]).toEqual({ kind: 'sheet', start: 'Übersicht', end: undefined });
    expect(() => tokenize("'Sheet1:'!A1")).toThrow('#ERROR! Empty sheet name at position 0');
    expect(() => tokenize("':Sheet1'!A1")).toThrow('#ERROR! Empty sheet name at position 0');
  });

  it('reads error literals, case-insensitively', () => {
    expect(kinds('#NULL! #div/0! #VALUE! #REF! #NAME? #NUM! #N/A #SPILL! #CALC!').map((t) => t.value)).toEqual([
      '#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#SPILL!', '#CALC!',
    ]);
    expect(() => tokenize('#FOO!')).toThrow('#ERROR! Unknown error literal at position 0');
    expect(() => tokenize('#')).toThrow('#ERROR! Unknown error literal at position 0');
  });

  it('reads # right after a cell as the spill operator, but not after a space', () => {
    expect(kinds('A1#').map((t) => t.kind)).toEqual(['cell', 'op']);
    expect(() => tokenize('A1 #')).toThrow('Unknown error literal');
    expect(kinds('B2#REF!').map((t) => t.kind)).toEqual(['cell', 'op', 'sheet']); // not an error literal: the parser rejects it
  });

  it('reads operators, including the two-character ones', () => {
    expect(kinds('+-*/^&= <> <= >= < > %:,;(){}@').map((t) => t.text)).toEqual([
      '+', '-', '*', '/', '^', '&', '=', '<>', '<=', '>=', '<', '>', '%', ':', ',', ';', '(', ')', '{', '}', '@',
    ]);
  });

  it('marks tokens that follow whitespace', () => {
    expect(tokenize('A1 B1 +C1').map((t) => t.space)).toEqual([false, true, true, false, false]);
  });

  it('records the position of every token', () => {
    expect(tokenize('SUM( A1)').map((t) => t.pos)).toEqual([0, 3, 5, 7, 8]);
  });

  it('reports a stray character', () => {
    expect(() => tokenize('1+$')).toThrow('#ERROR! Unexpected "$" at position 2');
    expect(() => tokenize('1 ? 2')).toThrow('#ERROR! Unexpected "?" at position 2');
    expect(() => tokenize('Sheet1!')).not.toThrow();
    expect(() => tokenize('.')).toThrow('#ERROR! Unexpected "." at position 0');
    expect(() => tokenize('a[1]')).toThrow('#ERROR! Unexpected "[" at position 1');
  });

  it('builds syntax errors with the #ERROR! prefix', () => {
    expect(syntaxError('x').message).toBe('#ERROR! x');
  });
});
