import { isFormulaEvaluable, rangeSpan, spanRef } from './xlsx-formulas';

describe('isFormulaEvaluable', () => {
  it.each([
    ['1+2'],
    ['SUM(A1:A3)'],
    ['IFS(A1>1,"a",TRUE,"b")'],
    ['IF(TRUE,1,2)'],
    ['SORT(A1:A3)'],
    ['@A1:A3'],
    ['SUM(A1#)'],
    ['XLOOKUP(1,A:A,B:B)'],
    ['A1+Sheet2!B2'],
  ])('can evaluate %s', (formula) => {
    expect(isFormulaEvaluable(formula)).toBe(true);
  });

  it.each([
    ['LET(x,1,x)'],
    ['HSTACK(A1:A2,B1:B2)'],
    ['A1+++'],
    ['SUM(A1'],
    ['SUM(Table1[col])'],
    ['Quarter1*2'],
    ['SUM(1)+NOSUCHFN(2)'],
    ['IF(1,NOSUCHFN(2),3)'],
    ['{1,2}+NOSUCHFN()'],
    ['('.repeat(4000) + '1' + ')'.repeat(4000)],
    ['A1+'.repeat(3000) + '1'],
  ])('keeps %s as the saved value', (formula) => {
    expect(isFormulaEvaluable(formula)).toBe(false);
  });

  it('looks for names and unknown functions in every operand', () => {
    expect(isFormulaEvaluable('1+(2*SUM(A1,B1:B2))-Name')).toBe(false);
    expect(isFormulaEvaluable('-SUM(1)%')).toBe(true);
  });
});

describe('rangeSpan and spanRef', () => {
  it.each([
    ['G11:G15', { rows: 5, cols: 1 }],
    ['A1', { rows: 1, cols: 1 }],
    ['$B$2:$D$3', { rows: 2, cols: 3 }],
    ['aa10:ab12', { rows: 3, cols: 2 }],
    ['A1:XFD1048576', { rows: 1048576, cols: 16384 }],
  ])('reads %s', (ref, span) => {
    expect(rangeSpan(ref)).toEqual(span);
  });

  it.each([[''], ['nonsense'], ['A0'], ['B1:A1'], ['A2:A1'], ['A1:'], ['1:3'], ['A:A'], ['A1:B2:C3'], ["'S'!A1"]])('does not read %p', (ref) => {
    expect(rangeSpan(ref)).toBeUndefined();
  });

  it('writes the range of a span, a single cell for 1 x 1', () => {
    expect(spanRef(10, 6, 5, 1)).toBe('G11:G15');
    expect(spanRef(0, 0, 1, 1)).toBe('A1');
    expect(spanRef(0, 25, 2, 3)).toBe('Z1:AB2');
    expect(spanRef(1, 0, 1, 4)).toBe('A2:D2');
  });
});
