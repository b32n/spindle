import { normalizeFunctionName, scanFormula, VOLATILE_FUNCTIONS } from './formula-scan';

describe('normalizeFunctionName', () => {
  it('strips the future-function prefixes and upper-cases', () => {
    expect(normalizeFunctionName('_xlfn.ifs')).toBe('IFS');
    expect(normalizeFunctionName('_xlfn._xlws.SORT')).toBe('SORT');
    expect(normalizeFunctionName('_XLFN.STDEV.S')).toBe('STDEV.S');
    expect(normalizeFunctionName('sum')).toBe('SUM');
  });
});

describe('scanFormula functions', () => {
  it('lists each function once, in first-seen order, prefixes removed', () => {
    expect(scanFormula('SUM(A1)+_xlfn.IFS(1,2)+sum(2)+_xlws.FILTER(A:A,1)+STDEV.S(B1:B3)').functions).toEqual(['SUM', 'IFS', 'FILTER', 'STDEV.S']);
  });

  it('ignores names in strings and quoted sheet names, and LET/LAMBDA parameter calls', () => {
    expect(scanFormula('IF(A1,"SUM(1)","it""s FOO(")').functions).toEqual(['IF']);
    expect(scanFormula("'My (sheet)'!A1+BAR(1)").functions).toEqual(['BAR']);
    expect(scanFormula('_xlfn.LET(_xlpm.x,1,_xlpm.x(2)+_xlpm.y)').functions).toEqual(['LET']);
  });

  it('does not take a reference, a table name or a number for a call', () => {
    expect(scanFormula('A1+Table1[Col (x)]+1.5+Sheet1!B2').functions).toEqual([]);
    expect(scanFormula('SUM (1)').functions).toEqual(['SUM']);
    expect(scanFormula('').functions).toEqual([]);
  });
});

describe('scanFormula references', () => {
  it.each([
    ['SUM(Table1[Amount])', true],
    ['Table1[[#This Row],[Amount]]*2', true],
    ['[@Amount]*2', true],
    ['SUM(Table1[#Data])', true],
    ['Table1[[#This Row],[It\'s]]', true],
    ['A1+B1', false],
    ['"Table1[x]"', false],
    ["'Sheet [x]'!A1", false],
    ['[1]Sheet1!A1', false],
  ])('structured reference in %s is %s', (text, expected) => {
    expect(scanFormula(text).structuredRef).toBe(expected);
  });

  it.each([
    ['[1]Sheet1!A1+1', true],
    ["'[2]My Sheet'!B2", true],
    ['SUM([3]S!A:A)', true],
    ['Table1[1]', false],
    ['SUM(Sales[2024])', false],
    ['Sales[[2024]]*2', false],
    ['T[[#This Row],[2024]]', false],
    ['SUM(Sales[2024])+[1]S!A1', true],
    ['"[1]x"', false],
    ['A1', false],
  ])('external reference in %s is %s', (text, expected) => {
    expect(scanFormula(text).externalRef).toBe(expected);
  });

  it('copes with an unterminated string, quoted name or bracket', () => {
    expect(() => scanFormula('SUM("abc')).not.toThrow();
    expect(scanFormula("A1+'abc").functions).toEqual([]);
    expect(scanFormula('Table1[[x').structuredRef).toBe(true);
    expect(scanFormula('A1]+SUM(1)').functions).toEqual(['SUM']);
  });
});

describe('VOLATILE_FUNCTIONS', () => {
  it('holds the functions that recalculate on every pass', () => {
    expect([...VOLATILE_FUNCTIONS].sort()).toEqual(['INDIRECT', 'INFO', 'NOW', 'OFFSET', 'RAND', 'RANDARRAY', 'RANDBETWEEN', 'TODAY']);
  });
});
