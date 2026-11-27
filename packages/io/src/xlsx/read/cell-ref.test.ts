import { MAX_COL, MAX_ROW, columnIndex, columnLetters, parseCellRef, parseRangeRef } from './cell-ref';

describe('columnIndex / columnLetters', () => {
  it.each([
    ['A', 0],
    ['Z', 25],
    ['AA', 26],
    ['AZ', 51],
    ['BA', 52],
    ['ZZ', 701],
    ['AAA', 702],
    ['XFD', 16383],
    ['xfd', 16383],
  ])('%s is column %i and back', (letters, index) => {
    expect(columnIndex(letters)).toBe(index);
    expect(columnLetters(index)).toBe(letters.toUpperCase());
  });

  it.each(['', 'XFE', 'ZZZ', 'AAAA', 'A1', '$A', 'é'])('%j is not a column', (letters) => {
    expect(columnIndex(letters)).toBeUndefined();
  });

  it('has the sheet limits', () => {
    expect(MAX_COL).toBe(16384);
    expect(MAX_ROW).toBe(1048576);
  });
});

describe('parseCellRef', () => {
  it.each([
    ['A1', 0, 0],
    ['$B$7', 6, 1],
    ['b7', 6, 1],
    ['XFD1048576', 1048575, 16383],
  ])('%s', (ref, row, col) => {
    expect(parseCellRef(ref)).toEqual({ row, col });
  });

  it.each(['', 'A', '1', 'A0', 'A1048577', 'XFE1', 'A1:B2', 'A-1', '__proto__', 'constructor', 'A99999999'])('%j is not a cell', (ref) => {
    expect(parseCellRef(ref)).toBeUndefined();
  });
});

describe('parseRangeRef', () => {
  it('reads a cell as a one-cell range and a range with its corners in order', () => {
    expect(parseRangeRef('B2')).toEqual({ startRow: 1, startCol: 1, endRow: 1, endCol: 1 });
    expect(parseRangeRef('A1:C3')).toEqual({ startRow: 0, startCol: 0, endRow: 2, endCol: 2 });
    expect(parseRangeRef(' $C$3:$A$1 ')).toEqual({ startRow: 0, startCol: 0, endRow: 2, endCol: 2 });
  });

  it.each(['', 'A1:B2:C3', 'A1:', ':B2', 'A:A', '1:1', 'A1:ZZZ1', 'x'])('%j is not a range', (ref) => {
    expect(parseRangeRef(ref)).toBeUndefined();
  });
});
