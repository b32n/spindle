import { columnLabelToIndex, parseRangeReference, parseCellReference } from './cell-reference';

describe('parseRangeReference: whole-column and whole-row forms', () => {
  it('reads a whole-column range with the nominal sheet extent', () => {
    expect(parseRangeReference('A:C')).toEqual({
      start: { row: 0, col: 0, rowAbsolute: true, colAbsolute: false, sheetName: undefined },
      end: { row: 1048575, col: 2, rowAbsolute: true, colAbsolute: false, sheetName: undefined },
      whole: 'column',
    });
    expect(parseRangeReference('$B:$B')?.start).toMatchObject({ col: 1, colAbsolute: true });
    expect(parseRangeReference('b:d')?.end.col).toBe(3);
  });

  it('reads a whole-row range with the nominal sheet extent', () => {
    expect(parseRangeReference('2:4')).toEqual({
      start: { row: 1, col: 0, rowAbsolute: false, colAbsolute: true, sheetName: undefined },
      end: { row: 3, col: 16383, rowAbsolute: false, colAbsolute: true, sheetName: undefined },
      whole: 'row',
    });
    expect(parseRangeReference('$1:$1')?.start).toMatchObject({ row: 0, rowAbsolute: true });
  });

  it('keeps the sheet on both corners, quoted or not', () => {
    expect(parseRangeReference('Data!A:A')?.end.sheetName).toBe('Data');
    expect(parseRangeReference("'My Data'!3:3")?.start.sheetName).toBe('My Data');
  });

  it('puts the sheet on both corners of a plain range', () => {
    const r = parseRangeReference('Data!A1:B2');
    expect([r?.start.sheetName, r?.end.sheetName]).toEqual(['Data', 'Data']);
    expect(parseRangeReference("'My Data'!$A$1:B2")?.end.sheetName).toBe('My Data');
  });

  it('still reads plain ranges, with no whole flag', () => {
    const r = parseRangeReference('A1:B2');
    expect(r?.whole).toBeUndefined();
    expect(r?.end).toMatchObject({ row: 1, col: 1 });
  });

  it('rejects anything else', () => {
    expect(parseRangeReference('A')).toBeNull();
    expect(parseRangeReference('A:')).toBeNull();
    expect(parseRangeReference(':A')).toBeNull();
    expect(parseRangeReference('A:1')).toBeNull();
    expect(parseRangeReference('1:A')).toBeNull();
    expect(parseRangeReference('A1:B')).toBeNull();
    expect(parseRangeReference('XFE:XFE')).toBeNull();
    expect(parseRangeReference('A:XFE')).toBeNull();
    expect(parseRangeReference('0:1')).toBeNull();
    expect(parseRangeReference('1:0')).toBeNull();
    expect(parseRangeReference('1048577:1')).toBeNull();
    expect(parseRangeReference('1:1048577')).toBeNull();
    expect(parseRangeReference('ABCD:ABCD')).toBeNull();
  });

  it('accepts the last column and row', () => {
    expect(parseRangeReference('XFD:XFD')?.start.col).toBe(16383);
    expect(parseRangeReference('1048576:1048576')?.start.row).toBe(1048575);
  });
});

describe('parseCellReference', () => {
  it('is unchanged for a plain cell', () => {
    expect(parseCellReference('$B$3')).toMatchObject({ row: 2, col: 1, rowAbsolute: true, colAbsolute: true });
    expect(parseCellReference('A:A')).toBeNull();
  });
});

describe('parseCellReference and columnLabelToIndex edge cases', () => {
  it('rejects row 0 and anything that is not letters then digits', () => {
    expect(parseCellReference('A0')).toBeNull();
    expect(parseCellReference('1A')).toBeNull();
    expect(parseCellReference('A')).toBeNull();
  });
  it('maps column letters and rejects a non-letter', () => {
    expect(columnLabelToIndex('A')).toBe(0);
    expect(columnLabelToIndex('xfd')).toBe(16383);
    expect(columnLabelToIndex('A1')).toBe(-1);
  });
});
