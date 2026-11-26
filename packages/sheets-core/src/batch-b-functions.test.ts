// Batch-B functions through the workbook: omitted arguments, spilled results and INDEX as a reference.

import { WorkbookImpl } from './workbook';

function sheetWith(cells: Record<string, unknown> = {}) {
  const w = new WorkbookImpl('wb', 'WB');
  const id = w.getSheet().id;
  const at = (a: string): [number, number] => {
    const m = /^([A-Z]+)(\d+)$/.exec(a) as RegExpExecArray;
    let col = 0;
    for (const ch of m[1]) col = col * 26 + ch.charCodeAt(0) - 64;
    return [Number(m[2]) - 1, col - 1];
  };
  for (const [ref, v] of Object.entries(cells)) {
    const [r, c] = at(ref);
    if (typeof v === 'string' && v.startsWith('=')) w.setFormula(id, r, c, v);
    else w.setCellValue(id, r, c, v as never);
  }
  return {
    set: (ref: string, formula: string) => {
      const [r, c] = at(ref);
      w.setFormula(id, r, c, formula);
    },
    value: (ref: string): unknown => {
      const [r, c] = at(ref);
      return w.getCellCalculatedValue(id, r, c);
    },
    block: (ref: string, rows: number, cols: number): unknown[][] => {
      const [r, c] = at(ref);
      return Array.from({ length: rows }, (_, i) => Array.from({ length: cols }, (_, j) => w.getCellCalculatedValue(id, r + i, c + j)));
    },
  };
}

const GRID = { A1: 1, B1: 2, C1: 3, A2: 4, B2: 5, C2: 6, A3: 7, B3: 8, C3: 9 };

describe('an omitted argument keeps the arguments after it in place', () => {
  it('NUMBERVALUE and FIXED take the argument after a blank one', () => {
    const s = sheetWith();
    s.set('A1', '=NUMBERVALUE("1;234",,";")');
    expect(s.value('A1')).toBe(1234);
    s.set('A2', '=FIXED(1234.5678,,TRUE)');
    expect(s.value('A2')).toBe('1234.57');
  });

  it('WORKDAY.INTL and NETWORKDAYS.INTL read holidays past a blank weekend', () => {
    const s = sheetWith({ D1: 44200 });
    s.set('A1', '=WORKDAY.INTL(44197,1,,D1)'); // Friday 1 Jan 2021: Monday 4 Jan is a holiday
    expect(s.value('A1')).toBe(44201);
    s.set('A2', '=NETWORKDAYS.INTL(44197,44201,,D1)');
    expect(s.value('A2')).toBe(2);
  });

  it('the other functions still skip an empty argument', () => {
    const s = sheetWith();
    s.set('A1', '=IF(TRUE,,2)');
    expect(s.value('A1')).toBe(2);
  });
});

describe('INDEX with a row or column of 0', () => {
  it('spills a whole column or row and feeds aggregates', () => {
    const s = sheetWith(GRID);
    s.set('E1', '=INDEX(A1:C3,0,2)');
    expect(s.block('E1', 3, 1)).toEqual([[2], [5], [8]]);
    s.set('G1', '=INDEX(A1:C3,2,0)');
    expect(s.block('G1', 1, 3)).toEqual([[4, 5, 6]]);
    s.set('G5', '=INDEX(A1:C3,,3)');
    expect(s.block('G5', 3, 1)).toEqual([[3], [6], [9]]);
    s.set('J1', '=SUM(INDEX(A1:C3,0,2))');
    expect(s.value('J1')).toBe(15);
    s.set('J2', '=SUM(INDEX(A1:C3,2,0))');
    expect(s.value('J2')).toBe(15);
    s.set('J3', '=ROWS(INDEX(A1:C3,0,1))');
    expect(s.value('J3')).toBe(3);
    s.set('J4', '=COLUMNS(INDEX(A1:C3,1,0))');
    expect(s.value('J4')).toBe(3);
  });

  it('is a reference: a range endpoint, and a whole column of a whole-column range', () => {
    const s = sheetWith(GRID);
    s.set('J1', '=SUM(A1:INDEX(A1:C3,2,2))');
    expect(s.value('J1')).toBe(12);
    s.set('J2', '=SUM(INDEX(A1:C3,0,2):INDEX(A1:C3,0,3))');
    expect(s.value('J2')).toBe(15 + 18);
    s.set('J3', '=SUM(INDEX(A:C,0,3))');
    expect(s.value('J3')).toBe(18);
    s.set('J4', '=ISREF(INDEX(A1:C3,0,2))');
    expect(s.value('J4')).toBe(true);
  });

  it('recalculates when the indexed cells change', () => {
    const s = sheetWith(GRID);
    s.set('J1', '=SUM(INDEX(A1:C3,0,2))');
    expect(s.value('J1')).toBe(15);
    s.set('B2', '=50');
    expect(s.value('J1')).toBe(60);
  });

  it('reports #REF! outside the range, #VALUE! for a negative index and refuses a bad argument count', () => {
    const s = sheetWith(GRID);
    s.set('J1', '=INDEX(A1:C3,4,0)');
    expect(s.value('J1')).toBe('#REF!');
    s.set('J2', '=INDEX(A1:C3,0,4)');
    expect(s.value('J2')).toBe('#REF!');
    s.set('J3', '=INDEX(A1:C3,-1,1)');
    expect(s.value('J3')).toBe('#VALUE!');
    s.set('J4', '=INDEX(A1:C3,1,1,1,1)');
    expect(s.value('J4')).toBe('#VALUE!');
    s.set('J5', '=INDEX(A1:C3,1,1,2)');
    expect(s.value('J5')).toBe('#REF!');
    s.set('J6', '=INDEX(SEQUENCE(3,3),0,2)');
    expect(s.block('J6', 3, 1)).toEqual([[2], [5], [8]]);
  });
});

describe('dates around the 1900 leap-year quirk', () => {
  it('count Excel\'s 29 February 1900', () => {
    const s = sheetWith();
    s.set('A1', '=DATE(1900,2,29)');
    expect(s.value('A1')).toBe(60);
    s.set('A2', '=DAYS(DATE(1900,3,1),DATE(1900,2,28))');
    expect(s.value('A2')).toBe(2);
    s.set('A3', '=WORKDAY.INTL(DATE(1900,2,27),2)');
    expect(s.value('A3')).toBe(60); // Monday 27 February, then Tuesday and Wednesday the 29th
    s.set('A4', '=NETWORKDAYS.INTL(DATE(1900,1,1),DATE(1900,3,31))');
    expect(s.value('A4')).toBe(65); // 91 days from serial 1 to 91 are 13 whole weeks
    s.set('A5', '=MONTH(60)&"/"&DAY(60)');
    expect(s.value('A5')).toBe('2/29');
  });
});
