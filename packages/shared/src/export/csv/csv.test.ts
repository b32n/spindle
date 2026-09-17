import { exportCsv, importCsv } from './index';

describe('CSV import', () => {
  it('handles CRLF line endings without a trailing \\r on the last column', () => {
    const { rows } = importCsv('a,b\r\nc,d\r\n');
    expect(rows[0][1]).toBe('b');
    expect(rows[1][1]).toBe('d');
  });

  it('keeps a quoted field that contains a newline as one cell', () => {
    const { rows } = importCsv('"line1\nline2",z\n');
    expect(rows[0][0]).toBe('line1\nline2');
    expect(rows[0][1]).toBe('z');
    // No phantom extra row from the embedded newline.
    expect(rows.length).toBe(1);
  });

  it('handles quoted commas and escaped quotes', () => {
    const { rows } = importCsv('"a,b","c""d"\n');
    expect(rows[0][0]).toBe('a,b');
    expect(rows[0][1]).toBe('c"d');
  });

  it('preserves a whitespace-only field between commas', () => {
    const { rows } = importCsv('a, ,c\n');
    expect(rows[0][1]).toBe(' ');
    expect(rows[0][2]).toBe('c');
  });

  it('coerces round-trip-safe numeric strings to numbers, and preserves the rest as text', () => {
    const { rows } = importCsv('42,3.14,007,1e21,5.0\n');
    expect(rows[0]).toEqual([42, 3.14, '007', '1e21', '5.0']);
  });
});

describe('CSV export', () => {
  it('quotes fields containing commas, quotes, or newlines', () => {
    const csv = exportCsv({ rows: [['a,b', 'c"d', 'e\nf', 'plain']] });
    expect(csv).toBe('"a,b","c""d","e\nf",plain');
  });

  it('guards a leading formula-trigger character in a string value', () => {
    const csv = exportCsv({ rows: [['=1+1', '+SUM(A1)', '-5text', '@cmd', 'safe']] });
    expect(csv).toBe("'=1+1,'+SUM(A1),'-5text,'@cmd,safe");
  });

  it('never guards or quotes numbers/booleans', () => {
    const csv = exportCsv({ rows: [[-5, 42, true, false]] });
    expect(csv).toBe('-5,42,true,false');
  });

  it('renders null as an empty field', () => {
    const csv = exportCsv({ rows: [[null, 'a', null]] });
    expect(csv).toBe(',a,');
  });
});

describe('CSV export -> import round-trips', () => {
  it('round-trips cells containing commas and newlines', () => {
    const csv = exportCsv({ rows: [['x\ny', 'a,b'], ['plain', null]] });
    const { rows } = importCsv(csv);
    expect(rows[0][0]).toBe('x\ny');
    expect(rows[0][1]).toBe('a,b');
    expect(rows[1][0]).toBe('plain');
  });
});
