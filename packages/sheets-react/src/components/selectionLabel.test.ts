import { formatSelectionLabel } from './selectionLabel';

const R = 1000;
const C = 100;
const rng = (startRow: number, endRow: number, startCol: number, endCol: number) => ({
  startRow,
  endRow,
  startCol,
  endCol,
});

describe('formatSelectionLabel', () => {
  it('returns empty for no selection', () => {
    expect(formatSelectionLabel(undefined, R, C)).toBe('');
  });
  it('labels a single cell and a block', () => {
    expect(formatSelectionLabel(rng(1, 1, 1, 1), R, C)).toBe('B2');
    expect(formatSelectionLabel(rng(1, 4, 1, 3), R, C)).toBe('B2:D5');
  });
  it('labels whole columns', () => {
    expect(formatSelectionLabel(rng(0, R - 1, 0, 0), R, C)).toBe('A:A');
    expect(formatSelectionLabel(rng(0, R - 1, 1, 3), R, C)).toBe('B:D');
  });
  it('labels whole rows', () => {
    expect(formatSelectionLabel(rng(1, 1, 0, C - 1), R, C)).toBe('2:2');
    expect(formatSelectionLabel(rng(1, 4, 0, C - 1), R, C)).toBe('2:5');
  });
  it('normalises reversed ranges', () => {
    expect(formatSelectionLabel(rng(4, 1, 3, 1), R, C)).toBe('B2:D5');
  });
  it('labels the whole sheet as a block', () => {
    expect(formatSelectionLabel(rng(0, R - 1, 0, C - 1), R, C)).toBe('A1:CV1000');
  });
});
