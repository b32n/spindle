import {
  BUDGET_100,
  budgetFor,
  countsEqual,
  fmt,
  overBudget,
  pageNumber,
  pairPages,
  quantile,
  resultsTable,
  stem,
  summarize,
  unsafeOut,
  visibleNumbers,
  type DeckResult,
  type Summary,
} from '../../../../scripts/review-pack/helpers';

const one = (v: number): Summary => summarize([v]);
const result = (over: Partial<DeckResult> = {}): DeckResult => ({
  deck: 'bench-100-slides',
  slides: 100,
  bytes: 126068,
  importMs: summarize([100, 200, 300]),
  filmstripMs: summarize([1000, 2000]),
  longestTaskMs: one(80),
  rssMiB: one(250),
  ...over,
});

describe('quantile and summarize', () => {
  it('interpolates between ranks and clamps the fraction', () => {
    expect(quantile([1, 2, 3, 4], 0.5)).toBe(2.5);
    expect(quantile([1, 2, 3, 4], 0)).toBe(1);
    expect(quantile([1, 2, 3, 4], 1)).toBe(4);
    expect(quantile([1, 2, 3, 4], -1)).toBe(1);
    expect(quantile([1, 2, 3, 4], 2)).toBe(4);
    expect(quantile([7], 0.95)).toBe(7);
  });
  it('rejects no samples', () => {
    expect(() => quantile([], 0.5)).toThrow('no samples');
  });
  it('summarizes unsorted samples without changing them', () => {
    const samples = [30, 10, 20];
    expect(summarize(samples)).toEqual({ n: 3, min: 10, median: 20, p95: 29, max: 30, mean: 20 });
    expect(samples).toEqual([30, 10, 20]);
  });
});

describe('budget', () => {
  it('scales linearly above 100 slides and never below the base', () => {
    expect(budgetFor(100)).toEqual(BUDGET_100);
    expect(budgetFor(50)).toEqual(BUDGET_100);
    expect(budgetFor(200)).toEqual({ importMs: 3000, filmstripMs: 10000, rssMiB: 800 });
  });
  it('is met by the example result', () => {
    expect(overBudget(result())).toEqual([]);
  });
  it('names each measurement that is over', () => {
    const r = result({ importMs: one(1600), filmstripMs: one(5001), rssMiB: one(401) });
    expect(overBudget(r)).toEqual([
      'bench-100-slides: import 1600 ms is over 1500 ms',
      'bench-100-slides: filmstrip 5001 ms is over 5000 ms',
      'bench-100-slides: memory 401 MiB is over 400 MiB',
    ]);
  });
  it('takes an explicit budget', () => {
    expect(overBudget(result(), { importMs: 1, filmstripMs: 1e9, rssMiB: 1e9 })).toHaveLength(1);
  });
});

describe('formatting', () => {
  it('rounds', () => {
    expect(fmt(1.5)).toBe('2');
    expect(fmt(1.4)).toBe('1');
  });
  it('renders the results table', () => {
    const table = resultsTable([result()]).split('\n');
    expect(table).toHaveLength(3);
    expect(table[2]).toBe(
      '| bench-100-slides | 100 | 123 KiB | 200 (300) | 1500 (2000) | 80 | 250 |'
    );
    expect(resultsTable([]).split('\n')).toHaveLength(2);
  });
});

describe('without a browser', () => {
  const r = result({ filmstripMs: null, longestTaskMs: null });
  it('is not checked against the filmstrip budget and prints n/a', () => {
    expect(overBudget(r)).toEqual([]);
    expect(resultsTable([r]).split('\n')[2]).toBe(
      '| bench-100-slides | 100 | 123 KiB | 200 (300) | n/a | n/a | 250 |'
    );
  });
});

describe('output directory safety', () => {
  const root = '/work/spindle';
  it('refuses the project, a parent of it, and a directory of other files', () => {
    expect(unsafeOut('/work/spindle', root, null, '.m')).toContain('contains, the project');
    expect(unsafeOut('/work', root, [], '.m')).toContain('contains, the project');
    expect(unsafeOut('/', root, [], '.m')).toContain('contains, the project');
    expect(unsafeOut('/home/me/docs', root, ['a.txt'], '.m')).toContain('not empty');
  });
  it('allows a new, empty, marked or nested directory', () => {
    expect(unsafeOut('/tmp/x', root, null, '.m')).toBeNull();
    expect(unsafeOut('/tmp/x', root, [], '.m')).toBeNull();
    expect(unsafeOut('/tmp/x', root, ['.m', 'lo'], '.m')).toBeNull();
    expect(unsafeOut('/work/spindle/review-pack-output/review', root, ['.m'], '.m')).toBeNull();
    expect(unsafeOut('/work/spindle2', root, [], '.m')).toBeNull();
  });
});

describe('hidden slides', () => {
  it('numbers the visible slides, as LibreOffice numbers its pages', () => {
    expect(visibleNumbers([false, true, false, false, true])).toEqual([1, null, 2, 3, null]);
    expect(visibleNumbers([])).toEqual([]);
  });
  it('compares the visible slides, capped, with the pages', () => {
    expect(countsEqual(4, 4, 12)).toBe(true);
    expect(countsEqual(30, 12, 12)).toBe(true);
    expect(countsEqual(4, 3, 12)).toBe(false);
    expect(countsEqual(30, 11, 12)).toBe(false);
  });
});

describe('pairing', () => {
  it('reads page numbers', () => {
    expect(pageNumber('deck-07.png')).toBe(7);
    expect(pageNumber('slide-012.PNG')).toBe(12);
    expect(pageNumber('deck.png')).toBeNull();
    expect(pageNumber('deck-1.jpg')).toBeNull();
  });
  it('pairs by number, in order, with nulls for a missing side', () => {
    expect(pairPages(['a-02.png', 'a-1.png', 'notes.txt'], ['b-01.png', 'b-03.png'])).toEqual([
      { page: 1, spindle: 'a-1.png', libreOffice: 'b-01.png' },
      { page: 2, spindle: 'a-02.png', libreOffice: null },
      { page: 3, spindle: null, libreOffice: 'b-03.png' },
    ]);
    expect(pairPages([], [])).toEqual([]);
  });
  it('takes the stem of a path', () => {
    expect(stem('/x/y/text-basic.pptx')).toBe('text-basic');
    expect(stem('C:\\x\\a.b.pptx')).toBe('a.b');
    expect(stem('plain')).toBe('plain');
  });
});
