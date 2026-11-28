/**
 * The pure parts of the local review pack (plan K17): sample statistics, the budget check and the Markdown table of the
 * benchmark, and the pairing of the Spindle and LibreOffice page images. Everything with a side effect (browser, soffice,
 * files) lives in `benchmark.ts` and `review-pack.ts`.
 */

export interface Summary {
  n: number;
  min: number;
  median: number;
  p95: number;
  max: number;
  mean: number;
}

/** The value at the given fraction (0 to 1) of the sorted samples, interpolated between the two nearest ranks. */
export function quantile(sorted: readonly number[], q: number): number {
  if (sorted.length === 0) throw new Error('quantile of no samples');
  const pos = Math.min(1, Math.max(0, q)) * (sorted.length - 1);
  const lo = Math.floor(pos);
  const hi = Math.ceil(pos);
  return sorted[lo] + (sorted[hi] - sorted[lo]) * (pos - lo);
}

export function summarize(samples: readonly number[]): Summary {
  const sorted = [...samples].sort((a, b) => a - b);
  return {
    n: sorted.length,
    min: sorted[0],
    median: quantile(sorted, 0.5),
    p95: quantile(sorted, 0.95),
    max: sorted[sorted.length - 1],
    mean: sorted.reduce((a, b) => a + b, 0) / sorted.length,
  };
}

/** One deck's measurements: the median of the runs, in milliseconds (times) and mebibytes (memory). */
export interface DeckResult {
  deck: string;
  slides: number;
  bytes: number;
  importMs: Summary;
  /** From the file chosen in the Open dialog until the filmstrip shows every slide. */
  filmstripMs: Summary;
  /** The longest main-thread task during that open. */
  longestTaskMs: Summary;
  /** Peak resident set of the import process, MiB. */
  rssMiB: Summary;
}

/** The K17 budget (launch-bar row L17): upper limits on the medians for the 100-slide deck, doubled for 200. */
export interface Budget {
  importMs: number;
  filmstripMs: number;
  rssMiB: number;
}

export const BUDGET_100: Budget = { importMs: 1500, filmstripMs: 5000, rssMiB: 400 };

/** The budget of a deck: linear in the slide count above 100 slides (a deck of 200 may take twice as long). */
export function budgetFor(slides: number, base: Budget = BUDGET_100): Budget {
  const k = Math.max(1, slides / 100);
  return {
    importMs: base.importMs * k,
    filmstripMs: base.filmstripMs * k,
    rssMiB: base.rssMiB * k,
  };
}

/** The measurements that are over the deck's budget, as sentences (empty when it is met). */
export function overBudget(r: DeckResult, budget: Budget = budgetFor(r.slides)): string[] {
  const out: string[] = [];
  if (r.importMs.median > budget.importMs)
    out.push(`${r.deck}: import ${fmt(r.importMs.median)} ms is over ${fmt(budget.importMs)} ms`);
  if (r.filmstripMs.median > budget.filmstripMs)
    out.push(
      `${r.deck}: filmstrip ${fmt(r.filmstripMs.median)} ms is over ${fmt(budget.filmstripMs)} ms`
    );
  if (r.rssMiB.median > budget.rssMiB)
    out.push(`${r.deck}: memory ${fmt(r.rssMiB.median)} MiB is over ${fmt(budget.rssMiB)} MiB`);
  return out;
}

export const fmt = (n: number): string => String(Math.round(n));

/** The results as a Markdown table (medians; the slowest run in brackets for the times). */
export function resultsTable(results: readonly DeckResult[]): string {
  const rows = results.map(
    (r) =>
      `| ${r.deck} | ${r.slides} | ${(r.bytes / 1024).toFixed(0)} KiB | ${fmt(r.importMs.median)} (${fmt(r.importMs.max)}) | ` +
      `${fmt(r.filmstripMs.median)} (${fmt(r.filmstripMs.max)}) | ${fmt(r.longestTaskMs.median)} | ${fmt(r.rssMiB.median)} |`
  );
  return [
    '| Deck | Slides | Size | Import ms (max) | Filmstrip ms (max) | Longest task ms | Peak RSS MiB |',
    '| --- | --- | --- | --- | --- | --- | --- |',
    ...rows,
  ].join('\n');
}

/** The 1-based page or slide number in a file written by `pdftoppm` (`name-07.png`) or by the pack (`slide-007.png`), or null. */
export function pageNumber(file: string): number | null {
  const m = /-(\d+)\.png$/i.exec(file);
  return m ? Number(m[1]) : null;
}

export interface Pair {
  page: number;
  spindle: string | null;
  libreOffice: string | null;
}

/** Pairs the two renders by page number, in page order; a page only one side has pairs with null. Files with no number are ignored. */
export function pairPages(spindle: readonly string[], libreOffice: readonly string[]): Pair[] {
  const index = (files: readonly string[]) => {
    const m = new Map<number, string>();
    for (const f of files) {
      const n = pageNumber(f);
      if (n !== null) m.set(n, f);
    }
    return m;
  };
  const a = index(spindle);
  const b = index(libreOffice);
  return [...new Set([...a.keys(), ...b.keys()])]
    .sort((x, y) => x - y)
    .map((page) => ({ page, spindle: a.get(page) ?? null, libreOffice: b.get(page) ?? null }));
}

/** A safe file stem from a fixture file name (`text-basic.pptx` -> `text-basic`). */
export const stem = (file: string): string => file.replace(/^.*[\\/]/, '').replace(/\.[^.]+$/, '');
