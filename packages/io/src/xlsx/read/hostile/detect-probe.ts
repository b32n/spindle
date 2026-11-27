import { readFileSync } from 'node:fs';
import { detectXlsxLoss } from '../../loss';
import { importXlsx } from '../../reader';

// Run by `child.ts` (`DETECT_PROBE`) in a child process with a small heap limit: the loss detector and the reader on the package at
// argv[2], timed one after the other, as one JSON line (the shape of `Probe`, with `detectMs` and the ids of the detector's items).
const bytes = new Uint8Array(readFileSync(process.argv[2]));
const timed = async <T>(f: () => Promise<T>): Promise<[T, number]> => {
  const t = Date.now();
  return [await f(), Date.now() - t];
};

(async () => {
  const [report, detectMs] = await timed(() => detectXlsxLoss(bytes, {}));
  const [result, ms] = await timed(() => importXlsx(bytes));
  console.log(JSON.stringify({ ms, detectMs, items: report.items.map((i) => [i.id, i.count]), hyperlinks: result.sheets.reduce((n, s) => n + (s.details?.hyperlinks.length ?? 0), 0), warnings: result.warnings.length }));
})().catch((e: unknown) => {
  console.log(JSON.stringify({ error: String(e) }));
  process.exitCode = 1;
});
