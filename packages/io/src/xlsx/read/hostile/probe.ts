import { readFileSync } from 'node:fs';
import { importXlsx } from '../../reader';

// Run by `child.ts` in a child process with a small heap limit: imports the package at argv[2] (with the limits of argv[3], a JSON
// object) and prints what it cost and what it gave, as one JSON line. The peak is the process's own (`maxRSS`, in KB), as nothing else
// runs in it. With `--expose-gc` it also prints what the result keeps once everything else is collected (`retained`).
const bytes = new Uint8Array(readFileSync(process.argv[2]));
const limits = process.argv[3] ? (JSON.parse(process.argv[3]) as Record<string, number>) : undefined;
const started = Date.now();
const mb = (n: number): number => Math.round(n / (1024 * 1024));
const gc = (globalThis as { gc?: () => void }).gc;

importXlsx(bytes, { limits }).then(
  async (result) => {
    const ms = Date.now() - started;
    const peakRssMb = mb(process.resourceUsage().maxRSS * 1024);
    let retained: { heapMb: number; externalMb: number } | undefined;
    if (gc) {
      // the engine keeps the subject of the last regular expression match (here the whole part); any later match lets it go
      /x/.test('x');
      // external strings are released when the collector's follow-up work has run, a turn of the event loop after the collection
      for (let i = 0; i < 3; i++) {
        gc();
        await new Promise((resolve) => setTimeout(resolve, 20));
      }
      const m = process.memoryUsage();
      retained = { heapMb: mb(m.heapUsed), externalMb: mb(m.external) };
    }
    console.log(
      JSON.stringify({
        ms,
        peakRssMb,
        retained,
        sheets: result.sheets.map((s) => ({
          name: s.name.length > 40 ? `${s.name.slice(0, 40)}…` : s.name,
          cells: s.cells.size,
          merges: s.mergedRanges?.length ?? 0,
          hyperlinks: s.details?.hyperlinks.length ?? 0,
          columns: s.details?.columns.length ?? 0,
          columnWidths: s.columnWidths?.size ?? 0,
          hiddenCols: s.hiddenCols?.size ?? 0,
          rows: s.details?.rows.length ?? 0,
          extLst: s.details?.extLst?.length ?? 0,
          dimension: s.details?.dimension,
          formulaChars: [...s.cells.values()].reduce((n, c) => n + (c.formula?.length ?? 0), 0),
        })),
        sharedStrings: result.sharedStrings?.length ?? 0,
        definedNames: result.workbook?.definedNames.length ?? 0,
        keptRelationships: result.workbook?.sheets.reduce((n, s) => n + s.rels.length, 0) ?? 0,
        warnings: result.warnings.length,
        warningChars: result.warnings.reduce((n, w) => n + w.length, 0),
        firstWarnings: result.warnings.slice(0, 40),
        limitLoss: result.limitLoss,
      })
    );
  },
  (e: unknown) => {
    console.log(JSON.stringify({ error: String(e), errorCode: (e as { code?: string }).code }));
    process.exitCode = 1;
  }
);
