import { readFileSync } from 'node:fs';
import { importXlsx } from '../../reader';

// Run by `hostile-memory.test.ts` in a child process with a small heap limit: imports the package at argv[2] and prints what it cost.
// The peak is the process's own (`maxRSS`, in KB), as nothing else runs in it.
const bytes = new Uint8Array(readFileSync(process.argv[2]));
const started = Date.now();
importXlsx(bytes).then(
  (result) =>
    console.log(
      JSON.stringify({
        ms: Date.now() - started,
        peakRssMb: Math.round(process.resourceUsage().maxRSS / 1024),
        sheets: result.sheets.length,
        cells: result.sheets.reduce((n, s) => n + s.cells.size, 0),
        sharedStrings: result.sharedStrings?.length ?? 0,
        warnings: result.warnings.length,
        warningChars: result.warnings.reduce((n, w) => n + w.length, 0),
        firstWarnings: result.warnings.slice(0, 3),
      })
    ),
  (e: unknown) => {
    console.log(JSON.stringify({ error: String(e) }));
    process.exitCode = 1;
  }
);
