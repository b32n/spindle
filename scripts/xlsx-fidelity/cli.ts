/**
 * XLSX fidelity corpus runner (the flags are the shared ones, see scripts/fidelity-core/cli.ts).
 *
 *   npm run sheets:xlsx-fidelity                  compare every fixture with its golden snapshots (exit 1 on drift)
 *   npm run sheets:xlsx-fidelity:update           rewrite the snapshots (and the README coverage matrix)
 *   npm run sheets:xlsx-fidelity -- --only id     run a subset
 *   npm run sheets:xlsx-fidelity -- --matrix      verify claims and print the coverage matrix
 *   npm run sheets:xlsx-fidelity -- --report md   also print a markdown report (--report json for JSON);
 *       add --report-out <file> to append it to a file (CI passes "$GITHUB_STEP_SUMMARY"); repeatable
 *   npm run sheets:xlsx-fidelity -- --runtime     also import the run-time files (bomb, huge, perf); no goldens
 *
 * Runs against the built @b32nio/spindle-sheets-core (run `npm run build` first).
 */
import { main, type CliConfig } from '../fidelity-core/cli';
import type { XlsxEditorApi } from './api';
import { loadBuiltApi } from './api';
import { CONSTRUCTS } from './constructs';
import { MATRIX_END, MATRIX_START, XLSX_FIXTURES_DIR, XLSX_MANIFEST_OPTIONS } from './index';
import { toXlsxReport } from './report';
import { xlsxAdapter, type XlsxResult } from './run';
import { runRuntime } from './runtime';

export function xlsxCliConfig(
  api: XlsxEditorApi,
  fixturesDir = XLSX_FIXTURES_DIR
): CliConfig<XlsxResult> {
  return {
    format: 'XLSX',
    updateCommand: 'npm run sheets:xlsx-fidelity:update',
    fixturesDir,
    manifestOptions: XLSX_MANIFEST_OPTIONS,
    registry: CONSTRUCTS,
    matrixMarkers: { start: MATRIX_START, end: MATRIX_END },
    adapter: xlsxAdapter(api),
    toReport: toXlsxReport,
    runRuntime: (entries, args, io) => runRuntime({ api, fixturesDir }, entries, args, io),
  };
}

/* istanbul ignore next -- process bootstrap: runs only when the file is executed (tsx), not when Jest imports it */
if (require.main === module) {
  loadBuiltApi().then(
    (api) => main(xlsxCliConfig(api), process.argv.slice(2)),
    (e) => {
      console.error(e instanceof Error ? e.stack : e);
      process.exit(2);
    }
  );
}
