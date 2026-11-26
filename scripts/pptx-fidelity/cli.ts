/**
 * PPTX fidelity corpus runner (the flags are the shared ones, see scripts/fidelity-core/cli.ts).
 *
 *   npm run slides:pptx-fidelity                  compare every fixture with its golden snapshots (exit 1 on drift)
 *   npm run slides:pptx-fidelity:update           rewrite the snapshots (and the README coverage matrix)
 *   npm run slides:pptx-fidelity -- --only id     run a subset
 *   npm run slides:pptx-fidelity -- --matrix      verify claims and print the coverage matrix
 *   npm run slides:pptx-fidelity -- --report md   also print a markdown report (--report json for JSON);
 *       add --report-out <file> to append it to a file (CI passes "$GITHUB_STEP_SUMMARY"); repeatable
 *   npm run slides:pptx-fidelity -- --runtime     also scan the run-time files (the 1000-slide deck); no goldens
 *
 * Runs against the built @b32nio/spindle-io/pptx (run `npm run build` first).
 */
import { main, type CliConfig } from '../fidelity-core/cli';
import type { PptxEditorApi } from './api';
import { loadBuiltApi } from './api';
import { CONSTRUCTS } from './constructs';
import { MATRIX_END, MATRIX_START, PPTX_FIXTURES_DIR, PPTX_MANIFEST_OPTIONS } from './index';
import { toPptxReport } from './report';
import { pptxAdapter, type PptxResult } from './run';
import { runRuntime } from './runtime';

export function pptxCliConfig(
  api: PptxEditorApi,
  fixturesDir = PPTX_FIXTURES_DIR
): CliConfig<PptxResult> {
  return {
    format: 'PPTX',
    updateCommand: 'npm run slides:pptx-fidelity:update',
    fixturesDir,
    manifestOptions: PPTX_MANIFEST_OPTIONS,
    registry: CONSTRUCTS,
    matrixMarkers: { start: MATRIX_START, end: MATRIX_END },
    adapter: pptxAdapter(api),
    toReport: toPptxReport,
    runRuntime: (entries, args, io) => runRuntime({ api, fixturesDir }, entries, args, io),
  };
}

/* istanbul ignore next -- process bootstrap: runs only when the file is executed (tsx), not when Jest imports it */
if (require.main === module) {
  loadBuiltApi().then(
    (api) => main(pptxCliConfig(api), process.argv.slice(2)),
    (e) => {
      console.error(e instanceof Error ? e.stack : e);
      process.exit(2);
    }
  );
}
