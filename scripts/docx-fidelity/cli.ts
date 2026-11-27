/**
 * DOCX fidelity corpus runner (the flags are the shared ones, see scripts/fidelity-core/cli.ts).
 *
 *   npm run docs:docx-fidelity                 compare every fixture with its golden snapshots (exit 1 on drift)
 *   npm run docs:docx-fidelity:update          rewrite the snapshots (and the README coverage matrix)
 *   npm run docs:docx-fidelity -- --only id    run a subset
 *   npm run docs:docx-fidelity -- --matrix     verify claims and print the coverage matrix
 *   npm run docs:docx-fidelity -- --report md  also print a markdown report (--report json for JSON);
 *       add --report-out <file> to append it to a file (CI passes "$GITHUB_STEP_SUMMARY"); repeatable
 *
 * Runs against the built @b32nio/spindle-docs-core (run `npm run build` first).
 */
import { main, type CliConfig } from '../fidelity-core/cli';
import { loadBuiltApi } from './api';
import { CONSTRUCTS } from './constructs';
import { DOCX_FIXTURES_DIR, DOCX_MANIFEST_OPTIONS, MATRIX_END, MATRIX_START } from './index';
import { toDocxReport } from './report';
import { docxAdapter } from './run';
import type { EditorApi, FixtureResult } from './types';

export function docxCliConfig(api: EditorApi, fixturesDir = DOCX_FIXTURES_DIR): CliConfig<FixtureResult> {
  return {
    format: 'DOCX',
    updateCommand: 'npm run docs:docx-fidelity:update',
    fixturesDir,
    manifestOptions: DOCX_MANIFEST_OPTIONS,
    registry: CONSTRUCTS,
    matrixMarkers: { start: MATRIX_START, end: MATRIX_END },
    adapter: docxAdapter(api),
    // A manifest that disagrees with the importer (expectImportError) fails the run for DOCX.
    strictImport: true,
    toReport: toDocxReport,
  };
}

/* istanbul ignore next -- process bootstrap: runs only when the file is executed (tsx), not when Jest imports it */
if (require.main === module) {
  loadBuiltApi().then(
    (api) => main(docxCliConfig(api), process.argv.slice(2)),
    (e) => {
      console.error(e instanceof Error ? e.stack : e);
      process.exit(2);
    }
  );
}
