/**
 * DOCX fidelity corpus runner.
 *
 *   npm run docs:docx-fidelity                 compare every fixture with its golden snapshots (exit 1 on drift)
 *   npm run docs:docx-fidelity:update          rewrite the snapshots (and the README coverage matrix)
 *   npm run docs:docx-fidelity -- --only id    run a subset
 *   npm run docs:docx-fidelity -- --matrix     verify claims and print the coverage matrix
 *
 * Runs against the built @b32nio/spindle-docs-core (run `npm run build` first).
 */
import { readFileSync, writeFileSync } from 'node:fs';
import path from 'node:path';
import { loadManifest } from './manifest';
import { describeOutcome, runHarness } from './harness';
import { MATRIX_END, MATRIX_START, buildMatrix, injectMatrix, uncoveredConstructs, verifyClaims } from './matrix';
import { SNAPSHOT_KINDS, type EditorApi } from './types';

const repoRoot = path.resolve(__dirname, '..', '..');
const fixturesDir = path.join(repoRoot, 'fixtures', 'docx');

async function loadApi(): Promise<EditorApi> {
  const core = await import('@b32nio/spindle-docs-core');
  return {
    createDocument: () => new core.DocumentImpl('fixture', 'Untitled'),
    importFromDocx: (bytes, document) => core.importFromDocx(bytes, document as InstanceType<typeof core.DocumentImpl>),
    exportToDocx: (document) => core.exportToDocx(document as InstanceType<typeof core.DocumentImpl>),
    inspectDocx: (bytes) => core.inspectDocx(bytes),
  };
}

function parseArgs(argv: string[]) {
  const args = { update: false, matrix: false, only: [] as string[] };
  for (let i = 0; i < argv.length; i++) {
    if (argv[i] === '--update') args.update = true;
    else if (argv[i] === '--matrix') args.matrix = true;
    else if (argv[i] === '--only') args.only.push(...(argv[++i] ?? '').split(',').filter(Boolean));
    else throw new Error(`unknown argument: ${argv[i]}`);
  }
  return args;
}

async function main(): Promise<number> {
  const args = parseArgs(process.argv.slice(2));
  const manifest = loadManifest(fixturesDir);

  const falseClaims = await verifyClaims(fixturesDir);
  if (falseClaims.length) {
    console.error(`Manifest claims not backed by the files:\n- ${falseClaims.join('\n- ')}`);
    return 1;
  }

  const readme = path.join(fixturesDir, 'README.md');
  const matrix = buildMatrix([...manifest.fixtures].sort((a, b) => a.id.localeCompare(b.id)));
  if (args.matrix) {
    console.log(matrix);
    const gaps = uncoveredConstructs(manifest.fixtures);
    if (gaps.length) console.log(`\nConstructs with no fixture: ${gaps.join(', ')}`);
    return 0;
  }

  const report = await runHarness({ api: await loadApi(), fixturesDir, update: args.update, only: args.only });
  for (const o of report.fixtures) {
    const state = SNAPSHOT_KINDS.map((k) => `${k}:${o.status[k]}`).join(' ');
    console.log(`${describeOutcome(o).length ? 'FAIL' : 'ok  '} ${o.entry.id.padEnd(28)} ${state}`);
    for (const line of describeOutcome(o)) console.log(`       ${line}`);
  }
  for (const f of report.orphanGoldens) console.log(`FAIL orphan golden without a manifest entry: ${f}`);

  if (args.update) {
    const current = readFileSync(readme, 'utf8');
    if (current.includes(MATRIX_START)) writeFileSync(readme, injectMatrix(current, matrix));
    console.log('Snapshots and README coverage matrix updated. Review the diff before committing.');
    return 0;
  }
  if (!report.ok) {
    console.error('\nSnapshots differ from the current importer output. If the change is intended, run `npm run docs:docx-fidelity:update` and review the diff.');
    return 1;
  }
  return 0;
}

main().then(
  (code) => process.exit(code),
  (e) => {
    console.error(e instanceof Error ? e.stack : e);
    process.exit(2);
  }
);
