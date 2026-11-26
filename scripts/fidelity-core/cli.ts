import { appendFileSync, readFileSync, writeFileSync } from 'node:fs';
import path from 'node:path';
import { parseArgs, type CliArgs } from './args';
import { verifyClaims } from './claims';
import { describeOutcome, runHarness, type HarnessAdapter, type HarnessReport } from './harness';
import { loadManifest } from './manifest';
import { buildMatrix, injectMatrix, uncoveredConstructs } from './matrix';
import type { OoxmlParts } from './ooxml-package';
import { formatJson, formatMarkdown, type FidelityReport } from './report';
import type { Construct, ManifestEntry, ManifestOptions } from './types';

export interface CliIo {
  log(line: string): void;
  error(line: string): void;
}

export interface CliConfig<R> {
  /** Display name, e.g. "XLSX". */
  format: string;
  /** What to run to rewrite the goldens, for the failure message. */
  updateCommand: string;
  fixturesDir: string;
  manifestOptions: ManifestOptions;
  registry: Array<Construct<OoxmlParts>>;
  matrixMarkers: { start: string; end: string };
  adapter: HarnessAdapter<R>;
  strictImport?: boolean;
  toReport(run: HarnessReport<R>): FidelityReport;
  /** Run the manifest's `generatedAtRuntime` files (opt-in `--runtime`); never compared with goldens. */
  runRuntime?(entries: ManifestEntry[], args: CliArgs, io: CliIo): Promise<void>;
}

const consoleIo: CliIo = { log: (l) => console.log(l), error: (l) => console.error(l) };

/**
 * The runner shared by every format:
 *
 *   <command>                      compare every fixture with its golden snapshots (exit 1 on drift)
 *   <command> --update             rewrite the snapshots (and the README coverage matrix)
 *   <command> --only id[,id]       run a subset
 *   <command> --matrix             verify claims and print the coverage matrix
 *   <command> --report md|json     also print a report; add --report-out <file> to append/write it (repeatable)
 *   <command> --runtime            also run the files generated at run time (bomb, huge, perf), without goldens
 *   <command> --timeout <seconds>  per-file wall-clock limit (default 60)
 *
 * Returns the process exit code: 0 ok, 1 drift or false claims, 2 bad arguments or an unexpected failure.
 */
export async function runCli<R>(
  config: CliConfig<R>,
  argv: string[],
  io: CliIo = consoleIo
): Promise<number> {
  let args: CliArgs;
  try {
    args = parseArgs(argv);
  } catch (e) {
    io.error((e as Error).message); // parseArgs only throws Errors
    return 2;
  }
  const manifest = loadManifest(config.fixturesDir, config.manifestOptions);

  const falseClaims = await verifyClaims(config.fixturesDir, manifest, config.registry);
  if (falseClaims.length) {
    io.error(`Manifest claims not backed by the files:\n- ${falseClaims.join('\n- ')}`);
    return 1;
  }

  const matrix = buildMatrix(config.registry, manifest.fixtures);
  if (args.matrix) {
    io.log(matrix);
    const gaps = uncoveredConstructs(config.registry, manifest.fixtures);
    if (gaps.length) io.log(`\nConstructs with no fixture: ${gaps.join(', ')}`);
    return 0;
  }

  const run = await runHarness({
    adapter: config.adapter,
    fixturesDir: config.fixturesDir,
    manifestOptions: config.manifestOptions,
    update: args.update,
    only: args.only,
    strictImport: config.strictImport,
    timeoutMs: args.timeoutSeconds * 1000,
  });
  const kinds = config.adapter.kinds;
  for (const o of run.fixtures) {
    const lines = describeOutcome(o, kinds);
    io.log(
      `${lines.length ? 'FAIL' : 'ok  '} ${o.entry.id.padEnd(28)} ${kinds.map((k) => `${k}:${o.status[k]}`).join(' ')}`
    );
    for (const line of lines) io.log(`       ${line}`);
  }
  for (const f of run.orphanGoldens) io.log(`FAIL orphan golden without a manifest entry: ${f}`);

  const report = config.toReport(run);
  for (const r of args.reports) {
    const text = r.format === 'md' ? formatMarkdown(report) : formatJson(report);
    if (!r.out) io.log('\n' + text);
    else if (r.format === 'md')
      appendFileSync(r.out, text + '\n'); // $GITHUB_STEP_SUMMARY is append-only
    else writeFileSync(r.out, text);
  }

  if (args.runtime) {
    const entries = manifest.fixtures.filter((f) => f.generatedAtRuntime);
    if (config.runRuntime) await config.runRuntime(entries, args, io);
    else io.log('This format has no run-time files.');
  }

  if (args.update) {
    const readme = path.join(config.fixturesDir, 'README.md');
    const current = readFileSync(readme, 'utf8');
    if (current.includes(config.matrixMarkers.start))
      writeFileSync(
        readme,
        injectMatrix(current, matrix, config.matrixMarkers.start, config.matrixMarkers.end)
      );
    io.log('Snapshots and README coverage matrix updated. Review the diff before committing.');
    return 0;
  }
  if (!run.ok) {
    io.error(
      `\nSnapshots differ from the current importer output. If the change is intended, run \`${config.updateCommand}\` and review the diff.`
    );
    return 1;
  }
  return 0;
}

/** Process entry point: runs the CLI and exits; an unexpected exception is exit code 2 with the stack. */
export function main<R>(config: CliConfig<R>, argv: string[]): void {
  runCli(config, argv).then(
    (code) => process.exit(code),
    (e) => {
      console.error(e instanceof Error ? e.stack : e);
      process.exit(2);
    }
  );
}
