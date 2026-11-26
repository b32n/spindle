import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import JSZip from 'jszip';
import { main, runCli, type CliConfig, type CliIo } from '../../../../../scripts/fidelity-core/cli';
import type { HarnessAdapter } from '../../../../../scripts/fidelity-core/harness';
import type { HarnessReport } from '../../../../../scripts/fidelity-core/harness';
import type { FidelityReport } from '../../../../../scripts/fidelity-core/report';
import type { Manifest } from '../../../../../scripts/fidelity-core/types';

/** A minimal report mapping for the CLI tests (the adapter result is just a size). */
export function toFidelityReportForTest(run: HarnessReport<{ n: number }>): FidelityReport {
  return {
    format: 'TEST',
    updateCommand: 'npm run test:update',
    ok: run.ok,
    kinds: ['model'],
    orphanGoldens: run.orphanGoldens,
    scoreboard: [],
    files: run.fixtures.map((o) => ({
      id: o.entry.id,
      title: o.entry.title,
      producer: o.entry.source.tool,
      snapshots: o.status,
      problems: o.problems,
      importError: null,
      importExpectation: o.importExpectation,
      warnings: [],
      loss: [],
      roundtrip: { status: 'ok', exportWarnings: [], reimportWarnings: [], diff: null },
    })),
  };
}

const START = '<!-- start -->';
const END = '<!-- end -->';

interface R {
  n: number;
}
const adapter: HarnessAdapter<R> = {
  kinds: ['model'],
  run: async (bytes) => ({ n: bytes.length }),
  crashed: () => ({ n: -1 }),
  snapshot: (r) => r,
  importError: () => null,
};

let dir: string;
let io: CliIo & { out: string[]; err: string[] };

async function zipBytes(): Promise<Buffer> {
  const zip = new JSZip();
  zip.file('a.xml', '<a/>');
  return zip.generateAsync({ type: 'nodebuffer' });
}

beforeEach(async () => {
  dir = mkdtempSync(path.join(tmpdir(), 'fidelity-cli-'));
  const manifest: Manifest = {
    version: 1,
    fixtures: [
      {
        id: 'one',
        filename: 'one.pkg',
        title: 'One',
        source: { tool: 't', os: 'linux' },
        license: 'l',
        constructs: ['has.a'],
        notes: 'n',
      },
      {
        id: 'rt',
        filename: 'rt.pkg',
        title: 'Runtime',
        source: { tool: 't', os: 'linux' },
        license: 'l',
        constructs: ['has.a'],
        notes: 'n',
        generatedAtRuntime: true,
        generator: 'g.py',
      },
    ],
  };
  writeFileSync(path.join(dir, 'manifest.json'), JSON.stringify(manifest));
  writeFileSync(path.join(dir, 'one.pkg'), await zipBytes());
  writeFileSync(path.join(dir, 'README.md'), `# T\n${START}\nold\n${END}\n`);
  const out: string[] = [];
  const err: string[] = [];
  io = { out, err, log: (l) => out.push(l), error: (l) => err.push(l) };
});
afterEach(() => rmSync(dir, { recursive: true, force: true }));

const config = (extra: Partial<CliConfig<R>> = {}): CliConfig<R> => ({
  format: 'TEST',
  updateCommand: 'npm run test:update',
  fixturesDir: dir,
  manifestOptions: { extensions: ['.pkg'], constructIds: ['has.a', 'has.b'], label: 'test' },
  registry: [
    { id: 'has.a', label: 'A', count: (p) => (p.names.includes('a.xml') ? 1 : 0) },
    { id: 'has.b', label: 'B', count: (p) => (p.names.includes('b.xml') ? 1 : 0) },
  ],
  matrixMarkers: { start: START, end: END },
  adapter,
  toReport: toFidelityReportForTest,
  ...extra,
});

describe('runCli', () => {
  it('rejects bad arguments with exit code 2', async () => {
    expect(await runCli(config(), ['--nope'], io)).toBe(2);
    expect(io.err).toEqual(['unknown argument: --nope']);
    expect(await runCli(config(), ['--timeout', 'x'], io)).toBe(2);
  });

  it('fails when a manifest claim is not backed by the file', async () => {
    const manifest = JSON.parse(readFileSync(path.join(dir, 'manifest.json'), 'utf8')) as Manifest;
    manifest.fixtures[0].constructs = ['has.b'];
    writeFileSync(path.join(dir, 'manifest.json'), JSON.stringify(manifest));
    expect(await runCli(config(), [], io)).toBe(1);
    expect(io.err[0]).toContain('one: claims "has.b" but the file does not contain it');
  });

  it('prints the coverage matrix and the constructs without a fixture', async () => {
    expect(await runCli(config(), ['--matrix'], io)).toBe(0);
    expect(io.out.join('\n')).toContain('| Construct |');
    expect(io.out.join('\n')).toContain('Constructs with no fixture: has.b');
    const all = config({ registry: [{ id: 'has.a', label: 'A', count: () => 1 }] });
    io.out.length = 0;
    expect(await runCli(all, ['--matrix'], io)).toBe(0);
    expect(io.out.join('\n')).not.toContain('Constructs with no fixture');
  });

  it('fails on missing goldens, then writes them (and the matrix) with --update, then passes', async () => {
    expect(await runCli(config(), [], io)).toBe(1);
    expect(io.out[0]).toMatch(/^FAIL one\s+model:missing/);
    expect(io.out[1]).toContain('one: model snapshot missing');
    expect(io.err.join('\n')).toContain('run `npm run test:update`');

    io.out.length = 0;
    expect(await runCli(config(), ['--update'], io)).toBe(0);
    expect(io.out.join('\n')).toContain('Snapshots and README coverage matrix updated');
    expect(readFileSync(path.join(dir, 'README.md'), 'utf8')).toContain('| `one` | One | t |');
    expect(existsSync(path.join(dir, 'goldens', 'one.model.json'))).toBe(true);

    io.out.length = 0;
    expect(await runCli(config(), [], io)).toBe(0);
    expect(io.out[0]).toMatch(/^ok {3}one\s+model:match/);
  });

  it('lists orphan goldens', async () => {
    await runCli(config(), ['--update'], io);
    writeFileSync(path.join(dir, 'goldens', 'gone.model.json'), '{}');
    io.out.length = 0;
    expect(await runCli(config(), [], io)).toBe(1);
    expect(io.out.join('\n')).toContain(
      'FAIL orphan golden without a manifest entry: gone.model.json'
    );
  });

  it('leaves a README without markers alone', async () => {
    writeFileSync(path.join(dir, 'README.md'), '# no markers\n');
    expect(await runCli(config(), ['--update'], io)).toBe(0);
    expect(readFileSync(path.join(dir, 'README.md'), 'utf8')).toBe('# no markers\n');
  });

  it('writes reports to stdout, appends markdown and overwrites JSON', async () => {
    await runCli(config(), ['--update'], io);
    io.out.length = 0;
    const md = path.join(dir, 'summary.md');
    const json = path.join(dir, 'report.json');
    writeFileSync(md, 'existing\n');
    expect(
      await runCli(
        config(),
        [
          '--only',
          'one',
          '--report',
          'md',
          '--report',
          'json',
          '--report',
          'md',
          '--report-out',
          md,
          '--report',
          'json',
          '--report-out',
          json,
        ],
        io
      )
    ).toBe(0);
    expect(io.out.join('\n')).toContain('## TEST fidelity report');
    expect(io.out.join('\n')).toContain('"format": "TEST"');
    expect(readFileSync(md, 'utf8').startsWith('existing\n## TEST fidelity report')).toBe(true);
    expect(JSON.parse(readFileSync(json, 'utf8')).files[0].id).toBe('one');
  });

  it('runs the run-time files only with --runtime, through the format hook', async () => {
    await runCli(config(), ['--update'], io);
    const seen: string[] = [];
    const withHook = config({
      runRuntime: async (entries, args, hookIo) => {
        seen.push(...entries.map((e) => e.id), String(args.timeoutSeconds));
        hookIo.log('hook ran');
      },
    });
    expect(await runCli(withHook, [], io)).toBe(0);
    expect(seen).toEqual([]);
    expect(await runCli(withHook, ['--runtime', '--timeout', '9'], io)).toBe(0);
    expect(seen).toEqual(['rt', '9']);
    expect(io.out).toContain('hook ran');
    expect(await runCli(config(), ['--runtime'], io)).toBe(0);
    expect(io.out).toContain('This format has no run-time files.');
  });

  it('logs to the console by default', async () => {
    const log = jest.spyOn(console, 'log').mockImplementation(() => undefined);
    const error = jest.spyOn(console, 'error').mockImplementation(() => undefined);
    try {
      expect(await runCli(config(), ['--bad'])).toBe(2);
      expect(error).toHaveBeenCalledWith('unknown argument: --bad');
      expect(await runCli(config(), ['--update'])).toBe(0);
      expect(log).toHaveBeenCalled();
    } finally {
      log.mockRestore();
      error.mockRestore();
    }
  });
});

describe('main', () => {
  it('exits with the code the run returns', async () => {
    const exit = jest.spyOn(process, 'exit').mockImplementation((() => undefined) as never);
    const error = jest.spyOn(console, 'error').mockImplementation(() => undefined);
    try {
      main(config(), ['--bad']);
      await new Promise((r) => setTimeout(r, 20));
      expect(exit).toHaveBeenCalledWith(2);
    } finally {
      exit.mockRestore();
      error.mockRestore();
    }
  });

  it('prints a non-Error exception as is', async () => {
    const exit = jest.spyOn(process, 'exit').mockImplementation((() => undefined) as never);
    const error = jest.spyOn(console, 'error').mockImplementation(() => undefined);
    const log = jest.spyOn(console, 'log').mockImplementation(() => undefined);
    try {
      main(
        config({
          toReport: () => {
            throw 'plain string';
          },
        }),
        ['--update']
      );
      await new Promise((r) => setTimeout(r, 50));
      expect(error).toHaveBeenCalledWith('plain string');
      expect(exit).toHaveBeenCalledWith(2);
    } finally {
      exit.mockRestore();
      error.mockRestore();
      log.mockRestore();
    }
  });

  it('prints the stack and exits 2 when the run throws', async () => {
    const exit = jest.spyOn(process, 'exit').mockImplementation((() => undefined) as never);
    const error = jest.spyOn(console, 'error').mockImplementation(() => undefined);
    try {
      writeFileSync(path.join(dir, 'manifest.json'), '{ not json');
      main(config(), []);
      await new Promise((r) => setTimeout(r, 20));
      expect(exit).toHaveBeenCalledWith(2);
      expect(String(error.mock.calls[0][0])).toContain('Invalid test manifest');
    } finally {
      exit.mockRestore();
      error.mockRestore();
    }
  });
});
