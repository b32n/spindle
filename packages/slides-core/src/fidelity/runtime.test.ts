import { spawnSync } from 'node:child_process';
import { existsSync, mkdtempSync, readdirSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import JSZip from 'jszip';
import type { CliIo } from '../../../../scripts/fidelity-core/cli';
import type { ManifestEntry } from '../../../../scripts/fidelity-core/types';
import type { PptxEditorApi } from '../../../../scripts/pptx-fidelity/api';
import {
  generateRuntimeFiles,
  runRuntime,
  runRuntimeFile,
} from '../../../../scripts/pptx-fidelity/runtime';
import { detectOnlyApi, sourceApi } from './support';

const python = spawnSync('python3', ['--version']).status === 0;
const itPy = python || process.env.CI ? it : it.skip;

async function deckBytes(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('ppt/presentation.xml', '<p:presentation xmlns:p="x"/>');
  zip.file('ppt/slides/slide1.xml', '<p:sld xmlns:p="x" show="0"/>');
  return zip.generateAsync({ type: 'uint8array' });
}

describe('runRuntimeFile', () => {
  it('scans a file and reports time, memory and the loss counts; "no importer" without one', async () => {
    const run = await runRuntimeFile(detectOnlyApi, 'small', await deckBytes(), 30_000);
    expect(run).toMatchObject({ id: 'small', importError: 'no importer', lossTypes: 1, lossCount: 1 });
    expect(run.elapsedMs).toBeGreaterThanOrEqual(0);
    expect(run.rssMb).toBeGreaterThan(0);
  });

  it('also imports when an importer exists, and records its failure with the checkout path removed', async () => {
    const withImporter: PptxEditorApi = { ...sourceApi, importPptx: async () => ({ warnings: [], presentation: {} }) };
    expect((await runRuntimeFile(withImporter, 'p', await deckBytes(), 30_000)).importError).toBeNull();
    const throwing: PptxEditorApi = { ...withImporter, importPptx: () => Promise.reject(`broke at ${process.cwd()}/x`) };
    expect((await runRuntimeFile(throwing, 'p', new Uint8Array(), 1000)).importError).toBe('broke at <repo>/x');
    const error: PptxEditorApi = { ...withImporter, importPptx: () => Promise.reject(new Error('plain')) };
    expect((await runRuntimeFile(error, 'p', new Uint8Array(), 1000)).importError).toBe('plain');
  });

  it('gives up on a scan or an import that does not finish', async () => {
    const hangingScan: PptxEditorApi = { inspectPptx: () => new Promise(() => undefined) };
    expect((await runRuntimeFile(hangingScan, 'slow', new Uint8Array(), 5)).importError).toBe('scan of slow did not finish within 5 ms');
    const hangingImport: PptxEditorApi = { ...sourceApi, importPptx: () => new Promise(() => undefined) };
    expect((await runRuntimeFile(hangingImport, 'slow', new Uint8Array(), 50)).importError).toBe('import of slow did not finish within 50 ms');
  });
});

describe('generateRuntimeFiles', () => {
  let dir: string;
  beforeEach(() => {
    dir = mkdtempSync(path.join(tmpdir(), 'pptx-runtime-gen-'));
  });
  afterEach(() => rmSync(dir, { recursive: true, force: true }));

  itPy('runs the generator script with the output directory and ids', () => {
    const script = path.join(dir, 'gen.py');
    writeFileSync(
      script,
      'import sys, os\nout = sys.argv[sys.argv.index("--out") + 1]\nfor name in sys.argv[sys.argv.index("--out") + 2:]:\n    open(os.path.join(out, name + ".pptx"), "w").write("x")\n'
    );
    generateRuntimeFiles(script, dir, ['one', 'two']);
    expect(readdirSync(dir).sort()).toEqual(['gen.py', 'one.pptx', 'two.pptx']);
  });

  itPy('throws with the script output when it fails', () => {
    const script = path.join(dir, 'fail.py');
    writeFileSync(script, 'import sys\nsys.stderr.write("it failed")\nsys.exit(3)\n');
    expect(() => generateRuntimeFiles(script, dir, [])).toThrow('generator failed (3): it failed');
  });

  itPy('throws with the spawn error when the script produces no output (timeout)', () => {
    const script = path.join(dir, 'slow.py');
    writeFileSync(script, 'import time\ntime.sleep(30)\n');
    expect(() => generateRuntimeFiles(script, dir, [], 50)).toThrow(/generator failed \(null\): Error: .*ETIMEDOUT/);
  });
});

describe('runRuntime', () => {
  const entry = (id: string, extra: Partial<ManifestEntry> = {}): ManifestEntry => ({
    id,
    filename: `${id}.pptx`,
    title: id,
    source: { tool: 't', os: 'o' },
    license: 'l',
    constructs: [],
    notes: '',
    generatedAtRuntime: true,
    ...extra,
  });
  const args = { update: false, matrix: false, runtime: true, only: [], reports: [], timeoutSeconds: 30 };
  const io = (): CliIo & { out: string[] } => {
    const out: string[] = [];
    return { out, log: (l) => out.push(l), error: (l) => out.push(`ERR ${l}`) };
  };

  it('generates, scans each file, prints one line per file and removes the temporary directory', async () => {
    let seen: { script: string; dir: string; ids: string[] } | undefined;
    const deck = await deckBytes();
    const sink = io();
    await runRuntime(
      {
        api: detectOnlyApi,
        fixturesDir: '/fixtures',
        generate: (script, dir, ids) => {
          seen = { script, dir, ids };
          writeFileSync(path.join(dir, 'ok.pptx'), deck);
          writeFileSync(path.join(dir, 'bad.pptx'), new Uint8Array([1, 2]));
        },
      },
      [entry('ok', { generator: 'generators/build_runtime.py' }), entry('bad')],
      args,
      sink
    );
    expect(seen?.script).toBe(path.join('/fixtures', 'generators/build_runtime.py'));
    expect(seen?.ids).toEqual(['ok', 'bad']);
    expect(existsSync(seen?.dir ?? '')).toBe(false);
    expect(sink.out[0]).toContain('Run-time files');
    expect(sink.out[1]).toMatch(/^ {5}ok\s+\d+ ms\s+rss\s+\d+ MB\s+loss types 1 \(1 objects\)\s+import: no importer$/);
    expect(sink.out[2]).toMatch(/^ {5}bad\s.*loss types 0 \(0 objects\)\s+import: no importer$/);
  });

  it('prints "ok" for a file the importer takes, once there is one', async () => {
    const deck = await deckBytes();
    const sink = io();
    await runRuntime(
      {
        api: { ...sourceApi, importPptx: async () => ({ warnings: [], presentation: {} }) },
        fixturesDir: '/fixtures',
        generate: (_s, dir) => writeFileSync(path.join(dir, 'ok.pptx'), deck),
      },
      [entry('ok')],
      args,
      sink
    );
    expect(sink.out[1]).toMatch(/import: ok$/);
  });

  it('falls back to the default generator path and still cleans up when generation fails', async () => {
    let dirSeen = '';
    await expect(
      runRuntime(
        {
          api: sourceApi,
          fixturesDir: '/fixtures',
          generate: (script, dir) => {
            dirSeen = dir;
            expect(script).toBe(path.join('/fixtures', 'generators/build_runtime.py'));
            throw new Error('no python');
          },
        },
        [entry('x')],
        args,
        io()
      )
    ).rejects.toThrow('no python');
    expect(existsSync(dirSeen)).toBe(false);
  });

  itPy('uses the real generator by default and scans and imports the 1000-slide deck', async () => {
    const sink = io();
    // A private TMPDIR, so a runtime directory created by a suite running in parallel cannot be mistaken for ours.
    const previous = process.env.TMPDIR;
    const own = mkdtempSync(path.join(tmpdir(), 'pptx-fidelity-test-'));
    process.env.TMPDIR = own;
    try {
      await runRuntime(
        { api: sourceApi, fixturesDir: path.resolve(__dirname, '../../../../fixtures/pptx') },
        [entry('edge-1000-slides', { generator: 'generators/build_runtime.py' })],
        { ...args, timeoutSeconds: 120 },
        sink
      );
      expect(sink.out[1]).toContain('edge-1000-slides');
      expect(sink.out[1]).toMatch(/import: ok$/);
      expect(readdirSync(own).filter((n) => n.startsWith('pptx-fidelity-runtime-'))).toEqual([]);
    } finally {
      if (previous === undefined) delete process.env.TMPDIR;
      else process.env.TMPDIR = previous;
      rmSync(own, { recursive: true, force: true });
    }
  }, 180_000);
});
