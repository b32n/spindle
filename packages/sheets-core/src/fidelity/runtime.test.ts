import { spawnSync } from 'node:child_process';
import { existsSync, mkdtempSync, readdirSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import type { CliIo } from '../../../../scripts/fidelity-core/cli';
import type { ManifestEntry } from '../../../../scripts/fidelity-core/types';
import type { XlsxEditorApi } from '../../../../scripts/xlsx-fidelity/api';
import {
  generateRuntimeFiles,
  runRuntime,
  runRuntimeFile,
} from '../../../../scripts/xlsx-fidelity/runtime';
import { exportToXlsx } from '../export/xlsx';
import { WorkbookImpl } from '../workbook';
import { sourceApi } from './support';

async function workbookBytes(rows: number): Promise<Uint8Array> {
  const wb = new WorkbookImpl('id', 'n');
  for (let r = 0; r < rows; r++) wb.setCellValue(undefined, r, 0, r);
  return exportToXlsx(wb);
}

const python = spawnSync('python3', ['--version']).status === 0;
const itPy = python || process.env.CI ? it : it.skip;

describe('runRuntimeFile', () => {
  it('imports a file and reports sizes, time and memory', async () => {
    const run = await runRuntimeFile(sourceApi, 'small', await workbookBytes(5), 30_000);
    expect(run).toMatchObject({
      id: 'small',
      importError: null,
      sheets: 1,
      cells: 5,
      lossTypes: 0,
    });
    expect(run.elapsedMs).toBeGreaterThanOrEqual(0);
    expect(run.rssMb).toBeGreaterThan(0);
  });

  it('records an import error, scrubbing the checkout path', async () => {
    const run = await runRuntimeFile(sourceApi, 'bad', new Uint8Array([1, 2, 3]), 30_000);
    expect(run.importError).toEqual(expect.any(String));
    expect(run).toMatchObject({ sheets: 0, cells: 0 });
    const throwing: XlsxEditorApi = {
      ...sourceApi,
      importFromXlsx: () => Promise.reject(`broke at ${process.cwd()}/x`),
    };
    expect((await runRuntimeFile(throwing, 'p', new Uint8Array(), 1000)).importError).toBe(
      'broke at <repo>/x'
    );
  });

  it('gives up on an import that does not finish', async () => {
    const hanging: XlsxEditorApi = {
      ...sourceApi,
      importFromXlsx: () => new Promise(() => undefined),
    };
    expect((await runRuntimeFile(hanging, 'slow', new Uint8Array(), 5)).importError).toBe(
      'import of slow did not finish within 5 ms'
    );
  });
});

describe('generateRuntimeFiles', () => {
  let dir: string;
  beforeEach(() => {
    dir = mkdtempSync(path.join(tmpdir(), 'xlsx-runtime-gen-'));
  });
  afterEach(() => rmSync(dir, { recursive: true, force: true }));

  itPy('runs the generator script with the output directory and ids', () => {
    const script = path.join(dir, 'gen.py');
    writeFileSync(
      script,
      'import sys, os\nout = sys.argv[sys.argv.index("--out") + 1]\nfor name in sys.argv[sys.argv.index("--out") + 2:]:\n    open(os.path.join(out, name + ".xlsx"), "w").write("x")\n'
    );
    generateRuntimeFiles(script, dir, ['one', 'two']);
    expect(readdirSync(dir).sort()).toEqual(['gen.py', 'one.xlsx', 'two.xlsx']);
  });

  itPy('throws with the script output when it fails', () => {
    const script = path.join(dir, 'fail.py');
    writeFileSync(script, 'import sys\nsys.stderr.write("it failed")\nsys.exit(3)\n');
    expect(() => generateRuntimeFiles(script, dir, [])).toThrow('generator failed (3): it failed');
  });

  itPy('throws with the spawn error when the script produces no output (timeout)', () => {
    const script = path.join(dir, 'slow.py');
    writeFileSync(script, 'import time\ntime.sleep(30)\n');
    expect(() => generateRuntimeFiles(script, dir, [], 50)).toThrow(
      /generator failed \(null\): Error: .*ETIMEDOUT/
    );
  });
});

describe('runRuntime', () => {
  const entry = (id: string, extra: Partial<ManifestEntry> = {}): ManifestEntry => ({
    id,
    filename: `${id}.xlsx`,
    title: id,
    source: { tool: 't', os: 'o' },
    license: 'l',
    constructs: [],
    notes: '',
    generatedAtRuntime: true,
    ...extra,
  });
  const args = {
    update: false,
    matrix: false,
    runtime: true,
    only: [],
    reports: [],
    timeoutSeconds: 30,
  };
  const io = (): CliIo & { out: string[] } => {
    const out: string[] = [];
    return { out, log: (l) => out.push(l), error: (l) => out.push(`ERR ${l}`) };
  };

  it('generates, imports each file, prints one line per file and removes the temporary directory', async () => {
    let seen: { script: string; dir: string; ids: string[] } | undefined;
    const small = await workbookBytes(3);
    const sink = io();
    await runRuntime(
      {
        api: sourceApi,
        fixturesDir: '/fixtures',
        generate: (script, dir, ids) => {
          seen = { script, dir, ids };
          writeFileSync(path.join(dir, 'ok.xlsx'), small);
          writeFileSync(path.join(dir, 'bad.xlsx'), new Uint8Array([1, 2]));
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
    expect(sink.out[1]).toMatch(
      /^ {5}ok\s+\d+ ms\s+rss\s+\d+ MB\s+loss types 0\s+1 sheets, 3 cells$/
    );
    expect(sink.out[2]).toMatch(/^ {5}bad\s.*import error: /);
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

  itPy('uses the real generator by default', async () => {
    const sink = io();
    // A private TMPDIR, so a runtime directory created by a suite running in parallel cannot be mistaken for ours.
    const previous = process.env.TMPDIR;
    const own = mkdtempSync(path.join(tmpdir(), 'xlsx-fidelity-test-'));
    process.env.TMPDIR = own;
    try {
      await runRuntime(
        { api: sourceApi, fixturesDir: path.resolve(__dirname, '../../../../fixtures/xlsx') },
        [entry('edge-huge-dimension', { generator: 'generators/build_runtime.py' })],
        args,
        sink
      );
      expect(sink.out[1]).toContain('edge-huge-dimension');
      expect(readdirSync(own).filter((n) => n.startsWith('xlsx-fidelity-runtime-'))).toEqual([]);
    } finally {
      if (previous === undefined) delete process.env.TMPDIR;
      else process.env.TMPDIR = previous;
      rmSync(own, { recursive: true, force: true });
    }
  });
});
