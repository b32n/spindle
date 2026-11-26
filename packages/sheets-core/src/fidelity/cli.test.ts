import { spawnSync } from 'node:child_process';
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { runCli, type CliIo } from '../../../../scripts/fidelity-core/cli';
import { xlsxCliConfig } from '../../../../scripts/xlsx-fidelity/cli';
import { XLSX_FIXTURES_DIR } from '../../../../scripts/xlsx-fidelity';
import { sourceApi } from './support';

const io = (): CliIo & { out: string[]; err: string[] } => {
  const out: string[] = [];
  const err: string[] = [];
  return { out, err, log: (l) => out.push(l), error: (l) => err.push(l) };
};

describe('xlsxCliConfig', () => {
  it('describes the XLSX run', () => {
    const config = xlsxCliConfig(sourceApi);
    expect(config).toMatchObject({
      format: 'XLSX',
      updateCommand: 'npm run sheets:xlsx-fidelity:update',
      fixturesDir: XLSX_FIXTURES_DIR,
    });
    expect(config.adapter.kinds).toEqual(['model', 'loss', 'roundtrip', 'ledger', 'values']);
    expect(config.strictImport).toBeUndefined();
  });

  it('runs a subset against the committed goldens and prints the report', async () => {
    const sink = io();
    const code = await runCli(
      xlsxCliConfig(sourceApi),
      ['--only', 'cell-types,openpyxl-image', '--report', 'md'],
      sink
    );
    expect(code).toBe(0);
    expect(sink.out[0]).toMatch(
      /^ok {3}cell-types\s+model:match loss:match roundtrip:match ledger:match values:match$/
    );
    expect(sink.out.join('\n')).toContain('## XLSX fidelity report');
    expect(sink.out.join('\n')).toContain('### Launch-bar scoreboard');
    expect(sink.out.join('\n')).toContain('`openpyxl-image` | **error**');
  });

  it('prints the coverage matrix', async () => {
    const sink = io();
    expect(await runCli(xlsxCliConfig(sourceApi), ['--matrix'], sink)).toBe(0);
    expect(sink.out.join('\n')).toContain('| Construct |');
  });

  it('rejects an unknown fixture id', async () => {
    await expect(runCli(xlsxCliConfig(sourceApi), ['--only', 'nope'], io())).rejects.toThrow(
      'unknown fixture id(s): nope'
    );
  });

  const python = spawnSync('python3', ['--version']).status === 0;
  (python || process.env.CI ? it : it.skip)(
    'wires --runtime to the run-time generator in the fixtures directory',
    async () => {
      const dir = mkdtempSync(path.join(tmpdir(), 'xlsx-cli-runtime-'));
      try {
        mkdirSync(path.join(dir, 'generators'));
        writeFileSync(path.join(dir, 'generators', 'build_runtime.py'), 'pass\n');
        const config = xlsxCliConfig(sourceApi, dir);
        expect(config.fixturesDir).toBe(dir);
        const sink = io();
        await config.runRuntime?.(
          [],
          { update: false, matrix: false, runtime: true, only: [], reports: [], timeoutSeconds: 5 },
          sink
        );
        expect(sink.out[0]).toContain('Run-time files');
      } finally {
        rmSync(dir, { recursive: true, force: true });
      }
    }
  );
});
