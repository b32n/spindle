import { spawnSync } from 'node:child_process';
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { runCli, type CliIo } from '../../../../scripts/fidelity-core/cli';
import { pptxCliConfig } from '../../../../scripts/pptx-fidelity/cli';
import { PPTX_FIXTURES_DIR } from '../../../../scripts/pptx-fidelity';
import { sourceApi } from './support';

const io = (): CliIo & { out: string[]; err: string[] } => {
  const out: string[] = [];
  const err: string[] = [];
  return { out, err, log: (l) => out.push(l), error: (l) => err.push(l) };
};

describe('pptxCliConfig', () => {
  it('describes the PPTX run', () => {
    const config = pptxCliConfig(sourceApi);
    expect(config).toMatchObject({
      format: 'PPTX',
      updateCommand: 'npm run slides:pptx-fidelity:update',
      fixturesDir: PPTX_FIXTURES_DIR,
    });
    expect(config.adapter.kinds).toEqual(['model', 'loss', 'roundtrip', 'ledger']);
    expect(config.strictImport).toBe(true);
  });

  it('runs a subset against the committed goldens and prints the report', async () => {
    const sink = io();
    const code = await runCli(pptxCliConfig(sourceApi), ['--only', 'kitchen-sink,edge-encrypted', '--report', 'md'], sink);
    expect(code).toBe(0);
    expect(sink.out[0]).toMatch(/^ok {3}kitchen-sink\s+model:match loss:match roundtrip:match ledger:match$/);
    const text = sink.out.join('\n');
    expect(text).toContain('## PPTX fidelity report');
    expect(text).toContain('### Launch-bar scoreboard');
    expect(text).toContain('`kitchen-sink`');
    expect(text).not.toContain('`kitchen-sink` | **error**');
    expect(text).toContain('| L13 | Slides |');
  });

  it('prints the coverage matrix', async () => {
    const sink = io();
    expect(await runCli(pptxCliConfig(sourceApi), ['--matrix'], sink)).toBe(0);
    expect(sink.out.join('\n')).toContain('| Construct |');
  });

  it('rejects an unknown fixture id', async () => {
    await expect(runCli(pptxCliConfig(sourceApi), ['--only', 'nope'], io())).rejects.toThrow('unknown fixture id(s): nope');
  });

  const python = spawnSync('python3', ['--version']).status === 0;
  (python || process.env.CI ? it : it.skip)('wires --runtime to the run-time generator in the fixtures directory', async () => {
    const dir = mkdtempSync(path.join(tmpdir(), 'pptx-cli-runtime-'));
    try {
      mkdirSync(path.join(dir, 'generators'));
      writeFileSync(path.join(dir, 'generators', 'build_runtime.py'), 'pass\n');
      const config = pptxCliConfig(sourceApi, dir);
      expect(config.fixturesDir).toBe(dir);
      const sink = io();
      await config.runRuntime?.([], { update: false, matrix: false, runtime: true, only: [], reports: [], timeoutSeconds: 5 }, sink);
      expect(sink.out[0]).toContain('Run-time files');
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });
});
