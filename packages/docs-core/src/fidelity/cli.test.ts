import { runCli, type CliIo } from '../../../../scripts/fidelity-core/cli';
import { docxCliConfig } from '../../../../scripts/docx-fidelity/cli';
import { DOCX_FIXTURES_DIR } from '../../../../scripts/docx-fidelity';
import { sourceApi } from './support';

const io = (): CliIo & { out: string[]; err: string[] } => {
  const out: string[] = [];
  const err: string[] = [];
  return { out, err, log: (l) => out.push(l), error: (l) => err.push(l) };
};

describe('docxCliConfig', () => {
  it('describes the DOCX run: three kinds and a strict import expectation', () => {
    const config = docxCliConfig(sourceApi);
    expect(config).toMatchObject({ format: 'DOCX', updateCommand: 'npm run docs:docx-fidelity:update', fixturesDir: DOCX_FIXTURES_DIR, strictImport: true });
    expect(config.adapter.kinds).toEqual(['model', 'loss', 'roundtrip']);
    expect(config.runRuntime).toBeUndefined();
  });

  it('runs a subset against the committed goldens and prints the report', async () => {
    const sink = io();
    const code = await runCli(docxCliConfig(sourceApi), ['--only', 'charts-objects,edge-no-body', '--report', 'md'], sink);
    expect(code).toBe(0);
    expect(sink.out[0]).toMatch(/^ok {3}charts-objects\s+model:match loss:match roundtrip:match$/);
    const text = sink.out.join('\n');
    expect(text).toContain('## DOCX fidelity report');
    expect(text).toContain('`edge-no-body` | error (expected)');
    expect(text).not.toContain('Launch-bar scoreboard');
  });

  it('prints the coverage matrix', async () => {
    const sink = io();
    expect(await runCli(docxCliConfig(sourceApi), ['--matrix'], sink)).toBe(0);
    expect(sink.out.join('\n')).toContain('| Construct |');
  });
});
