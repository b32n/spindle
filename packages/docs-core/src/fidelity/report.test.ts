import { formatJson, formatMarkdown } from '../../../../scripts/fidelity-core/report';
import type { HarnessReport } from '../../../../scripts/fidelity-core/harness';
import { toDocxReport } from '../../../../scripts/docx-fidelity/report';
import type { FixtureResult } from '../../../../scripts/docx-fidelity/types';
import { diffModels } from '../../../../scripts/docx-fidelity/diff';

const run = (over: Partial<HarnessReport<FixtureResult>> = {}): HarnessReport<FixtureResult> => ({
  ok: true,
  orphanGoldens: [],
  fixtures: [
    {
      entry: { id: 'one', filename: 'one.docx', title: 'One', source: { tool: 'tool', os: 'os' }, license: 'l', constructs: ['headings'], notes: 'n' },
      result: {
        loss: { format: 'docx', items: [{ id: 'docx.comments', label: 'c', count: 1, effect: 'e' }] },
        model: { importError: null, warnings: ['w'], document: {} },
        roundtrip: {
          status: 'ok',
          exportWarnings: [],
          reimportWarnings: [],
          diff: diffModels(
            { sections: [{ blocks: [{ type: 'heading', content: [{ type: 'text', text: 'T' }] }, { type: 'paragraph', content: [{ type: 'text', text: 'body' }] }] }] },
            { sections: [{ blocks: [{ type: 'paragraph', content: [{ type: 'text', text: 'body' }] }] }] }
          ),
        },
      },
      status: { model: 'match', loss: 'match', roundtrip: 'created' },
      problems: [],
      importExpectation: 'as-expected',
    },
  ],
  ...over,
});

describe('toDocxReport', () => {
  it('maps a harness run onto the neutral report', () => {
    const r = toDocxReport(run());
    expect(r).toMatchObject({
      format: 'DOCX',
      updateCommand: 'npm run docs:docx-fidelity:update',
      ok: true,
      kinds: ['model', 'loss', 'roundtrip'],
      orphanGoldens: [],
      scoreboard: [],
    });
    expect(r.files[0]).toMatchObject({ id: 'one', title: 'One', producer: 'tool', warnings: ['w'], importError: null, importExpectation: 'as-expected', snapshots: { roundtrip: 'created' } });
    expect(r.files[0].loss).toHaveLength(1);
    expect(r.files[0].ledger).toBeUndefined();
    expect(r.files[0].values).toBeUndefined();
  });

  it('carries failures, problems and orphans through', () => {
    const base = run();
    const failed = toDocxReport({ ...base, ok: false, orphanGoldens: ['gone.model.json'], fixtures: [{ ...base.fixtures[0], problems: ['boom'], importExpectation: 'unexpected-error' }] });
    expect(failed).toMatchObject({ ok: false, orphanGoldens: ['gone.model.json'] });
    expect(failed.files[0]).toMatchObject({ problems: ['boom'], importExpectation: 'unexpected-error' });
  });

  it('renders through the shared markdown and JSON formatters', () => {
    const report = toDocxReport(run());
    const md = formatMarkdown(report);
    expect(md).toContain('## DOCX fidelity report');
    expect(md).toContain('| `one` | ok | 1 | 1 (1 type) | 1 difference | match | match | match |');
    expect(md).toContain('Before: 1 sections, 1 heading, 1 paragraph, 5 textChars.');
    expect(md).not.toContain('Launch-bar scoreboard');
    const json = formatJson(report);
    expect(json.endsWith('\n')).toBe(true);
    expect(JSON.parse(json).files[0].id).toBe('one');
  });
});
