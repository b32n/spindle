import { parseArgs } from '../../../../scripts/docx-fidelity/args';
import type { HarnessReport } from '../../../../scripts/docx-fidelity/harness';
import { formatJson, formatMarkdown, mdEscape, toFidelityReport, type FidelityReport, type ReportFile } from '../../../../scripts/docx-fidelity/report';
import type { DiffReport, RoundtripSnapshot } from '../../../../scripts/docx-fidelity/types';

const counts = (blocks: Record<string, number>, textChars = 10) => ({ sections: 1, blocks, textChars });
const diff = (over: Partial<DiffReport> = {}): DiffReport => ({
  total: 0,
  added: 0,
  removed: 0,
  changed: 0,
  byField: {},
  before: counts({ paragraph: 2 }),
  after: counts({ paragraph: 2 }),
  entries: [],
  truncated: false,
  ...over,
});
const rt = (over: Partial<RoundtripSnapshot> = {}): RoundtripSnapshot => ({ status: 'ok', exportWarnings: [], reimportWarnings: [], diff: diff(), ...over });
const file = (id: string, over: Partial<ReportFile> = {}): ReportFile => ({
  id,
  title: `Title of ${id}`,
  producer: 'test tool',
  snapshots: { model: 'match', loss: 'match', roundtrip: 'match' },
  problems: [],
  importError: null,
  warnings: [],
  loss: [],
  roundtrip: rt(),
  ...over,
});
const report = (files: ReportFile[], over: Partial<FidelityReport> = {}): FidelityReport => ({ format: 'DOCX', ok: true, orphanGoldens: [], files, ...over });

describe('formatMarkdown', () => {
  it('summarises a passing run with losses as recorded behaviour, not failures', () => {
    const md = formatMarkdown(
      report([
        file('alpha', {
          warnings: ['w1'],
          loss: [
            { id: 'docx.comments', label: 'comment(s)', count: 3, effect: 'not imported' },
            { id: 'docx.fields', label: 'field(s)', count: 5, effect: 'dropped' },
          ],
          roundtrip: rt({ diff: diff({ total: 2, changed: 2, byField: { 'changed a.b': 2 }, entries: [{ path: 'a.b', kind: 'changed', before: '1', after: '2' }] }) }),
        }),
        file('beta'),
      ])
    );
    expect(md).toContain('## DOCX fidelity report');
    expect(md).toContain('**Result: all snapshots match their goldens.**');
    expect(md).toContain('2 files · 6/6 snapshots match · 0 import errors · 1 file with round-trip differences · 8 loss occurrences in 2 loss types');
    expect(md).toContain('| `alpha` | ok | 1 | 8 (2 types) | 2 differences | match | match | match |');
    expect(md).toContain('| `beta` | ok | 0 | 0 | clean | match | match | match |');
    expect(md).toContain('| `docx.fields` | 1 | 5 | field(s): dropped |');
    expect(md.indexOf('docx.fields')).toBeLessThan(md.indexOf('docx.comments')); // most frequent first
    expect(md).toContain('Differences: **2** (0 added, 0 removed, 2 changed).');
    expect(md).toContain('- changed `a.b`: 1 → 2');
    expect(md).toContain('No differences (1 section, 2 blocks).');
    expect(md).toContain('Loss report: nothing reported.');
  });

  it('reports a failed run with the failing snapshots first and open', () => {
    const md = formatMarkdown(
      report([file('good'), file('bad', { snapshots: { model: 'match', loss: 'mismatch', roundtrip: 'missing' }, problems: ['importFromDocx threw unexpectedly: boom'], importError: 'boom', roundtrip: rt({ status: 'skipped-import-error', diff: null }) })], {
        ok: false,
        orphanGoldens: ['gone.model.json'],
      })
    );
    expect(md).toContain('**Result: FAILED.** 2 snapshots differ from the goldens, 1 problem, 1 orphaned golden.');
    expect(md).toContain('npm run docs:docx-fidelity:update');
    expect(md).toContain('Orphaned goldens (no manifest entry): `gone.model.json`');
    expect(md).toContain('| `bad` | **error** | 0 | 0 | skipped (import error) | match | **mismatch** | **missing** |');
    expect(md).toContain('- **Problem:** importFromDocx threw unexpectedly: boom');
    expect(md).toContain('**Import error:** boom');
    expect(md.indexOf('<summary><code>bad</code>')).toBeLessThan(md.indexOf('<summary><code>good</code>'));
    expect(md).toContain('<details open>');
    expect(md).toContain('<details>');
    expect(md).toContain('1 import error');
  });

  it('shows round-trip failures and warnings', () => {
    const md = formatMarkdown(
      report([
        file('rtfail', { roundtrip: rt({ status: 'reimport-error', error: 'cannot read', diff: null, exportWarnings: ['export warned'], reimportWarnings: ['re-import warned'] }) }),
        file('exportfail', { roundtrip: rt({ status: 'export-error', error: 'no export', diff: null }) }),
      ])
    );
    expect(md).toContain('| `rtfail` | ok | 0 | 0 | **reimport-error** |');
    expect(md).toContain('2 round-trip failures');
    expect(md).toContain('Round trip (import → export → import): **reimport-error**: cannot read');
    expect(md).toContain('- export warning: export warned');
    expect(md).toContain('- re-import warning: re-import warned');
    expect(formatMarkdown(report([file('one', { roundtrip: rt({ status: 'export-error', diff: null }) })]))).toContain('1 round-trip failure');
  });

  it('lists block counts, added/removed entries, top fields and truncation notes', () => {
    const entries = Array.from({ length: 12 }, (_, i) => ({ path: `p${i}`, kind: (['added', 'removed', 'changed'] as const)[i % 3], before: 'b', after: 'a' }));
    const byField = Object.fromEntries(Array.from({ length: 15 }, (_, i) => [`changed f${String(i).padStart(2, '0')}`, 15 - i]));
    const md = formatMarkdown(
      report([file('big', { warnings: Array.from({ length: 13 }, (_, i) => `warning ${i}`), roundtrip: rt({ diff: diff({ total: 12, added: 4, removed: 4, changed: 4, entries, byField, before: counts({ heading: 1, paragraph: 3 }, 40), after: counts({ paragraph: 3 }, 30) }) }) })]),
      { maxWarnings: 2, maxDiffEntries: 3, maxFields: 4 }
    );
    expect(md).toContain('Before: 1 section, 1 heading, 3 paragraph, 40 characters.');
    expect(md).toContain('After: 1 section, 3 paragraph, 30 characters.');
    expect(md).toContain('- warning 0');
    expect(md).not.toContain('- warning 2');
    expect(md).toContain('…and 11 more');
    expect(md).toContain('…and 11 more fields.');
    expect(md).toContain('First 3 differences:');
    expect(md).toContain('- added `p0`: a');
    expect(md).toContain('- removed `p1`: b');
    expect(md).toContain('- changed `p2`: b → a');
    expect(md).not.toContain('`p3`');
    expect(md.indexOf('f00')).toBeLessThan(md.indexOf('f01')); // biggest first
    expect(md).not.toContain('f04');
  });

  it('describes a document with no blocks and a root-level difference', () => {
    const md = formatMarkdown(report([file('empty', { roundtrip: rt({ diff: diff({ total: 1, changed: 1, entries: [{ path: '', kind: 'changed', before: '1', after: '2' }], before: counts({}), after: counts({}) }) }) })]));
    expect(md).toContain('no blocks');
    expect(md).toContain('`(root)`');
  });

  it('keeps the summary under the size limit by omitting details, never the overview', () => {
    const files = Array.from({ length: 30 }, (_, i) => file(`f${String(i).padStart(2, '0')}`, { warnings: Array.from({ length: 10 }, (_, k) => `a fairly long warning text number ${k} `.repeat(4)) }));
    const md = formatMarkdown(report(files), { maxChars: 6000 });
    expect(md.length).toBeLessThan(8000);
    expect(md).toContain('| `f29` |');
    expect(md).toMatch(/Details omitted to stay under the summary size limit: `f\d\d`/);
  });

  it('escapes markdown and HTML in free text', () => {
    const md = formatMarkdown(report([file('x', { warnings: ['a | b <script>`tick`\nnext'], title: 'A <b> title' })]));
    expect(md).toContain('- a \\| b &lt;script>\'tick\' next');
    expect(md).toContain('A &lt;b> title');
    expect(mdEscape('back\\slash')).toBe('back\\\\slash');
  });

  it('handles an empty run', () => {
    const md = formatMarkdown(report([]));
    expect(md).toContain('0 files · 0/0 snapshots match');
    expect(md).not.toContain('Loss by type');
  });
});

describe('formatJson / toFidelityReport', () => {
  const run: HarnessReport = {
    ok: true,
    orphanGoldens: [],
    fixtures: [
      {
        entry: { id: 'one', filename: 'one.docx', title: 'One', source: { tool: 'tool', os: 'os' }, license: 'l', constructs: ['headings'], notes: 'n' },
        result: {
          loss: { format: 'docx', items: [{ id: 'docx.comments', label: 'c', count: 1, effect: 'e' }] },
          model: { importError: null, warnings: ['w'], document: {} },
          roundtrip: rt(),
        },
        status: { model: 'match', loss: 'match', roundtrip: 'created' },
        problems: [],
      },
    ],
  };

  it('maps a harness run onto the neutral report', () => {
    const r = toFidelityReport('DOCX', run);
    expect(r).toMatchObject({ format: 'DOCX', ok: true, orphanGoldens: [] });
    expect(r.files[0]).toMatchObject({ id: 'one', title: 'One', producer: 'tool', warnings: ['w'], importError: null, snapshots: { roundtrip: 'created' } });
    expect(r.files[0].loss).toHaveLength(1);
  });

  it('serialises with sorted keys and a trailing newline', () => {
    const json = formatJson(toFidelityReport('DOCX', run));
    expect(json.endsWith('\n')).toBe(true);
    expect(JSON.parse(json).files[0].id).toBe('one');
    expect(json.indexOf('"files"')).toBeLessThan(json.indexOf('"format"'));
  });
});

describe('parseArgs', () => {
  it('has safe defaults', () => {
    expect(parseArgs([])).toEqual({ update: false, matrix: false, only: [], reports: [] });
  });

  it('parses flags, repeated --only and paired reports', () => {
    const args = parseArgs(['--update', '--matrix', '--only', 'a,b', '--only', 'c', '--report', 'md', '--report-out', 'out.md', '--report', 'json', '--report', 'md']);
    expect(args.update).toBe(true);
    expect(args.matrix).toBe(true);
    expect(args.only).toEqual(['a', 'b', 'c']);
    expect(args.reports).toEqual([{ format: 'md', out: 'out.md' }, { format: 'json' }, { format: 'md' }]);
    expect(parseArgs(['--only']).only).toEqual([]);
  });

  it.each([
    [['--nope'], /unknown argument: --nope/],
    [['--report', 'html'], /expects "md" or "json", got html/],
    [['--report'], /expects "md" or "json", got undefined/],
    [['--report-out', 'x'], /must follow a --report/],
    [['--report', 'md', '--report-out', 'a', '--report-out', 'b'], /must follow a --report/],
    [['--report', 'md', '--report-out'], /needs a file path/],
  ])('rejects %j', (argv, message) => {
    expect(() => parseArgs(argv)).toThrow(message);
  });
});
