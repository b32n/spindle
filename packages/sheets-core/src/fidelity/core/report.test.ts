import {
  formatJson,
  formatMarkdown,
  mdEscape,
  plural,
  type FidelityReport,
  type ReportFile,
} from '../../../../../scripts/fidelity-core/report';
import type { DiffReport, SnapshotStatus } from '../../../../../scripts/fidelity-core/types';

const kinds = ['model', 'loss'];
const good = (): Record<string, SnapshotStatus> => ({ model: 'match', loss: 'match' });
const diff = (over: Partial<DiffReport> = {}): DiffReport => ({
  total: 3,
  added: 1,
  removed: 1,
  changed: 1,
  byField: { 'changed a[]': 2, 'added b': 1 },
  before: { sheets: 1, cells: 4 },
  after: { sheets: 1, cells: 5 },
  entries: [
    { path: 'a[0]', kind: 'changed', before: '1', after: '2' },
    { path: 'b', kind: 'added', after: '"x"' },
    { path: '', kind: 'removed', before: '"y"' },
  ],
  truncated: false,
  ...over,
});
const file = (id: string, extra: Partial<ReportFile> = {}): ReportFile => ({
  id,
  title: `Title of ${id}`,
  producer: 'prod|ucer',
  snapshots: good(),
  problems: [],
  importError: null,
  importExpectation: 'as-expected',
  warnings: [],
  loss: [],
  roundtrip: { status: 'ok', exportWarnings: [], reimportWarnings: [], diff: null },
  ...extra,
});
const report = (files: ReportFile[], extra: Partial<FidelityReport> = {}): FidelityReport => ({
  format: 'TEST',
  updateCommand: 'npm run test:update',
  ok: true,
  kinds,
  orphanGoldens: [],
  files,
  scoreboard: [],
  ...extra,
});

describe('helpers', () => {
  it('escapes markdown table text', () => {
    expect(mdEscape('a|b`c<d>\\e\nf\r\ng')).toBe("a\\|b'c&lt;d>\\\\e f g");
  });
  it('pluralises', () => {
    expect([plural(1, 'file'), plural(2, 'file'), plural(0, 'file')]).toEqual([
      '1 file',
      '2 files',
      '0 files',
    ]);
  });
});

describe('formatJson', () => {
  it('is stable JSON with sorted keys', () => {
    const text = formatJson(report([file('a')]));
    expect(text.endsWith('\n')).toBe(true);
    expect(JSON.parse(text).files[0].id).toBe('a');
    expect(text.indexOf('"files"')).toBeLessThan(text.indexOf('"format"'));
  });
});

describe('formatMarkdown', () => {
  it('summarises a clean run without optional columns or scoreboard', () => {
    const md = formatMarkdown(report([file('a'), file('b')]));
    expect(md).toContain('## TEST fidelity report');
    expect(md).toContain('**Result: all snapshots match their goldens.**');
    expect(md).toContain(
      '2 files · 4/4 snapshots match · 0 import errors · 0 files with round-trip differences · 0 loss occurrences in 0 loss types'
    );
    expect(md).not.toContain('Launch-bar scoreboard');
    expect(md).toContain('| File | Import | Warnings | Loss | Round trip | model | loss |');
    expect(md).toContain('| `a` | ok | 0 | 0 | clean | match | match |');
    expect(md).toContain('Loss report: nothing reported.');
    expect(md).toContain('Import: succeeded.');
    expect(md).not.toContain('Loss by type');
  });

  it('shows the scoreboard with the three statuses', () => {
    const md = formatMarkdown(
      report([file('a')], {
        scoreboard: [
          { id: 'L4', area: 'Sheets', criterion: 'c|1', measured: 'm', target: '0', status: 'met' },
          {
            id: 'L5',
            area: 'Sheets',
            criterion: 'c2',
            measured: 'm2',
            target: '98%',
            status: 'not-met',
          },
          {
            id: 'L6',
            area: 'Sheets',
            criterion: 'c3',
            measured: 'why',
            target: 't',
            status: 'not-measured',
          },
        ],
      })
    );
    expect(md).toContain('### Launch-bar scoreboard');
    expect(md).toContain('| L4 | Sheets | c\\|1 | m | 0 | met |');
    expect(md).toContain('| L5 | Sheets | c2 | m2 | 98% | not met |');
    expect(md).toContain('| L6 | Sheets | c3 | why | t | not measured here |');
  });

  it('explains a failed run, orphans, problems and unchecked files', () => {
    const md = formatMarkdown(
      report(
        [
          file('a', { snapshots: { model: 'mismatch', loss: 'missing' }, problems: ['bad thing'] }),
          file('b', { snapshots: { model: 'unchecked', loss: 'created' } }),
          file('c', { snapshots: { model: 'updated', loss: 'match' } }),
        ],
        { ok: false, orphanGoldens: ['old.model.json'] }
      )
    );
    expect(md).toContain(
      '**Result: FAILED.** 2 snapshots differ from the goldens, 1 problem, 1 orphaned golden.'
    );
    expect(md).toContain('run `npm run test:update`');
    expect(md).toContain('Orphaned goldens (no manifest entry): `old.model.json`');
    expect(md).toContain('| **mismatch** | **missing** |');
    expect(md).toContain('| not compared | match |');
    expect(md).toContain('- **Problem:** bad thing');
    expect(md.indexOf('<code>a</code>')).toBeLessThan(md.indexOf('<code>b</code>'));
    expect(md).toContain('<details open>');
    const single = formatMarkdown(
      report([file('a', { snapshots: { model: 'mismatch', loss: 'match' } })], { ok: false })
    );
    expect(single).toContain('1 snapshot differs');
  });

  it('describes import errors and expectations', () => {
    const md = formatMarkdown(
      report([
        file('thrown', { importError: 'kaput', importExpectation: 'unexpected-error' }),
        file('expected', { importError: 'meant to', importExpectation: 'as-expected' }),
        file('surprise', { importExpectation: 'unexpected-success' }),
      ])
    );
    expect(md).toContain('| `thrown` | **error** |');
    expect(md).toContain('| `expected` | error (expected) |');
    expect(md).toContain('| `surprise` | **ok (expected error)** |');
    expect(md).toContain('**Import error:** kaput');
    expect(md).toContain('**Import error (expected):** meant to');
    expect(md).toContain('The manifest says this file should fail to import, but it imported.');
    expect(md).toContain('2 import errors');
  });

  it('lists warnings with a cap, loss tables and loss totals', () => {
    const loss = [
      { id: 'x.a', label: 'A', count: 3, effect: 'lost' },
      { id: 'x.b', label: 'B', count: 1, effect: 'kept' },
    ];
    const md = formatMarkdown(
      report([
        file('a', { warnings: ['w1', 'w2', 'w3'], loss }),
        file('b', { loss: [loss[0], { id: 'x.c', label: 'C', count: 1, effect: 'kept' }] }),
      ]),
      { maxWarnings: 2 }
    );
    expect(md).toContain('Import warnings (3):');
    expect(md).toContain('- …and 1 more');
    expect(md).toContain('Loss report (2 item types, 4 occurrences):');
    expect(md).toContain('| `x.a` | 3 | A: lost |');
    expect(md).toContain('### Loss by type (all files)');
    expect(md).toContain('| `x.a` | 2 | 6 | A: lost |');
    expect(md.indexOf('| `x.b` | 1 | 1 |')).toBeLessThan(md.indexOf('| `x.c` | 1 | 1 |'));
    expect(md).toContain('| `a` | ok | 3 | 4 (2 types) |');
    expect(md).toContain('8 loss occurrences in 3 loss types');
  });

  it('prints round-trip results: statuses, warnings, clean and dirty diffs', () => {
    const rt = (
      status: 'ok' | 'skipped-import-error' | 'export-error' | 'reimport-error',
      extra = {}
    ) => ({ status, exportWarnings: [], reimportWarnings: [], diff: null, ...extra });
    const many = Object.fromEntries(Array.from({ length: 5 }, (_, i) => [`changed f${i}`, 1]));
    const md = formatMarkdown(
      report(
        [
          file('clean', { roundtrip: rt('ok', { diff: diff({ total: 0, entries: [] }) }) }),
          file('dirty', {
            roundtrip: rt('ok', {
              diff: diff(),
              exportWarnings: ['ew'],
              reimportWarnings: ['r1', 'r2'],
            }),
          }),
          file('many', { roundtrip: rt('ok', { diff: diff({ byField: many, entries: [] }) }) }),
          file('skip', { roundtrip: rt('skipped-import-error') }),
          file('exp', { roundtrip: rt('export-error', { error: 'no export' }) }),
          file('re', { roundtrip: rt('reimport-error', { error: 'no re-import' }) }),
        ],
        {}
      ),
      { maxFields: 2, maxDiffEntries: 2, maxWarnings: 1 }
    );
    expect(md).toContain('No differences (1 sheets, 4 cells).');
    expect(md).toContain('Differences: **3** (1 added, 1 removed, 1 changed).');
    expect(md).toContain('Before: 1 sheets, 4 cells.');
    expect(md).toContain('| `changed a[]` | 2 |');
    expect(md).toContain('…and 3 more fields.');
    expect(md).toContain('First 2 differences:');
    expect(md).toContain('- changed `a[0]`: 1 → 2');
    expect(md).toContain('- added `b`: "x"');
    expect(md).not.toContain('(root)');
    expect(md).toContain('- export warning: ew');
    expect(md).toContain('- re-import warning: r1');
    expect(md).not.toContain('r2');
    expect(md).toContain('skipped (import error)');
    expect(md).toContain('**export-error**: no export');
    expect(md).toContain('**reimport-error**: no re-import');
    expect(md).toContain('2 round-trip failures');
    expect(md).toContain('2 files with round-trip differences');
    const removedRoot = formatMarkdown(
      report([
        file('r', {
          roundtrip: rt('ok', {
            diff: diff({
              entries: [
                { path: '', kind: 'removed', before: '"y"' },
                { path: '', kind: 'added', after: '1' },
              ],
            }),
          }),
        }),
      ])
    );
    expect(removedRoot).toContain('- removed `(root)`: "y"');
    expect(removedRoot).toContain('- added `(root)`: 1');
    const noCounts = formatMarkdown(
      report([
        file('z', {
          roundtrip: rt('ok', { diff: diff({ total: 0, before: {}, byField: {}, entries: [] }) }),
        }),
      ])
    );
    expect(noCounts).toContain('No differences (nothing counted).');
    const grouped = formatMarkdown(
      report([
        file('g', {
          roundtrip: rt('ok', {
            diff: diff({
              total: 0,
              before: { sections: 1, blocks: { heading: 1, paragraph: 10 }, textChars: 589 },
              byField: {},
              entries: [],
            }),
          }),
        }),
      ])
    );
    expect(grouped).toContain(
      'No differences (1 sections, 1 heading, 10 paragraph, 589 textChars).'
    );
    const noFields = formatMarkdown(
      report([file('y', { roundtrip: rt('ok', { diff: diff({ byField: {}, entries: [] }) }) })])
    );
    expect(noFields).not.toContain('| Field |');
  });

  it('adds ledger and values columns and detail sections when files carry them', () => {
    const md = formatMarkdown(
      report([
        file('with', {
          ledger: {
            mapped: 5,
            opaque: 0,
            dropped: 2,
            unaccounted: 3,
            total: 10,
            unaccountedParts: ['docProps/app.xml'],
          },
          values: {
            status: 'ok',
            compared: 4,
            matched: 3,
            mismatched: 1,
            firstMismatches: [
              { sheet: 'S|1', cell: 'A1', expected: '1', actual: '2' },
              { sheet: 'S', cell: 'B2', expected: '3', actual: '4' },
            ],
          },
        }),
        file('empty', {
          ledger: {
            mapped: 0,
            opaque: 0,
            dropped: 0,
            unaccounted: 0,
            total: 0,
            unaccountedParts: [],
          },
          values: {
            status: 'skipped-import-error',
            compared: 0,
            matched: 0,
            mismatched: 0,
            firstMismatches: [],
          },
        }),
        file('plain'),
      ]),
      { maxMismatches: 1 }
    );
    expect(md).toContain('| Ledger (unaccounted/parts) | Values (matched/compared) |');
    expect(md).toContain('| 3/10 | 3/4 |');
    expect(md).toContain('| 0/0 | skipped-import-error |');
    expect(md).toContain('| `plain` | ok | 0 | 0 | clean |  |  | match | match |');
    expect(md).toContain(
      'Package ledger: 5 mapped, 0 opaque, 2 dropped, **3 unaccounted** of 10 parts.'
    );
    expect(md).toContain('Unaccounted: `docProps/app.xml`');
    expect(md).toContain(
      'Values (engine recalculation vs cached values): ok; 4 compared, 3 matched, **1 mismatched**.'
    );
    expect(md).toContain('| S\\|1 | `A1` | 1 | 2 |');
    expect(md).not.toContain('`B2`');
    const onlyValues = formatMarkdown(
      report([
        file('v', {
          values: { status: 'ok', compared: 0, matched: 0, mismatched: 0, firstMismatches: [] },
        }),
      ])
    );
    expect(onlyValues).not.toContain('Ledger (unaccounted');
    expect(onlyValues).toContain('| ok |');
    const onlyLedger = formatMarkdown(
      report([
        file('l', {
          ledger: {
            mapped: 1,
            opaque: 0,
            dropped: 0,
            unaccounted: 0,
            total: 1,
            unaccountedParts: [],
          },
        }),
      ])
    );
    expect(onlyLedger).not.toContain('Values (matched');
  });

  it('drops per-file details to stay under the size limit, failures first', () => {
    const files = [
      file('aaa'),
      file('bbb', { snapshots: { model: 'mismatch', loss: 'match' } }),
      file('ccc'),
    ];
    const md = formatMarkdown(report(files, { ok: false }), { maxChars: 1100 });
    expect(md).toContain('Details omitted to stay under the summary size limit');
    expect(md).toContain('<code>bbb</code>');
    const omitted = /omitted[^:]*: (.*)\. The JSON/.exec(md)?.[1] ?? '';
    expect(omitted).not.toContain('bbb');
    expect(omitted).toContain('`ccc`');
  });
});
