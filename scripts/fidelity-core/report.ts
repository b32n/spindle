/**
 * Human-readable summary of a fidelity run, written for the CI job summary (`$GITHUB_STEP_SUMMARY`)
 * and as a JSON artifact. The report shape is format-neutral: each format adapter maps its harness
 * results onto `FidelityReport` and reuses `formatMarkdown` / `formatJson` unchanged.
 */
import { stableStringify } from './snapshot';
import type { ScoreRow } from './scoreboard';
import {
  GOOD_STATUSES,
  type DiffReport,
  type ImportExpectation,
  type LedgerTotals,
  type LossItem,
  type RoundtripSnapshot,
  type SnapshotStatus,
  type ValuesTotals,
} from './types';

export interface LedgerSummary extends LedgerTotals {
  /** Names of the parts nothing accounts for (capped by the adapter). */
  unaccountedParts: string[];
}

export interface ValueMismatch {
  sheet: string;
  cell: string;
  expected: string;
  actual: string;
}

export interface ValuesSummary extends ValuesTotals {
  status: string;
  /** Formulas the engine could not parse, which keep the file's value: not part of the totals above. */
  passthrough?: number;
  firstMismatches: ValueMismatch[];
}

export interface ReportFile {
  id: string;
  title: string;
  producer: string;
  snapshots: Record<string, SnapshotStatus>;
  /** Problems other than snapshot drift (for example an unexpected import error under `strictImport`). */
  problems: string[];
  importError: string | null;
  importExpectation: ImportExpectation;
  warnings: string[];
  loss: LossItem[];
  roundtrip: RoundtripSnapshot;
  ledger?: LedgerSummary;
  values?: ValuesSummary;
}

export interface FidelityReport {
  /** Display name of the format, e.g. "XLSX". */
  format: string;
  /** The command that rewrites the goldens, shown when a run fails. */
  updateCommand: string;
  ok: boolean;
  /** Snapshot kinds, in column order. */
  kinds: string[];
  orphanGoldens: string[];
  files: ReportFile[];
  /** Launch-bar rows this run can measure (plan 7.3); empty for a format that has none yet. */
  scoreboard: ScoreRow[];
}

export interface FormatOptions {
  /** Warnings listed per file. */
  maxWarnings?: number;
  /** Round-trip differences listed per file. */
  maxDiffEntries?: number;
  /** Per-field difference counts listed per file. */
  maxFields?: number;
  /** Value mismatches listed per file. */
  maxMismatches?: number;
  /** GitHub caps a step summary at 1 MiB; per-file details are dropped to stay below this many characters. */
  maxChars?: number;
}

const DEFAULTS: Required<FormatOptions> = {
  maxWarnings: 10,
  maxDiffEntries: 10,
  maxFields: 12,
  maxMismatches: 10,
  maxChars: 900_000,
};

export function formatJson(report: FidelityReport): string {
  return stableStringify(report);
}

const isGood = (s: SnapshotStatus) => GOOD_STATUSES.includes(s);

/** Escape text for a markdown table cell or list item. */
export function mdEscape(text: string): string {
  return text
    .replace(/\\/g, '\\\\')
    .replace(/\|/g, '\\|')
    .replace(/`/g, "'")
    .replace(/[\r\n]+/g, ' ')
    .replace(/</g, '&lt;');
}

const code = (s: string) => `\`${s.replace(/`/g, "'")}\``;
export const plural = (n: number, word: string) => `${n} ${word}${n === 1 ? '' : 's'}`;
const sum = (items: LossItem[]) => items.reduce((n, i) => n + i.count, 0);

function roundtripCell(rt: RoundtripSnapshot): string {
  if (rt.status === 'skipped-import-error') return 'skipped (import error)';
  if (rt.status !== 'ok') return `**${rt.status}**`;
  return rt.diff && rt.diff.total > 0 ? plural(rt.diff.total, 'difference') : 'clean';
}

const snapshotCell = (s: SnapshotStatus) =>
  isGood(s) ? (s === 'unchecked' ? 'not compared' : 'match') : `**${s}**`;

function lossTable(items: LossItem[]): string[] {
  return [
    '| Loss id | Count | Effect |',
    '|---|--:|---|',
    ...items.map((i) => `| ${code(i.id)} | ${i.count} | ${mdEscape(`${i.label}: ${i.effect}`)} |`),
  ];
}

const countsText = (c: DiffReport['before']) =>
  Object.entries(c)
    .map(([k, n]) => `${n} ${k}`)
    .join(', ') || 'nothing counted';

function diffLines(diff: DiffReport, o: Required<FormatOptions>): string[] {
  const out: string[] = [];
  out.push(
    `Differences: **${diff.total}** (${diff.added} added, ${diff.removed} removed, ${diff.changed} changed).`
  );
  out.push(`Before: ${countsText(diff.before)}.`, `After: ${countsText(diff.after)}.`);
  const fields = Object.entries(diff.byField).sort(
    (a, b) => b[1] - a[1] || a[0].localeCompare(b[0])
  );
  if (fields.length) {
    out.push(
      '',
      '| Field | Differences |',
      '|---|--:|',
      ...fields.slice(0, o.maxFields).map(([f, n]) => `| ${code(f)} | ${n} |`)
    );
    if (fields.length > o.maxFields)
      out.push('', `…and ${fields.length - o.maxFields} more fields.`);
  }
  if (diff.entries.length) {
    out.push('', `First ${Math.min(diff.entries.length, o.maxDiffEntries)} differences:`, '');
    for (const e of diff.entries.slice(0, o.maxDiffEntries)) {
      const [before, after] = [mdEscape(e.before ?? ''), mdEscape(e.after ?? '')];
      const change =
        e.kind === 'changed' ? `${before} → ${after}` : e.kind === 'added' ? after : before;
      out.push(`- ${e.kind} ${code(e.path || '(root)')}: ${change}`);
    }
  }
  return out;
}

const valuesCell = (v: ValuesSummary | undefined) =>
  v && v.compared > 0 ? `${v.matched}/${v.compared}` : v ? v.status : '';
const ledgerCell = (l: LedgerSummary | undefined) => (l ? `${l.unaccounted}/${l.total}` : '');

function isClean(f: ReportFile, kinds: string[]): boolean {
  return f.problems.length === 0 && kinds.every((k) => isGood(f.snapshots[k]));
}

function fileDetails(f: ReportFile, kinds: string[], o: Required<FormatOptions>): string {
  const bad = !isClean(f, kinds);
  const lines: string[] = [
    `<details${bad ? ' open' : ''}>`,
    `<summary><code>${mdEscape(f.id)}</code>: ${mdEscape(f.title)}</summary>`,
    '',
  ];
  lines.push(`Produced by: ${mdEscape(f.producer)}`, '');
  lines.push(
    `Snapshots: ${kinds.map((k) => `${k} ${snapshotCell(f.snapshots[k])}`).join(', ')}`,
    ''
  );
  for (const p of f.problems) lines.push(`- **Problem:** ${mdEscape(p)}`);
  if (f.problems.length) lines.push('');
  lines.push(
    f.importError
      ? `**Import error${f.importExpectation === 'as-expected' ? ' (expected)' : ''}:** ${mdEscape(f.importError)}`
      : 'Import: succeeded.',
    ''
  );
  if (f.importExpectation === 'unexpected-success')
    lines.push('The manifest says this file should fail to import, but it imported.', '');
  if (f.warnings.length) {
    lines.push(`Import warnings (${f.warnings.length}):`, '');
    for (const w of f.warnings.slice(0, o.maxWarnings)) lines.push(`- ${mdEscape(w)}`);
    if (f.warnings.length > o.maxWarnings)
      lines.push(`- …and ${f.warnings.length - o.maxWarnings} more`);
    lines.push('');
  }
  if (f.loss.length)
    lines.push(
      `Loss report (${plural(f.loss.length, 'item type')}, ${sum(f.loss)} occurrences):`,
      '',
      ...lossTable(f.loss),
      ''
    );
  else lines.push('Loss report: nothing reported.', '');
  const rt = f.roundtrip;
  lines.push(
    `Round trip (import → export → import): **${rt.status}**${rt.error ? `: ${mdEscape(rt.error)}` : ''}`,
    ''
  );
  for (const w of rt.exportWarnings) lines.push(`- export warning: ${mdEscape(w)}`);
  for (const w of rt.reimportWarnings.slice(0, o.maxWarnings))
    lines.push(`- re-import warning: ${mdEscape(w)}`);
  if (rt.exportWarnings.length || rt.reimportWarnings.length) lines.push('');
  if (rt.diff)
    lines.push(
      ...(rt.diff.total === 0
        ? [`No differences (${countsText(rt.diff.before)}).`]
        : diffLines(rt.diff, o)),
      ''
    );
  if (f.ledger) {
    const l = f.ledger;
    lines.push(
      `Package ledger: ${l.mapped} mapped, ${l.opaque} opaque, ${l.dropped} dropped, **${l.unaccounted} unaccounted** of ${l.total} parts.`,
      ''
    );
    if (l.unaccountedParts.length)
      lines.push(`Unaccounted: ${l.unaccountedParts.map(code).join(', ')}`, '');
  }
  if (f.values) {
    const v = f.values;
    lines.push(
      `Values (engine recalculation vs cached values): ${v.status}; ${v.compared} compared, ${v.matched} matched, **${v.mismatched} mismatched**.`,
      ''
    );
    if (v.firstMismatches.length) {
      lines.push('| Sheet | Cell | Expected | Actual |', '|---|---|---|---|');
      for (const m of v.firstMismatches.slice(0, o.maxMismatches))
        lines.push(
          `| ${mdEscape(m.sheet)} | ${code(m.cell)} | ${mdEscape(m.expected)} | ${mdEscape(m.actual)} |`
        );
      lines.push('');
    }
  }
  lines.push('</details>', '');
  return lines.join('\n');
}

const STATUS_TEXT: Record<ScoreRow['status'], string> = {
  met: 'met',
  'not-met': 'not met',
  'not-measured': 'not measured here',
};

function scoreboardLines(rows: ScoreRow[]): string[] {
  if (rows.length === 0) return [];
  return [
    '### Launch-bar scoreboard',
    '',
    'Rows of the launch-ready bar (plan section 7.3) that this run can measure. A row that is not met is the recorded baseline, not a failure.',
    '',
    '| # | Area | Criterion | Measured | Target | Status |',
    '|---|---|---|---|---|---|',
    ...rows.map(
      (r) =>
        `| ${r.id} | ${mdEscape(r.area)} | ${mdEscape(r.criterion)} | ${mdEscape(r.measured)} | ${mdEscape(r.target)} | ${STATUS_TEXT[r.status]} |`
    ),
    '',
  ];
}

/** GitHub-flavoured markdown report, in the order a reviewer needs it: verdict, scoreboard, overview, loss totals, per-file details. */
export function formatMarkdown(report: FidelityReport, options: FormatOptions = {}): string {
  const o = { ...DEFAULTS, ...options };
  const { files, kinds } = report;
  const snapshotTotal = files.length * kinds.length;
  const drifted = files.reduce(
    (n, f) => n + kinds.filter((k) => !isGood(f.snapshots[k])).length,
    0
  );
  const importErrors = files.filter((f) => f.importError).length;
  const rtDirty = files.filter(
    (f) => f.roundtrip.status === 'ok' && f.roundtrip.diff && f.roundtrip.diff.total > 0
  ).length;
  const rtBroken = files.filter(
    (f) => f.roundtrip.status === 'export-error' || f.roundtrip.status === 'reimport-error'
  ).length;
  const lossAll = files.flatMap((f) => f.loss);
  const withLedger = files.some((f) => f.ledger);
  const withValues = files.some((f) => f.values);

  const out: string[] = [`## ${report.format} fidelity report`, ''];
  const problems = files.reduce((n, f) => n + f.problems.length, 0);
  if (report.ok)
    out.push(
      '**Result: all snapshots match their goldens.** Losses, import errors and round-trip differences below are the current, recorded behaviour, not failures.',
      ''
    );
  else {
    out.push(
      `**Result: FAILED.** ${plural(drifted, 'snapshot')} ${drifted === 1 ? 'differs' : 'differ'} from the goldens, ${plural(problems, 'problem')}, ${plural(report.orphanGoldens.length, 'orphaned golden')}.`,
      ''
    );
    out.push(
      `If the change is intended, run \`${report.updateCommand}\`, review the diff and commit the new goldens.`,
      ''
    );
  }
  out.push(...scoreboardLines(report.scoreboard));
  out.push(
    `${plural(files.length, 'file')} · ${snapshotTotal - drifted}/${snapshotTotal} snapshots match · ${plural(importErrors, 'import error')} · ` +
      `${plural(rtDirty, 'file')} with round-trip differences${rtBroken ? ` · ${plural(rtBroken, 'round-trip failure')}` : ''} · ${sum(lossAll)} loss occurrences in ${new Set(lossAll.map((i) => i.id)).size} loss types`,
    ''
  );
  if (report.orphanGoldens.length)
    out.push(
      'Orphaned goldens (no manifest entry): ' + report.orphanGoldens.map(code).join(', '),
      ''
    );

  const extra = [
    ...(withLedger ? ['Ledger (unaccounted/parts)'] : []),
    ...(withValues ? ['Values (matched/compared)'] : []),
  ];
  out.push(
    '### Files',
    '',
    `| File | Import | Warnings | Loss | Round trip |${extra.map((e) => ` ${e} |`).join('')} ${kinds.join(' | ')} |`,
    `|---|---|--:|--:|---|${extra.map(() => '--:|').join('')}${kinds.map(() => '---|').join('')}`
  );
  for (const f of files) {
    const lossCell = f.loss.length ? `${sum(f.loss)} (${plural(f.loss.length, 'type')})` : '0';
    const importCell = f.importError
      ? f.importExpectation === 'as-expected'
        ? 'error (expected)'
        : '**error**'
      : f.importExpectation === 'unexpected-success'
        ? '**ok (expected error)**'
        : 'ok';
    const extraCells = [
      ...(withLedger ? [ledgerCell(f.ledger)] : []),
      ...(withValues ? [valuesCell(f.values)] : []),
    ];
    out.push(
      `| ${code(f.id)} | ${importCell} | ${f.warnings.length} | ${lossCell} | ${roundtripCell(f.roundtrip)} |${extraCells.map((c) => ` ${c} |`).join('')} ${kinds.map((k) => snapshotCell(f.snapshots[k])).join(' | ')} |`
    );
  }
  out.push('');

  const byLoss = new Map<string, { label: string; effect: string; files: number; count: number }>();
  for (const f of files) {
    for (const i of f.loss) {
      const cur = byLoss.get(i.id) ?? { label: i.label, effect: i.effect, files: 0, count: 0 };
      cur.files++;
      cur.count += i.count;
      byLoss.set(i.id, cur);
    }
  }
  if (byLoss.size) {
    out.push(
      '### Loss by type (all files)',
      '',
      '| Loss id | Files | Occurrences | Effect |',
      '|---|--:|--:|---|'
    );
    for (const [id, v] of [...byLoss].sort(
      (a, b) => b[1].count - a[1].count || a[0].localeCompare(b[0])
    )) {
      out.push(
        `| ${code(id)} | ${v.files} | ${v.count} | ${mdEscape(`${v.label}: ${v.effect}`)} |`
      );
    }
    out.push('');
  }

  const head = out.join('\n');
  // Failures first, so they are visible even when the details are cut for size.
  const ordered = [...files].sort(
    (a, b) => Number(isClean(a, kinds)) - Number(isClean(b, kinds)) || a.id.localeCompare(b.id)
  );
  let details = '### Details\n\n';
  const omitted: string[] = [];
  for (const f of ordered) {
    const block = fileDetails(f, kinds, o);
    if (head.length + details.length + block.length > o.maxChars) omitted.push(f.id);
    else details += block + '\n';
  }
  if (omitted.length)
    details += `_Details omitted to stay under the summary size limit: ${omitted.map(code).join(', ')}. The JSON artifact has everything._\n`;
  return `${head}\n${details}`;
}
