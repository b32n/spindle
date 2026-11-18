/**
 * Human-readable summary of a fidelity run, written for the CI job summary (`$GITHUB_STEP_SUMMARY`)
 * and as a JSON artifact. The report shape is format-neutral on purpose: the XLSX and PPTX corpora
 * can map their own harness results onto `FidelityReport` and reuse `formatMarkdown` unchanged.
 */
import { stableStringify } from './snapshot';
import type { HarnessReport } from './harness';
import { SNAPSHOT_KINDS, type DiffReport, type LossItem, type RoundtripSnapshot, type SnapshotKind, type SnapshotStatus } from './types';

export interface ReportFile {
  id: string;
  title: string;
  producer: string;
  snapshots: Record<SnapshotKind, SnapshotStatus>;
  /** Problems other than snapshot drift (for example an unexpected import error). */
  problems: string[];
  importError: string | null;
  warnings: string[];
  loss: LossItem[];
  roundtrip: RoundtripSnapshot;
}

export interface FidelityReport {
  /** Display name of the format, e.g. "DOCX". */
  format: string;
  ok: boolean;
  orphanGoldens: string[];
  files: ReportFile[];
}

export interface FormatOptions {
  /** Warnings listed per file. */
  maxWarnings?: number;
  /** Round-trip differences listed per file. */
  maxDiffEntries?: number;
  /** Per-field difference counts listed per file. */
  maxFields?: number;
  /** GitHub caps a step summary at 1 MiB; per-file details are dropped to stay below this many characters. */
  maxChars?: number;
}

const DEFAULTS: Required<FormatOptions> = { maxWarnings: 10, maxDiffEntries: 10, maxFields: 12, maxChars: 900_000 };

/** Map a docx harness run onto the format-neutral report. */
export function toFidelityReport(format: string, run: HarnessReport): FidelityReport {
  return {
    format,
    ok: run.ok,
    orphanGoldens: run.orphanGoldens,
    files: run.fixtures.map((o) => ({
      id: o.entry.id,
      title: o.entry.title,
      producer: o.entry.source.tool,
      snapshots: o.status,
      problems: o.problems,
      importError: o.result.model.importError,
      warnings: o.result.model.warnings,
      loss: o.result.loss.items,
      roundtrip: o.result.roundtrip,
    })),
  };
}

export function formatJson(report: FidelityReport): string {
  return stableStringify(report);
}

const GOOD: SnapshotStatus[] = ['match', 'created', 'updated'];
const isGood = (s: SnapshotStatus) => GOOD.includes(s);

/** Escape text for a markdown table cell or list item. */
export function mdEscape(text: string): string {
  return text.replace(/\\/g, '\\\\').replace(/\|/g, '\\|').replace(/`/g, "'").replace(/[\r\n]+/g, ' ').replace(/</g, '&lt;');
}

const code = (s: string) => `\`${s.replace(/`/g, "'")}\``;
const plural = (n: number, word: string) => `${n} ${word}${n === 1 ? '' : 's'}`;
const sum = (items: LossItem[]) => items.reduce((n, i) => n + i.count, 0);

function roundtripCell(rt: RoundtripSnapshot): string {
  if (rt.status === 'skipped-import-error') return 'skipped (import error)';
  if (rt.status !== 'ok') return `**${rt.status}**`;
  return rt.diff && rt.diff.total > 0 ? plural(rt.diff.total, 'difference') : 'clean';
}

function snapshotCell(s: SnapshotStatus): string {
  return isGood(s) ? 'match' : `**${s}**`;
}

function lossTable(items: LossItem[]): string[] {
  return ['| Loss id | Count | Effect |', '|---|--:|---|', ...items.map((i) => `| ${code(i.id)} | ${i.count} | ${mdEscape(`${i.label}: ${i.effect}`)} |`)];
}

function diffLines(diff: DiffReport, o: Required<FormatOptions>): string[] {
  const out: string[] = [];
  out.push(`Differences: **${diff.total}** (${diff.added} added, ${diff.removed} removed, ${diff.changed} changed).`);
  const counts = (c: DiffReport['before']) => `${plural(c.sections, 'section')}, ${Object.entries(c.blocks).map(([t, n]) => `${n} ${t}`).join(', ') || 'no blocks'}, ${c.textChars} characters`;
  out.push(`Before: ${counts(diff.before)}.`, `After: ${counts(diff.after)}.`);
  const fields = Object.entries(diff.byField).sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]));
  if (fields.length) {
    out.push('', '| Field | Differences |', '|---|--:|', ...fields.slice(0, o.maxFields).map(([f, n]) => `| ${code(f)} | ${n} |`));
    if (fields.length > o.maxFields) out.push('', `…and ${fields.length - o.maxFields} more fields.`);
  }
  if (diff.entries.length) {
    out.push('', `First ${Math.min(diff.entries.length, o.maxDiffEntries)} differences:`, '');
    for (const e of diff.entries.slice(0, o.maxDiffEntries)) {
      const change = e.kind === 'changed' ? `${mdEscape(e.before ?? '')} → ${mdEscape(e.after ?? '')}` : e.kind === 'added' ? mdEscape(e.after ?? '') : mdEscape(e.before ?? '');
      out.push(`- ${e.kind} ${code(e.path || '(root)')}: ${change}`);
    }
  }
  return out;
}

function fileDetails(f: ReportFile, o: Required<FormatOptions>): string {
  const bad = SNAPSHOT_KINDS.some((k) => !isGood(f.snapshots[k])) || f.problems.length > 0;
  const lines: string[] = [`<details${bad ? ' open' : ''}>`, `<summary><code>${mdEscape(f.id)}</code>: ${mdEscape(f.title)}</summary>`, ''];
  lines.push(`Produced by: ${mdEscape(f.producer)}`, '');
  lines.push(`Snapshots: ${SNAPSHOT_KINDS.map((k) => `${k} ${snapshotCell(f.snapshots[k])}`).join(', ')}`, '');
  for (const p of f.problems) lines.push(`- **Problem:** ${mdEscape(p)}`);
  if (f.problems.length) lines.push('');
  lines.push(f.importError ? `**Import error:** ${mdEscape(f.importError)}` : 'Import: succeeded.', '');
  if (f.warnings.length) {
    lines.push(`Import warnings (${f.warnings.length}):`, '');
    for (const w of f.warnings.slice(0, o.maxWarnings)) lines.push(`- ${mdEscape(w)}`);
    if (f.warnings.length > o.maxWarnings) lines.push(`- …and ${f.warnings.length - o.maxWarnings} more`);
    lines.push('');
  }
  if (f.loss.length) lines.push(`Loss report (${plural(f.loss.length, 'item type')}, ${sum(f.loss)} occurrences):`, '', ...lossTable(f.loss), '');
  else lines.push('Loss report: nothing reported.', '');
  const rt = f.roundtrip;
  lines.push(`Round trip (import → export → import): **${rt.status}**${rt.error ? `: ${mdEscape(rt.error)}` : ''}`, '');
  for (const w of rt.exportWarnings) lines.push(`- export warning: ${mdEscape(w)}`);
  for (const w of rt.reimportWarnings.slice(0, o.maxWarnings)) lines.push(`- re-import warning: ${mdEscape(w)}`);
  if (rt.exportWarnings.length || rt.reimportWarnings.length) lines.push('');
  if (rt.diff) lines.push(...(rt.diff.total === 0 ? [`No differences (${plural(rt.diff.before.sections, 'section')}, ${Object.values(rt.diff.before.blocks).reduce((a, b) => a + b, 0)} blocks).`] : diffLines(rt.diff, o)), '');
  lines.push('</details>', '');
  return lines.join('\n');
}

/** GitHub-flavoured markdown report, in the order a reviewer needs it: verdict, overview, loss totals, per-file details. */
export function formatMarkdown(report: FidelityReport, options: FormatOptions = {}): string {
  const o = { ...DEFAULTS, ...options };
  const files = report.files;
  const snapshotTotal = files.length * SNAPSHOT_KINDS.length;
  const drifted = files.reduce((n, f) => n + SNAPSHOT_KINDS.filter((k) => !isGood(f.snapshots[k])).length, 0);
  const importErrors = files.filter((f) => f.importError).length;
  const rtDirty = files.filter((f) => f.roundtrip.status === 'ok' && f.roundtrip.diff && f.roundtrip.diff.total > 0).length;
  const rtBroken = files.filter((f) => f.roundtrip.status === 'export-error' || f.roundtrip.status === 'reimport-error').length;
  const lossAll = files.flatMap((f) => f.loss);

  const out: string[] = [`## ${report.format} fidelity report`, ''];
  const problems = files.reduce((n, f) => n + f.problems.length, 0);
  if (report.ok) out.push('**Result: all snapshots match their goldens.** Losses and round-trip differences below are the current, recorded behaviour, not failures.', '');
  else {
    out.push(`**Result: FAILED.** ${plural(drifted, 'snapshot')} ${drifted === 1 ? 'differs' : 'differ'} from the goldens, ${plural(problems, 'problem')}, ${plural(report.orphanGoldens.length, 'orphaned golden')}.`, '');
    out.push('If the change is intended, run `npm run docs:docx-fidelity:update`, review the diff and commit the new goldens.', '');
  }
  out.push(
    `${plural(files.length, 'file')} · ${snapshotTotal - drifted}/${snapshotTotal} snapshots match · ${plural(importErrors, 'import error')} · ` +
      `${plural(rtDirty, 'file')} with round-trip differences${rtBroken ? ` · ${plural(rtBroken, 'round-trip failure')}` : ''} · ${sum(lossAll)} loss occurrences in ${new Set(lossAll.map((i) => i.id)).size} loss types`,
    ''
  );
  if (report.orphanGoldens.length) out.push('Orphaned goldens (no manifest entry): ' + report.orphanGoldens.map(code).join(', '), '');

  out.push('### Files', '', '| File | Import | Warnings | Loss | Round trip | model | loss | roundtrip |', '|---|---|--:|--:|---|---|---|---|');
  for (const f of files) {
    const lossCell = f.loss.length ? `${sum(f.loss)} (${plural(f.loss.length, 'type')})` : '0';
    out.push(`| ${code(f.id)} | ${f.importError ? '**error**' : 'ok'} | ${f.warnings.length} | ${lossCell} | ${roundtripCell(f.roundtrip)} | ${SNAPSHOT_KINDS.map((k) => snapshotCell(f.snapshots[k])).join(' | ')} |`);
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
    out.push('### Loss by type (all files)', '', '| Loss id | Files | Occurrences | Effect |', '|---|--:|--:|---|');
    for (const [id, v] of [...byLoss].sort((a, b) => b[1].count - a[1].count || a[0].localeCompare(b[0]))) {
      out.push(`| ${code(id)} | ${v.files} | ${v.count} | ${mdEscape(`${v.label}: ${v.effect}`)} |`);
    }
    out.push('');
  }

  const head = out.join('\n');
  // Failures first, so they are visible even when the details are cut for size.
  const ordered = [...files].sort((a, b) => Number(isClean(a)) - Number(isClean(b)) || a.id.localeCompare(b.id));
  let details = '### Details\n\n';
  const omitted: string[] = [];
  for (const f of ordered) {
    const block = fileDetails(f, o);
    if (head.length + details.length + block.length > o.maxChars) omitted.push(f.id);
    else details += block + '\n';
  }
  if (omitted.length) details += `_Details omitted to stay under the summary size limit: ${omitted.map(code).join(', ')}. The JSON artifact has everything._\n`;
  return `${head}\n${details}`;
}

function isClean(f: ReportFile): boolean {
  return f.problems.length === 0 && SNAPSHOT_KINDS.every((k) => isGood(f.snapshots[k]));
}
