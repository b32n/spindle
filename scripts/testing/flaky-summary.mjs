// Summarise flaky / failing tests from a Playwright JSON report (as produced by
// `playwright merge-reports --reporter=json`). Used by the E2E and nightly workflows.
//   node scripts/testing/flaky-summary.mjs report.json
// Writes a markdown table to $GITHUB_STEP_SUMMARY (when set), emits one ::error
// annotation per flaky/failing test, and exits 1 if there is any.
//   "flaky"      = failed then passed on a retry (PR runs: retries=2)
//   "failing"    = failed on every attempt
// Nightly runs use --retries=0 --repeat-each=N, so each repeat is its own test
// entry; entries are grouped by file/project/title and the failure rate is shown.
import { readFileSync, appendFileSync } from 'node:fs';

const file = process.argv[2];
if (!file) {
  console.error('usage: flaky-summary.mjs <report.json>');
  process.exit(2);
}
const report = JSON.parse(readFileSync(file, 'utf8'));

/** key -> { runs, flaky, failed } */
const groups = new Map();
function walk(suite) {
  for (const spec of suite.specs ?? []) {
    for (const t of spec.tests ?? []) {
      const key = `${spec.file} › ${t.projectName} › ${spec.title}`;
      const g = groups.get(key) ?? { runs: 0, flaky: 0, failed: 0 };
      if (t.status === 'skipped') continue;
      g.runs += 1;
      if (t.status === 'flaky') g.flaky += 1;
      if (t.status === 'unexpected') g.failed += 1;
      groups.set(key, g);
    }
  }
  for (const s of suite.suites ?? []) walk(s);
}
for (const s of report.suites ?? []) walk(s);

if (groups.size === 0) {
  console.log('::error title=E2E report::no tests found in the JSON report');
  process.exit(1);
}

const bad = [...groups].filter(([, g]) => g.flaky || g.failed);
const lines = [];
if (bad.length === 0) {
  lines.push('### E2E: no flaky or failing tests', '');
} else {
  lines.push(`### E2E: ${bad.length} flaky/failing test(s)`, '', '| Test | Runs | Flaky (passed on retry) | Failed |', '|---|---|---|---|');
  for (const [key, g] of bad) {
    lines.push(`| ${key.replaceAll('|', '\\|')} | ${g.runs} | ${g.flaky} | ${g.failed} |`);
    const kind = g.failed ? 'failing' : 'flaky';
    console.log(`::error title=E2E ${kind} test::${key} (${g.failed} failed, ${g.flaky} flaky of ${g.runs} runs)`);
  }
  lines.push('');
}
const out = lines.join('\n');
console.log(out);
if (process.env.GITHUB_STEP_SUMMARY) appendFileSync(process.env.GITHUB_STEP_SUMMARY, out + '\n');
process.exit(bad.length ? 1 : 0);
