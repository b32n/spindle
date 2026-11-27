// Local second opinion on Office files (plan A4): opens each file in headless LibreOffice, converts it to PDF and reports the exit
// status and the page count, with an optional PNG per page (`pdftoppm`). NOT part of CI (the Playwright image has no LibreOffice):
// an agent or a developer runs it in a checkout that has `soffice`.
//
//   node scripts/fidelity-core/lo-check.mjs [--png] [--out <dir>] file.pptx ...
//   node scripts/fidelity-core/lo-check.mjs --export-pptx [--png] [--out <dir>] file.pptx ...
//
// `--export-pptx` first reads each file with the built `@b32nio/spindle-io/pptx` (run `npm run build`), writes it again with
// `exportPptx` into `<out>/exported/` and checks that file; it also prints the writer's own "not exported yet" notes.
// Exit status 1 when any conversion fails or gives no pages.
import { spawnSync } from 'node:child_process';
import { mkdirSync, mkdtempSync, readFileSync, writeFileSync } from 'node:fs';
import os from 'node:os';
import path from 'node:path';

const args = process.argv.slice(2);
const flag = (name) => args.includes(name) && args.splice(args.indexOf(name), 1).length > 0;
const option = (name) => {
  const i = args.indexOf(name);
  return i === -1 ? undefined : args.splice(i, 2)[1];
};
const png = flag('--png');
const exportFirst = flag('--export-pptx');
const out = path.resolve(option('--out') ?? mkdtempSync(path.join(os.tmpdir(), 'lo-check-')));
const files = args;
if (files.length === 0) {
  console.error('usage: lo-check.mjs [--png] [--export-pptx] [--out dir] file ...');
  process.exit(2);
}
mkdirSync(out, { recursive: true });

const run = (cmd, argv) => spawnSync(cmd, argv, { encoding: 'utf8', timeout: 120_000 });
if (run('soffice', ['--version']).status !== 0) {
  console.error('soffice is not available: nothing was checked');
  process.exit(2);
}

/** The pages of a PDF: `pdfinfo` when present, else the `/Type /Page` objects. */
function pageCount(pdf) {
  const info = run('pdfinfo', [pdf]);
  const m = /Pages:\s+(\d+)/.exec(info.stdout ?? '');
  if (m) return Number(m[1]);
  return (readFileSync(pdf, 'latin1').match(/\/Type\s*\/Page[^s]/g) ?? []).length;
}

let failed = 0;
const exporter = exportFirst ? await import('@b32nio/spindle-io/pptx') : undefined;
for (const file of files) {
  let target = path.resolve(file);
  const notes = [];
  if (exporter) {
    const result = await exporter.exportPptx(await exporter.importPptx(new Uint8Array(readFileSync(target))));
    notes.push(...result.warnings);
    mkdirSync(path.join(out, 'exported'), { recursive: true });
    target = path.join(out, 'exported', path.basename(file));
    writeFileSync(target, result.bytes);
  }
  const base = path.basename(target).replace(/\.[^.]+$/, '');
  // a private profile per run, so a LibreOffice already open on the box is neither disturbed nor waited for
  const profile = mkdtempSync(path.join(os.tmpdir(), 'lo-profile-'));
  const conv = run('soffice', [`-env:UserInstallation=file://${profile}`, '--headless', '--convert-to', 'pdf', '--outdir', out, target]);
  const pdf = path.join(out, `${base}.pdf`);
  let pages = 0;
  if (conv.status === 0) {
    try {
      pages = pageCount(pdf);
    } catch {
      pages = 0;
    }
  }
  if (png && pages > 0) run('pdftoppm', ['-png', '-r', '40', pdf, path.join(out, base)]);
  const ok = conv.status === 0 && pages > 0;
  if (!ok) failed++;
  console.log(`${ok ? 'ok  ' : 'FAIL'} ${path.basename(file)}: soffice exit ${conv.status}, ${pages} page(s)${notes.length ? `; ${notes.length} note(s)` : ''}`);
  for (const n of notes) console.log(`       ${n}`);
}
console.log(`output in ${out}`);
process.exit(failed > 0 ? 1 : 0);
