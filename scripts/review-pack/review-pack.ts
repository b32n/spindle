/**
 * Local review pack (plan K17) for Bharat's visual sign-off ledger (launch-bar row L16): for each .pptx of the corpus, the
 * Spindle render (the slides demo in Chromium, one PNG per slide) beside the LibreOffice render (`soffice` to PDF, then
 * `pdftoppm`), one side-by-side PNG per slide, and an `index.md` ledger with a blank verdict column. NOT a CI gate: needs
 * `soffice`, `pdftoppm` and a Chromium (see browser.ts); nothing is committed (the output directory is git-ignored).
 *
 *   npm run slides:review-pack [-- --out DIR] [--max-slides N] [--dpi N] [file.pptx ...]
 *
 * Default files: fixtures/pptx/*.pptx except the edge-* files (they are rejected by design, nothing to look at).
 */
import { spawnSync } from 'node:child_process';
import { mkdirSync, readdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import path from 'node:path';
import type { Page } from '@playwright/test';
import { demoPage, launch, openFile, ROOT, serveDemo } from './browser';
import { pairPages, stem } from './helpers';

const args = process.argv.slice(2);
const option = (name: string): string | undefined => {
  const i = args.indexOf(name);
  return i === -1 ? undefined : args.splice(i, 2)[1];
};
const out = path.resolve(option('--out') ?? path.join(ROOT, 'review-pack-output', 'review'));
const maxSlides = Number(option('--max-slides') ?? 12);
const dpi = option('--dpi') ?? '60';
const corpus = path.join(ROOT, 'fixtures', 'pptx');
const files = args.length
  ? args.map((f) => path.resolve(f))
  : readdirSync(corpus)
      .filter((f) => f.endsWith('.pptx') && !f.startsWith('edge-'))
      .sort()
      .map((f) => path.join(corpus, f));

const run = (cmd: string, argv: string[]) =>
  spawnSync(cmd, argv, { encoding: 'utf8', timeout: 180_000 });
for (const tool of [
  ['soffice', '--version'],
  ['pdftoppm', '-v'],
]) {
  if (run(tool[0], tool.slice(1)).error) {
    console.error(`${tool[0]} is not available: the review pack needs soffice and pdftoppm`);
    process.exit(2);
  }
}

const pad = (n: number, width: number) => String(n).padStart(width, '0');
const png = (dir: string) =>
  readdirSync(dir)
    .filter((f) => f.endsWith('.png'))
    .sort();

/** LibreOffice's pages: `lo/<stem>-NN.png`. Returns the page count (0 when soffice gave no PDF). */
function libreOfficeRender(file: string, dir: string, profile: string): number {
  mkdirSync(dir, { recursive: true });
  const c = run('soffice', [
    `-env:UserInstallation=file://${profile}`,
    '--headless',
    '--convert-to',
    'pdf',
    '--outdir',
    dir,
    file,
  ]);
  const pdf = path.join(dir, `${stem(file)}.pdf`);
  if (c.status !== 0 || !readdirSync(dir).includes(path.basename(pdf))) return 0;
  run('pdftoppm', ['-png', '-r', dpi, '-l', String(maxSlides), pdf, path.join(dir, stem(file))]);
  rmSync(pdf);
  return png(dir).length;
}

/** Spindle's slides: `spindle/<stem>/slide-NNN.png`. Returns the slide count of the deck as imported. */
async function spindleRender(page: Page, file: string, dir: string): Promise<number> {
  mkdirSync(dir, { recursive: true });
  await openFile(page, file);
  await page.waitForFunction(
    (name) =>
      (window as unknown as { __lastOpened?: { fileName: string } }).__lastOpened?.fileName ===
      name,
    path.basename(file),
    { timeout: 60_000 }
  );
  const thumbs = page.locator('[data-slide-thumb]');
  const count = await thumbs.count();
  for (let i = 0; i < Math.min(count, maxSlides); i++) {
    await thumbs.nth(i).click();
    await page.waitForTimeout(150);
    await page
      .locator('[data-slide-stage] [data-slide-id]')
      .first()
      .screenshot({ path: path.join(dir, `slide-${pad(i + 1, 3)}.png`) });
  }
  return count;
}

/** Runs in the page (plain JavaScript, a string so the TypeScript runner adds no helpers to it): draws the two images and returns the PNG. */
const DRAW = `
async ([l, r, text]) => {
  const load = (b) =>
    new Promise((res) => {
      if (!b) return res(null);
      const img = new Image();
      img.onload = () => res(img);
      img.src = \`data:image/png;base64,\${b}\`;
    });
  const [a, b] = await Promise.all([load(l), load(r)]);
  const h = 540;
  const fit = (i) => (i ? Math.round((i.width / i.height) * h) : 480);
  const wa = fit(a);
  const wb = fit(b);
  const c = document.createElement('canvas');
  c.width = wa + wb + 30;
  c.height = h + 30;
  const g = c.getContext('2d');
  g.fillStyle = '#888';
  g.fillRect(0, 0, c.width, c.height);
  g.fillStyle = '#fff';
  g.font = '14px sans-serif';
  g.fillText(\`\${text}   |   Spindle (left)   LibreOffice (right)\`, 8, 18);
  for (const [img, x, w] of [[a, 10, wa], [b, wa + 20, wb]]) {
    if (img) g.drawImage(img, x, 26, w, h);
    else {
      g.fillStyle = '#444';
      g.fillRect(x, 26, w, h);
      g.fillStyle = '#fff';
      g.fillText('no render', x + 10, 50);
    }
  }
  return c.toDataURL('image/png').split(',')[1];
}
`;

/** Two PNGs side by side with a caption, drawn on a canvas in the page (no image library needed). */
async function sideBySide(
  page: Page,
  left: string | null,
  right: string | null,
  caption: string
): Promise<Buffer> {
  const b64 = (f: string | null) => (f ? readFileSync(f).toString('base64') : null);
  const data = await page.evaluate<string>(
    `(${DRAW})(${JSON.stringify([b64(left), b64(right), caption])})`
  );
  return Buffer.from(data, 'base64');
}

async function main(): Promise<void> {
  rmSync(out, { recursive: true, force: true });
  mkdirSync(out, { recursive: true });
  const profile = path.join(out, '.lo-profile');
  const server = await serveDemo();
  const browser = await launch();
  const rows: string[] = [];
  try {
    const page = await demoPage(browser);
    for (const file of files) {
      const name = stem(file);
      process.stdout.write(`${name} ... `);
      let slides = 0;
      let note = '';
      try {
        slides = await spindleRender(page, file, path.join(out, 'spindle', name));
      } catch (e) {
        note = `Spindle: ${(e as Error).message.split('\n')[0]}`;
      }
      const loDir = path.join(out, 'lo', name);
      const pages = libreOfficeRender(file, loDir, profile);
      const sDir = path.join(out, 'spindle', name);
      const sFiles = (() => {
        try {
          return png(sDir);
        } catch {
          return [];
        }
      })();
      const pairs = pairPages(sFiles, png(loDir));
      for (const p of pairs) {
        const buf = await sideBySide(
          page,
          p.spindle && path.join(sDir, p.spindle),
          p.libreOffice && path.join(loDir, p.libreOffice),
          `${name} ${p.page}`
        );
        mkdirSync(path.join(out, 'side-by-side'), { recursive: true });
        writeFileSync(path.join(out, 'side-by-side', `${name}-${pad(p.page, 3)}.png`), buf);
      }
      rows.push(
        `| ${name} | ${slides} | ${pages} | ${slides === pages ? 'yes' : 'NO'} | ${pairs.length} | ${note} |  |  |`
      );
      console.log(`${slides} slides, ${pages} LibreOffice pages, ${pairs.length} pairs`);
    }
  } finally {
    await browser.close();
    server.stop();
    rmSync(profile, { recursive: true, force: true });
  }
  const md = [
    '# PPTX visual review pack',
    '',
    `Slides rendered: the first ${maxSlides} of each deck. Side-by-side images: \`side-by-side/<fixture>-<page>.png\` (Spindle left, LibreOffice right).`,
    '',
    '| Fixture | Spindle slides | LibreOffice pages | Counts equal | Pairs | Note | Verdict (acceptable / listed) | Listed differences |',
    '| --- | --- | --- | --- | --- | --- | --- | --- |',
    ...rows,
    '',
  ].join('\n');
  writeFileSync(path.join(out, 'index.md'), md);
  console.log(`\nWrote ${path.join(out, 'index.md')}`);
}

main().catch((e) => {
  console.error(e);
  process.exit(2);
});
