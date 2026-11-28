/**
 * Local benchmark (plan K17, launch-bar row L17): import time, memory and filmstrip render time of a 100-slide and a
 * 200-slide deck. NOT a CI gate and not part of the jest runs.
 *
 *   npm run build && npm run slides:bench [-- --runs 5] [--out DIR] [--json FILE] [--no-browser]
 *
 * The decks are written by fixtures/pptx/generators/build_bench.py (deterministic, never committed). Each import runs in
 * its own process (`--worker`) so the peak resident set is that run's; the filmstrip time is measured in Chromium
 * against the slides demo, from the file handed to the file chooser until the filmstrip holds every slide. That is the Vite
 * dev server (unminified, with its module loading), so the figure is a conservative upper bound for a built demo; the
 * browser's own memory is not measured (the RSS is the import process's). The medians are checked against the L17
 * budget (helpers.ts) and the table is printed as Markdown; `--no-browser` skips the filmstrip (n/a, not checked).
 */
import { spawnSync } from 'node:child_process';
import { mkdirSync, readFileSync, statSync, writeFileSync } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import type { Browser } from '@playwright/test';
import { demoPage, launch, longestTaskSince, openFile, pageNow, ROOT, serveDemo } from './browser';
import { overBudget, resultsTable, summarize, type DeckResult } from './helpers';

const DECKS = [
  { deck: 'bench-100-slides', slides: 100 },
  { deck: 'bench-200-slides', slides: 200 },
];

/** One import in this process: prints `{importMs, rssMiB}` (the child side of `--worker`). */
async function worker(file: string): Promise<void> {
  const { DeckImpl, importFromPptx } = await import('@b32nio/spindle-slides-core');
  const bytes = new Uint8Array(readFileSync(file));
  const deck = new DeckImpl();
  const t0 = performance.now();
  await importFromPptx(bytes, deck);
  const importMs = performance.now() - t0;
  const rssMiB = process.resourceUsage().maxRSS / 1024; // KiB on Linux
  console.log(JSON.stringify({ importMs, rssMiB, slides: deck.getData().slides.length }));
}

function option(args: string[], name: string): string | undefined {
  const i = args.indexOf(name);
  return i === -1 ? undefined : args[i + 1];
}

async function main(): Promise<void> {
  const args = process.argv.slice(2);
  if (args[0] === '--worker') return worker(args[1]);
  const runs = Number(option(args, '--runs') ?? 5);
  const out = path.resolve(option(args, '--out') ?? path.join(ROOT, 'review-pack-output', 'bench'));
  const jsonFile = option(args, '--json');
  mkdirSync(out, { recursive: true });

  const gen = spawnSync('python3', ['build_bench.py', '--out', out], {
    cwd: path.join(ROOT, 'fixtures/pptx/generators'),
    encoding: 'utf8',
  });
  if (gen.status !== 0) throw new Error(`generator failed: ${gen.stderr}`);

  const noBrowser = args.includes('--no-browser');
  let server: { stop: () => void } | undefined;
  let browser: Browser | undefined;
  let browserVersion: string | null = null;
  const results: DeckResult[] = [];
  try {
    if (!noBrowser) {
      server = await serveDemo();
      browser = await launch();
      browserVersion = browser.version();
    }
    for (const { deck, slides } of DECKS) {
      const file = path.join(out, `${deck}.pptx`);
      const importMs: number[] = [];
      const rssMiB: number[] = [];
      const filmstripMs: number[] = [];
      const longestTaskMs: number[] = [];
      for (let i = 0; i < runs; i++) {
        const w = spawnSync(
          'npx',
          ['tsx', path.join(__dirname, 'benchmark.ts'), '--worker', file],
          { cwd: ROOT, encoding: 'utf8' }
        );
        if (w.status !== 0) throw new Error(`import worker failed: ${w.stderr}`);
        const r = JSON.parse(w.stdout.trim().split('\n').pop() ?? '{}') as {
          importMs: number;
          rssMiB: number;
          slides: number;
        };
        if (r.slides !== slides)
          throw new Error(`${deck}: imported ${r.slides} slides, expected ${slides}`);
        importMs.push(r.importMs);
        rssMiB.push(r.rssMiB);
        if (browser) {
          const page = await demoPage(browser);
          const mark = await pageNow(page);
          let t0 = 0;
          await openFile(page, file, () => (t0 = Date.now()));
          await page.waitForFunction(
            (n) => document.querySelectorAll('[data-slide-thumb]').length >= n,
            slides,
            { timeout: 120_000, polling: 'raf' }
          );
          filmstripMs.push(Date.now() - t0);
          longestTaskMs.push(await longestTaskSince(page, mark));
          await page.close();
        }
      }
      results.push({
        deck,
        slides,
        bytes: statSync(file).size,
        importMs: summarize(importMs),
        rssMiB: summarize(rssMiB),
        filmstripMs: browser ? summarize(filmstripMs) : null,
        longestTaskMs: browser ? summarize(longestTaskMs) : null,
      });
    }
  } finally {
    await browser?.close();
    server?.stop();
  }
  const table = resultsTable(results);
  const over = results.flatMap((r) => overBudget(r));
  console.log(table);
  if (noBrowser)
    console.log(
      '--no-browser: filmstrip and longest task not measured, only import and memory are checked.'
    );
  console.log(over.length ? `OVER BUDGET:\n${over.join('\n')}` : 'Within the L17 budget.');
  if (jsonFile)
    writeFileSync(
      jsonFile,
      JSON.stringify(
        {
          runs,
          node: process.version,
          os: `${os.type()} ${os.release()}`,
          cpus: { count: os.cpus().length, model: os.cpus()[0]?.model ?? 'unknown' },
          browser: browserVersion,
          results,
        },
        null,
        2
      ) + '\n'
    );
  process.exitCode = over.length ? 1 : 0;
}

main().catch((e) => {
  console.error(e);
  process.exit(2);
});
