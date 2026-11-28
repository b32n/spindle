/** Shared by the local review pack and the benchmark (plan K17): serve the slides demo, drive it with Playwright's Chromium. */
import { spawn, type ChildProcess } from 'node:child_process';
import { existsSync } from 'node:fs';
import path from 'node:path';
import { chromium, type Browser, type Page } from '@playwright/test';

export const ROOT = path.resolve(__dirname, '..', '..');
export const PORT = 5188;
export const DEMO_URL = `http://127.0.0.1:${PORT}`;

/** The browser: `SPINDLE_CHROME`, else /usr/bin/google-chrome when it exists, else the Chromium Playwright installed. */
export function chromePath(): string | undefined {
  const env = process.env.SPINDLE_CHROME;
  if (env) return env;
  return existsSync('/usr/bin/google-chrome') ? '/usr/bin/google-chrome' : undefined;
}

export async function launch(): Promise<Browser> {
  return chromium.launch({ executablePath: chromePath(), args: ['--no-sandbox'] });
}

/** `vite examples/slides-demo` on a fixed port, stopped by `stop()` (only the process started here is signalled). */
export async function serveDemo(): Promise<{ stop: () => void }> {
  const child: ChildProcess = spawn(
    'npx',
    ['vite', 'examples/slides-demo', '--port', String(PORT), '--strictPort', '--host', '127.0.0.1'],
    { cwd: ROOT, stdio: 'ignore', detached: true }
  );
  const stop = () => {
    try {
      if (child.pid) process.kill(-child.pid); // the whole group npx started, by the id of this child
    } catch {
      /* already gone */
    }
  };
  for (let i = 0; i < 120; i++) {
    try {
      if ((await fetch(DEMO_URL)).ok) return { stop };
    } catch {
      /* not up yet */
    }
    await new Promise((r) => setTimeout(r, 500));
  }
  stop();
  throw new Error('the slides demo did not start');
}

/** A page of the demo with the web-font request blocked (no network) and every import confirmation skipped. */
export async function demoPage(browser: Browser): Promise<Page> {
  const page = await browser.newPage({ viewport: { width: 1600, height: 900 } });
  await page.route(/fonts\.b32n\.io/, (r) => r.abort());
  page.on('dialog', (d) => void d.accept());
  await page.addInitScript(() => {
    const w = window as unknown as { __tasks: { start: number; duration: number }[] };
    w.__tasks = [];
    new PerformanceObserver((list) => {
      for (const e of list.getEntries())
        w.__tasks.push({ start: e.startTime, duration: e.duration });
    }).observe({ type: 'longtask', buffered: true });
  });
  await page.goto(`${DEMO_URL}/?importConfirmation=never`);
  await page.waitForFunction(() => !!(window as unknown as { __spindle?: unknown }).__spindle);
  return page;
}

/** The page clock now: the mark for `longestTaskSince`, taken just before a file is chosen. */
export const pageNow = (page: Page): Promise<number> => page.evaluate(() => performance.now());

/** The longest main-thread task (ms) that started at or after the mark (the demo's own load is not counted). */
export function longestTaskSince(page: Page, mark: number): Promise<number> {
  return page.evaluate((m) => {
    const tasks = (window as unknown as { __tasks: { start: number; duration: number }[] }).__tasks;
    return tasks.reduce((max, t) => (t.start >= m ? Math.max(max, t.duration) : max), 0);
  }, mark);
}

/**
 * Open a file through the toolbar's Open button (or the "More tools" menu when the toolbar has collapsed it);
 * `onChosen` runs right before the file is handed to the file chooser (the benchmark starts its clock there).
 */
export async function openFile(page: Page, file: string, onChosen?: () => void): Promise<void> {
  const name = 'Open file (.pptx)';
  const button = page.getByRole('button', { name, exact: true });
  const [chooser] = await Promise.all([
    page.waitForEvent('filechooser'),
    (async () => {
      if ((await button.count()) === 0) {
        await page.getByRole('button', { name: 'More tools', exact: true }).click();
        const menu = page.getByRole('menu');
        await menu
          .getByRole('button', { name })
          .or(menu.getByRole('menuitem', { name }))
          .first()
          .click();
      } else await button.click();
    })(),
  ]);
  onChosen?.();
  await chooser.setFiles(file);
}
