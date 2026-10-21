import { readFileSync } from 'node:fs';
import path from 'node:path';
import type { Page } from '@playwright/test';

const FONT_DIR = path.join(__dirname, '..', 'fixtures', 'fonts');
const b64 = (f: string) => readFileSync(path.join(FONT_DIR, f)).toString('base64');

// Every family the demos (and the sample content) ask for resolves to the same committed Liberation
// Sans subset, so text metrics do not depend on which fonts the machine has installed.
const FAMILIES = [
  'Inter',
  'Segoe UI',
  'Roboto',
  'Oxygen',
  'Ubuntu',
  'Cantarell',
  'Arial',
  'Helvetica',
  'Calibri',
  'Times New Roman',
  'Georgia',
  'Courier New',
];

/**
 * Make rendering deterministic for screenshots: embedded fonts (loaded before the app starts, and
 * re-applied once ready), no animations/transitions, hidden caret. Call before `goto`.
 */
export async function prepareVisual(page: Page): Promise<void> {
  const faces = FAMILIES.flatMap((family) => [
    { family, weight: '400', data: b64('LiberationSans-Regular.subset.ttf') },
    { family, weight: '700', data: b64('LiberationSans-Bold.subset.ttf') },
  ]);
  await page.addInitScript((faces) => {
    const install = () => {
      const style = document.createElement('style');
      style.textContent =
        `*,*::before,*::after{animation:none!important;transition:none!important;caret-color:transparent!important;scroll-behavior:auto!important}.selection-caret,.remote-cursor,.remote-cursor-label{visibility:hidden!important}` +
        faces
          .map(
            (f) =>
              `@font-face{font-family:'${f.family}';font-weight:${f.weight};src:url(data:font/ttf;base64,${f.data}) format('truetype')}`
          )
          .join('');
      document.head.appendChild(style);
    };
    if (document.head) install();
    else document.addEventListener('DOMContentLoaded', install);
  }, faces);
}

/** Wait until the embedded fonts are decoded and then let layout/canvas repaint with them. */
export async function fontsReady(page: Page): Promise<void> {
  await page.evaluate(async () => {
    await Promise.all(
      [
        '400 14px Inter',
        '700 14px Inter',
        '400 14px "Segoe UI"',
        '700 14px "Segoe UI"',
        '400 14px Arial',
        '700 14px Arial',
      ].map((f) => document.fonts.load(f))
    );
    await document.fonts.ready;
  });
}

/**
 * After the fonts are ready, force a relayout/repaint (canvas grids drew with a fallback font
 * before the embedded one was decoded) by nudging the viewport, then wait for two frames.
 */
export async function repaint(page: Page): Promise<void> {
  await fontsReady(page);
  const vp = page.viewportSize()!;
  await page.setViewportSize({ width: vp.width + 1, height: vp.height });
  await page.setViewportSize(vp);
  await page.evaluate(
    () => new Promise<void>((r) => requestAnimationFrame(() => requestAnimationFrame(() => r())))
  );
}
