import { readFileSync } from 'node:fs';
import path from 'node:path';
import type { Page } from '@playwright/test';

const FONT_DIR = path.join(__dirname, '..', 'fixtures', 'fonts');
const b64 = (f: string) => readFileSync(path.join(FONT_DIR, f)).toString('base64');

// Unique name that no system font can share; the stylesheet forces the DOM onto it.
const UNIQUE_FAMILY = 'SpindleVisualSans';

// Canvas paths (sheets/slides) take their font stack from the app theme, which cannot be overridden
// from CSS, so the demo family names stay registered too. fontsReady() asserts that every one of
// them (unique name included) is actually loaded from the embedded data.
// Every family the demos (and the sample content) ask for resolves to the same committed Liberation
// Sans subset, so text metrics do not depend on which fonts the machine has installed.
const FAMILIES = [
  UNIQUE_FAMILY,
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
        `*,*::before,*::after{font-family:'SpindleVisualSans',sans-serif!important}` +
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

/**
 * Wait until the embedded fonts are decoded, then fail clearly unless every embedded face (the
 * unique family and the demo-name aliases) reports `loaded` in `document.fonts`.
 */
export async function fontsReady(page: Page): Promise<void> {
  const missing = await page.evaluate(
    async ({ families }) => {
      await Promise.all(
        families.flatMap((f) =>
          [`400 14px "${f}"`, `700 14px "${f}"`].map((d) => document.fonts.load(d))
        )
      );
      await document.fonts.ready;
      const loaded = new Set<string>();
      document.fonts.forEach((face) => {
        if (face.status === 'loaded')
          loaded.add(`${face.family.replace(/["']/g, '')}/${face.weight}`);
      });
      return families
        .flatMap((f) => ['400', '700'].map((w) => `${f}/${w}`))
        .filter((k) => !loaded.has(k));
    },
    { families: FAMILIES }
  );
  if (missing.length) throw new Error(`Embedded visual fonts did not load: ${missing.join(', ')}`);
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
