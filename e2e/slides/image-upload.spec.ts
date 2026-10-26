import { test, expect, type Page } from '@playwright/test';
import { readFileSync } from 'node:fs';
import { SlidesPage, SLIDES_URL } from '../support/slides';
import { collectAlerts, collectPageErrors, fixture } from '../support/file-io';
import { makePng } from '../support/png';

interface ImageEl {
  id: string;
  type: string;
  src: string;
  naturalW: number;
  naturalH: number;
  w: number;
  h: number;
}

test.describe('slides: image upload variants', () => {
  let s: SlidesPage;
  let pageErrors: string[];
  let alerts: string[];
  test.beforeEach(async ({ page }) => {
    pageErrors = collectPageErrors(page);
    alerts = collectAlerts(page);
    s = new SlidesPage(page);
    await s.goto();
    await s.gotoSlide(2);
  });
  test.afterEach(() => {
    expect(pageErrors, 'uncaught page errors').toEqual([]);
    expect(alerts, 'unexpected alert dialogs').toEqual([]);
  });

  const ids = async () => (await s.els()).map((e) => e.id);
  const images = (page: Page, before: string[]) =>
    page.evaluate((known) => {
      const d = (window as any).__spindle.deck;
      return d
        .getElementsForSlide(d.getActiveSlideId())
        .filter((e: any) => !known.includes(e.id))
        .map((e: any) => ({
          id: e.id,
          type: e.type,
          src: e.src,
          naturalW: e.naturalW,
          naturalH: e.naturalH,
          w: e.w,
          h: e.h,
        }));
    }, before) as Promise<ImageEl[]>;

  const variants = [
    ['SL-IMG-01', 'png', 'logo.png', 'image/png'],
    ['SL-IMG-02', 'jpeg', 'logo.jpg', 'image/jpeg'],
    ['SL-IMG-03', 'gif', 'logo.gif', 'image/gif'],
    ['SL-IMG-04', 'webp', 'logo.webp', 'image/webp'],
    ['SL-IMG-05', 'svg', 'logo.svg', 'image/svg+xml'],
  ] as const;
  for (const [id, label, file, mime] of variants) {
    test(`${id} uploading a ${label} inserts an image element with the file embedded as a data URI`, async ({
      page,
    }) => {
      const before = await ids();
      await page.locator('input[type="file"]').setInputFiles(fixture(file));
      await expect.poll(async () => (await images(page, before)).length).toBe(1);
      const [el] = await images(page, before);
      expect(el.type).toBe('image');
      expect(el.src.startsWith(`data:${mime};base64,`)).toBe(true);
      // The embedded bytes are exactly the uploaded file.
      expect(Buffer.from(el.src.split(',')[1], 'base64').equals(readFileSync(fixture(file)))).toBe(
        true
      );
      expect([el.naturalW, el.naturalH]).toEqual([120, 60]);
      await s.expectSelection([el.id]);
      await expect(s.stage.locator(`[data-element-id="${el.id}"]`)).toBeVisible();
    });
  }

  test('SL-IMG-06 a 4000x3000 image is inserted scaled to fit the slide and keeps its aspect ratio', async ({
    page,
  }) => {
    const before = await ids();
    const png = makePng(4000, 3000, (x, y) => [x & 255, y & 255, (x + y) & 255]);
    const started = Date.now();
    await page
      .locator('input[type="file"]')
      .setInputFiles({ name: 'huge.png', mimeType: 'image/png', buffer: png });
    await expect.poll(async () => (await images(page, before)).length).toBe(1);
    expect(Date.now() - started, 'insert time for a 4000x3000 image').toBeLessThan(10_000);
    const [el] = await images(page, before);
    expect([el.naturalW, el.naturalH]).toEqual([4000, 3000]);
    const slide = await page.evaluate(() => (window as any).__spindle.deck.getSlideSize());
    expect(el.w).toBeLessThanOrEqual(slide.w);
    expect(el.h).toBeLessThanOrEqual(slide.h);
    expect(el.w / el.h).toBeCloseTo(4 / 3, 2);
    // Still responsive afterwards.
    await s.click(el.id);
    await page.keyboard.press('Delete');
    await expect.poll(async () => (await images(page, before)).length).toBe(0);
  });

  test('SL-IMG-07 cancelling the file chooser inserts nothing', async ({ page }) => {
    const before = await ids();
    const [chooser] = await Promise.all([
      page.waitForEvent('filechooser'),
      s.toolbar('Image (upload)').click(),
    ]);
    await chooser.setFiles([]); // what the browser does when the user cancels
    // Nothing to wait on for "nothing happened": give any async insert a full render cycle.
    await page.evaluate(
      () => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)))
    );
    expect(await ids()).toEqual(before);
    expect(await s.selection()).toEqual([]);
  });

  test('SL-IMG-08 the upload button opens a chooser that accepts images', async ({ page }) => {
    const [chooser] = await Promise.all([
      page.waitForEvent('filechooser'),
      s.toolbar('Image (upload)').click(),
    ]);
    expect(chooser.isMultiple()).toBe(false);
    await expect(page.locator('input[type="file"]')).toHaveAttribute('accept', 'image/*');
    const before = await ids();
    await chooser.setFiles(fixture('logo.png'));
    await expect.poll(async () => (await images(page, before)).length).toBe(1);
  });

  test('SL-IMG-09 a corrupt "image" (random bytes named .png) inserts nothing and raises no error', async ({
    page,
  }) => {
    // Current behaviour: the decode failure is silent (no message) and the slide is unchanged.
    const before = await ids();
    await page.locator('input[type="file"]').setInputFiles({
      name: 'broken.png',
      mimeType: 'image/png',
      buffer: Buffer.from(Array.from({ length: 2048 }, (_, i) => (i * 31 + 7) & 0xff)),
    });
    await page.evaluate(() => new Promise((r) => setTimeout(r, 500)));
    expect(await ids()).toEqual(before);
  });

  test('SL-IMG-10 a non-image file chosen anyway (bypassing accept) inserts nothing', async ({
    page,
  }) => {
    const before = await ids();
    await page.locator('input[type="file"]').setInputFiles(fixture('notes.txt'));
    await page.evaluate(() => new Promise((r) => setTimeout(r, 500)));
    expect(await ids()).toEqual(before);
  });
});

test.describe('slides: image hooks (no URL fetching by default)', () => {
  test.beforeEach(async ({ page }) => {
    await page.route(/fonts\.b32n\.io/, (r) => r.abort());
  });

  const pngBytes = readFileSync(fixture('logo.png'));
  const slideImages = (page: Page) =>
    page.evaluate(() => {
      const d = (window as any).__spindle.deck;
      return d
        .getElementsForSlide(d.getActiveSlideId())
        .filter((e: any) => e.type === 'image')
        .map((e: any) => e.src as string);
    });

  test('SL-IMG-11 without resolveImage there is no "Image from URL" button and no image request is made', async ({
    page,
  }) => {
    const requested: string[] = [];
    page.on('request', (r) => {
      if (!r.url().startsWith(SLIDES_URL) && !r.url().startsWith('https://fonts.b32n.io/'))
        requested.push(r.url()); // anything leaving the demo origin (except the font mirror)
    });
    const s = new SlidesPage(page);
    await s.goto();
    await expect(s.toolbar('Image (upload)')).toBeVisible();
    await expect(s.toolbar('Image from URL')).toHaveCount(0);
    expect(requested).toEqual([]);
  });

  test('SL-IMG-12 resolveImage: the URL is fetched by the host hook and stored as a data URL', async ({
    page,
  }) => {
    await page.route('**/good/**', (route) =>
      route.fulfill({ status: 200, contentType: 'image/png', body: pngBytes })
    );
    const s = new SlidesPage(page);
    await s.goto('?resolveImage=1');
    await s.gotoSlide(2);
    const before = (await slideImages(page)).length;
    page.once('dialog', (d) => d.accept(`${SLIDES_URL}/good/logo.png`));
    await s.toolbar('Image from URL').click();
    await expect.poll(async () => (await slideImages(page)).length).toBe(before + 1);
    const srcs = await slideImages(page);
    expect(srcs[srcs.length - 1]).toMatch(/^data:image\/png;base64,/);
  });

  test('SL-IMG-13 uploadImage: the element stores the URL the host returned', async ({ page }) => {
    await page.route('**/uploaded/**', (route) =>
      route.fulfill({ status: 200, contentType: 'image/png', body: pngBytes })
    );
    const s = new SlidesPage(page);
    await s.goto('?uploadImage=1');
    await s.gotoSlide(2);
    const before = (await slideImages(page)).length;
    await page.locator('input[type="file"]').setInputFiles(fixture('logo.png'));
    await expect.poll(async () => (await slideImages(page)).length).toBe(before + 1);
    const srcs = await slideImages(page);
    expect(srcs[srcs.length - 1]).toBe(`${SLIDES_URL}/uploaded/logo.png`);
  });

  test('SL-IMG-14 maxImageBytes refuses a larger file', async ({ page }) => {
    const s = new SlidesPage(page);
    await s.goto('?maxImageBytes=1024');
    await s.gotoSlide(2);
    const before = (await slideImages(page)).length;
    await page
      .locator('input[type="file"]')
      .setInputFiles({ name: 'big.png', mimeType: 'image/png', buffer: Buffer.alloc(4096) });
    await page.evaluate(
      () => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)))
    );
    expect((await slideImages(page)).length).toBe(before);
  });
});
