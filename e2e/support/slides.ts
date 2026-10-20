import { expect, type Locator, type Page } from '@playwright/test';

export const SLIDES_URL = 'http://127.0.0.1:5177';
export const MOD = process.platform === 'darwin' ? 'Meta' : 'Control';

export interface ElSnapshot {
  id: string;
  type: string;
  x: number;
  y: number;
  w: number;
  h: number;
  rotation: number;
}

/** Page object for the DOM-based slides editor. */
export class SlidesPage {
  readonly stage: Locator;
  readonly slide: Locator;
  readonly elements: Locator;
  readonly thumbs: Locator;

  constructor(readonly page: Page) {
    this.stage = page.locator('[data-slide-stage]').first();
    this.slide = this.stage.locator('[data-slide-id]').first();
    this.elements = this.stage.locator('[data-element-id]');
    this.thumbs = page.locator('[data-slide-thumb]');
  }

  async goto(query = ''): Promise<void> {
    await this.page.goto(`${SLIDES_URL}/${query}`);
    await expect(this.slide).toBeVisible();
    await this.page.waitForFunction(
      () => !!(window as unknown as { __spindle?: unknown }).__spindle
    );
  }

  toolbar(name: string | RegExp) {
    return this.page.getByRole('button', { name, exact: typeof name === 'string' });
  }

  // ---- geometry: slide space (1280x720) <-> page pixels -------------------
  async toPage(x: number, y: number) {
    const box = await this.slide.boundingBox();
    if (!box) throw new Error('slide not laid out');
    const scale = box.width / 1280;
    return { x: box.x + x * scale, y: box.y + y * scale, scale };
  }
  async elementBox(id: string) {
    const box = await this.stage.locator(`[data-element-id="${id}"]`).boundingBox();
    if (!box) throw new Error(`element ${id} not rendered`);
    return box;
  }
  async elementCenter(id: string) {
    const b = await this.elementBox(id);
    return { x: b.x + b.width / 2, y: b.y + b.height / 2 };
  }

  // ---- model --------------------------------------------------------------
  async els(): Promise<ElSnapshot[]> {
    return this.page.evaluate(() => {
      const d = (window as any).__spindle.deck;
      return d
        .getElementsForSlide(d.getActiveSlideId())
        .map((e: any) => ({
          id: e.id,
          type: e.type,
          x: e.x,
          y: e.y,
          w: e.w,
          h: e.h,
          rotation: e.rotation ?? 0,
        }));
    });
  }
  async el(id: string): Promise<ElSnapshot> {
    const e = (await this.els()).find((x) => x.id === id);
    if (!e) throw new Error(`no element ${id}`);
    return e;
  }
  async firstOfType(type: string): Promise<ElSnapshot> {
    const e = (await this.els()).find((x) => x.type === type);
    if (!e) throw new Error(`no ${type} element`);
    return e;
  }
  async selection(): Promise<string[]> {
    return this.page.evaluate(() => (window as any).__spindle.deck.getSelection().elementIds);
  }
  async slideIds(): Promise<string[]> {
    return this.page.evaluate(() => (window as any).__spindle.deck.getSlideIds());
  }
  async activeSlide(): Promise<string> {
    return this.page.evaluate(() => (window as any).__spindle.deck.getActiveSlideId());
  }
  /** Open slide `n` (1-based) via the filmstrip and wait for the stage to render it. */
  async gotoSlide(n: number) {
    await this.page.getByRole('button', { name: `Go to slide ${n}` }).click();
    const id = (await this.slideIds())[n - 1];
    await expect(this.slide).toHaveAttribute('data-slide-id', id);
  }
  async expectSelection(ids: string[]) {
    await expect
      .poll(async () => (await this.selection()).slice().sort())
      .toEqual(ids.slice().sort());
  }

  async click(id: string, modifiers: readonly string[] = []) {
    const c = await this.elementCenter(id);
    for (const m of modifiers) await this.page.keyboard.down(m);
    await this.page.mouse.click(c.x, c.y);
    for (const m of [...modifiers].reverse()) await this.page.keyboard.up(m);
  }
  async drag(from: { x: number; y: number }, to: { x: number; y: number }) {
    await this.page.mouse.move(from.x, from.y);
    await this.page.mouse.down();
    await this.page.mouse.move((from.x + to.x) / 2, (from.y + to.y) / 2, { steps: 4 });
    await this.page.mouse.move(to.x, to.y, { steps: 4 });
    await this.page.mouse.up();
  }
}
