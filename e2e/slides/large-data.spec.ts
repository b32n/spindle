import { test, expect } from '@playwright/test';
import { SlidesPage, MOD } from '../support/slides';
import { collectPageErrors } from '../support/file-io';

const N = 100;
const BUILD_BUDGET = 20_000;
const ACTION_BUDGET = 10_000;

test.describe('slides: many elements', () => {
  test('SL-BIG-01 a slide with 100 elements renders, selects, drags, multi-moves and deletes within budget', async ({
    page,
  }, info) => {
    const errors = collectPageErrors(page);
    const s = new SlidesPage(page);
    await s.goto();
    await page.getByRole('button', { name: 'New slide' }).click();
    await expect(s.thumbs).toHaveCount(4);

    const base = (await s.els()).length; // a new slide starts with its layout placeholders
    const TOTAL = base + N;

    // Build the elements through the model, in a 10 x 10 grid of small shapes.
    const t0 = Date.now();
    const slideId = await s.activeSlide();
    await page.evaluate(
      ([id, n]) => {
        const d = (window as any).__spindle.deck;
        for (let i = 0; i < (n as number); i++) {
          d.addElement(id, {
            type: 'shape',
            shape: i % 2 ? 'ellipse' : 'rect',
            x: 20 + (i % 10) * 120,
            y: 20 + Math.floor(i / 10) * 66,
            w: 90,
            h: 50,
          });
        }
      },
      [slideId, N] as const
    );
    await expect(s.elements).toHaveCount(TOTAL, { timeout: BUILD_BUDGET });
    const renderMs = Date.now() - t0;
    const els = await s.els();
    expect(els).toHaveLength(TOTAL);

    // Click-select one in the middle of the grid.
    const t1 = Date.now();
    const target = els[base + 44];
    await s.click(target.id);
    await s.expectSelection([target.id]);
    const selectMs = Date.now() - t1;

    // Drag it by 40 slide px right and 20 down (the stage scale is ~0.74 on a 1280px viewport).
    const t2 = Date.now();
    const c = await s.elementCenter(target.id);
    const { scale } = await s.toPage(0, 0);
    await s.drag(c, { x: c.x + 40 * scale, y: c.y + 20 * scale });
    await expect
      .poll(async () => (await s.el(target.id)).x, { timeout: ACTION_BUDGET })
      .toBeGreaterThan(target.x + 30);
    const moved = await s.el(target.id);
    expect(Math.abs(moved.x - (target.x + 40))).toBeLessThanOrEqual(10);
    expect(Math.abs(moved.y - (target.y + 20))).toBeLessThanOrEqual(10);
    const dragMs = Date.now() - t2;
    // Everything else stayed where it was.
    const after = await s.els();
    expect(
      after.filter((e, i) => i !== base + 44 && (e.x !== els[i].x || e.y !== els[i].y))
    ).toEqual([]);

    // Select all and nudge: every element moves.
    const t3 = Date.now();
    await page.keyboard.press(`${MOD}+a`);
    await expect
      .poll(async () => (await s.selection()).length, { timeout: ACTION_BUDGET })
      .toBe(TOTAL);
    await page.keyboard.press('Shift+ArrowDown');
    await expect
      .poll(async () => (await s.els()).filter((e, i) => e.y === after[i].y + 10).length, {
        timeout: ACTION_BUDGET,
      })
      .toBe(TOTAL);
    const allMs = Date.now() - t3;

    // Delete them all; undo brings all back.
    await page.keyboard.press('Delete');
    await expect(s.elements).toHaveCount(0, { timeout: ACTION_BUDGET });
    await page.keyboard.press(`${MOD}+z`);
    await expect(s.elements).toHaveCount(TOTAL, { timeout: ACTION_BUDGET });

    console.log(
      `[SL-BIG-01] render ${renderMs} ms, select ${selectMs} ms, drag ${dragMs} ms, select-all+move ${allMs} ms`
    );
    info.annotations.push({
      type: 'timings',
      description: `render ${renderMs} ms, select ${selectMs} ms, drag ${dragMs} ms, all ${allMs} ms`,
    });
    expect(errors, 'uncaught page errors').toEqual([]);
  });
});
