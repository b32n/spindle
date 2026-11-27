import { readFileSync } from 'node:fs';
import path from 'node:path';
import { test, expect } from '@playwright/test';
import JSZip from 'jszip';
import { importPptx } from '@b32nio/spindle-io/pptx';
import { SlidesPage } from '../support/slides';

// Opaque content from an imported PPTX (plan K11). The Open UI arrives with K15, so the generated charts deck is read by the
// PPTX reader here (Node side) and loaded into the deck through the demo's `__spindle.deck` hook, as the other specs drive the model.

const CHARTS = path.resolve(__dirname, '../../fixtures/pptx/charts.pptx');

/** Chart frames in the source: `c:chart` relationships of the slides, counted from the package, not from the reader. */
async function chartFramesInSource(bytes: Buffer): Promise<number> {
  const zip = await JSZip.loadAsync(bytes);
  let n = 0;
  for (const name of Object.keys(zip.files).filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f)))
    n += (await zip.file(name)!.async('string')).match(/<c:chart\b/g)?.length ?? 0;
  return n;
}

test.describe('slides: opaque content', () => {
  test('SL-OPQ-01 the charts deck shows one labelled placeholder per source chart and keeps the chart parts', async ({
    page,
  }) => {
    const bytes = readFileSync(CHARTS);
    const expected = await chartFramesInSource(bytes);
    expect(expected).toBeGreaterThan(1);
    const imported = await importPptx(new Uint8Array(bytes));

    // what the importer returned, as a deck: the opaque elements of each slide (K15 maps the rest)
    const data = {
      id: 'charts',
      title: 'Charts',
      slideSize: imported.slideSize,
      slides: imported.slides.map((s, i) => ({
        id: `s${i + 1}`,
        elements: (s.elements ?? []).flatMap((e, j) =>
          e.kind !== 'opaque'
            ? []
            : [
                {
                  id: `s${i + 1}-o${j}`,
                  containerId: `s${i + 1}`,
                  type: 'opaque',
                  kind: e.opaqueKind,
                  partRefs: e.parts,
                  title: e.title,
                  alt: e.altText,
                  pptxId: e.pptxId,
                  x: e.frame?.x ?? 0,
                  y: e.frame?.y ?? 0,
                  w: e.frame?.w ?? 100,
                  h: e.frame?.h ?? 100,
                  rotation: 0,
                },
              ]
        ),
      })),
      opaqueParts: imported.opaqueParts,
    };

    const s = new SlidesPage(page);
    await s.goto();
    await page.evaluate((d) => (window as any).__spindle.deck.setData(d), data as never);
    expect(await s.slideIds()).toHaveLength(imported.slides.length);

    // one placeholder per chart frame of the source, over all slides
    const total = await page.evaluate(() => {
      const deck = (window as any).__spindle.deck;
      return deck
        .getSlideIds()
        .flatMap((id: string) => deck.getElementsForSlide(id))
        .filter((e: any) => e.type === 'opaque').length;
    });
    expect(total).toBe(expected);

    // each slide draws its placeholder with the label, and it is a real element: selectable and deletable
    let seen = 0;
    for (let n = 1; n <= imported.slides.length; n++) {
      await s.gotoSlide(n);
      const holders = s.stage.locator('[data-opaque-kind="chart"]');
      const count = await holders.count();
      seen += count;
      for (let k = 0; k < count; k++)
        await expect(holders.nth(k)).toContainText('(kept, not editable)');
    }
    expect(seen).toBe(expected);
    await expect(s.stage.locator('[data-opaque-kind="chart"]').first()).toHaveAccessibleName(
      /^Chart.* \(kept, not editable\)/
    );

    // the parts the chart refers to are in the deck, byte for byte (chart xml, embedded workbook, ...)
    const kept = await page.evaluate(() => {
      const deck = (window as any).__spindle.deck;
      const els = deck
        .getElementsForSlide(deck.getActiveSlideId())
        .filter((e: any) => e.type === 'opaque');
      return els.flatMap((e: any) => e.partRefs.map((r: any) => !!deck.getOpaquePart(r.hash)));
    });
    expect(kept.length).toBeGreaterThan(0);
    expect(kept.every(Boolean)).toBe(true);
  });
});
