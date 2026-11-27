import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import path from 'node:path';
import { test, expect } from '@playwright/test';
import JSZip from 'jszip';
import { SlidesPage } from '../support/slides';

// Opaque content from an imported PPTX (plan K11). The Open UI arrives with K15, so the generated charts deck is read straight
// from its package here (Node side, JSZip only: no dependency on a built io package, which not every E2E job builds) and loaded
// into the deck through the demo's `__spindle.deck` hook, as the other specs drive the model.

const CHARTS = path.resolve(__dirname, '../../fixtures/pptx/charts.pptx');
const PX = 9525; // EMU per px

interface ChartFrame {
  name: string;
  part: string;
  x: number;
  y: number;
  w: number;
  h: number;
}

/** The chart frames of each slide, read from the package: the `c:chart` frames and the chart parts their relationships name. */
async function chartFramesOf(bytes: Buffer) {
  const zip = await JSZip.loadAsync(bytes);
  const text = (name: string) => zip.file(name)!.async('string');
  const slides: ChartFrame[][] = [];
  const names = Object.keys(zip.files)
    .filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f))
    .sort((a, b) => Number(/\d+/.exec(a)![0]) - Number(/\d+/.exec(b)![0]));
  const parts = new Map<string, Uint8Array>();
  for (const name of names) {
    const xml = await text(name);
    const rels = await text(name.replace('slides/', 'slides/_rels/') + '.rels');
    const target = (id: string) => {
      const rel = [...rels.matchAll(/<Relationship\b[^>]*>/g)]
        .map((m) => m[0])
        .find((r) => r.includes(`Id="${id}"`))!;
      return path.posix.normalize(`ppt/slides/${/Target="([^"]*)"/.exec(rel)![1]}`);
    };
    const frames: ChartFrame[] = [];
    for (const m of xml.matchAll(/<p:graphicFrame>[\s\S]*?<\/p:graphicFrame>/g)) {
      const id = /<c:chart\b[^>]*r:id="([^"]*)"/.exec(m[0])?.[1];
      if (!id) continue;
      const n = (re: RegExp) => Number(re.exec(m[0])![1]) / PX;
      const part = target(id);
      parts.set(part, await zip.file(part)!.async('uint8array'));
      frames.push({
        name: /<p:cNvPr[^>]*name="([^"]*)"/.exec(m[0])![1],
        part,
        x: n(/<a:off x="(\d+)"/),
        y: n(/<a:off [^>]*y="(\d+)"/),
        w: n(/<a:ext cx="(\d+)"/),
        h: n(/<a:ext [^>]*cy="(\d+)"/),
      });
    }
    slides.push(frames);
  }
  const size = /<p:sldSz cx="(\d+)" cy="(\d+)"/.exec(await text('ppt/presentation.xml'))!;
  return { slides, parts, slideSize: { w: Number(size[1]) / PX, h: Number(size[2]) / PX } };
}

test.describe('slides: opaque content', () => {
  test('SL-OPQ-01 the charts deck shows one labelled placeholder per source chart and keeps the chart parts', async ({
    page,
  }) => {
    const source = await chartFramesOf(readFileSync(CHARTS));
    const expected = source.slides.flat().length;
    expect(expected).toBe(4);
    const hash = (b: Uint8Array) => createHash('sha256').update(b).digest('hex');

    const data = {
      id: 'charts',
      title: 'Charts',
      slideSize: source.slideSize,
      slides: source.slides.map((frames, i) => ({
        id: `s${i + 1}`,
        elements: frames.map((f, j) => ({
          id: `s${i + 1}-o${j}`,
          containerId: `s${i + 1}`,
          type: 'opaque',
          kind: 'chart',
          partRefs: [{ hash: hash(source.parts.get(f.part)!), partName: f.part, role: 'chart' }],
          title: f.name,
          x: f.x,
          y: f.y,
          w: f.w,
          h: f.h,
          rotation: 0,
        })),
      })),
      opaqueParts: [...source.parts].map(([partName, bytes]) => ({
        hash: hash(bytes),
        contentType: 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
        bytes,
      })),
    };
    const slideCount = source.slides.length;

    const s = new SlidesPage(page);
    await s.goto();
    await page.evaluate((d) => (window as any).__spindle.deck.setData(d), data as never);
    expect(await s.slideIds()).toHaveLength(slideCount);

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
    for (let n = 1; n <= slideCount; n++) {
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
