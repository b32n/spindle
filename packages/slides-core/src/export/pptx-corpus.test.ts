import { readFileSync, readdirSync } from 'node:fs';
import path from 'node:path';
import { DeckImpl } from '../deck';
import { richTextToPlainText } from '../text/model';
import type { SlideElement } from '../scene/types';
import { exportToPptx, importFromPptx, inspectPptx } from './pptx';

// The whole PPTX corpus through the editor: import, export, import again. The deck keeps its slides and the kind of every element,
// and a second round trip is stable (what the writer states is what the reader gives back).

const dir = path.resolve(__dirname, '../../../../fixtures/pptx');
const files = readdirSync(dir).filter((f) => /\.pptx$/.test(f) && !f.startsWith('edge-'));
const load = (f: string) => new Uint8Array(readFileSync(path.join(dir, f)));
const kinds = (d: DeckImpl) => d.getData().slides.map((s) => s.elements.map((e) => e.type).join(','));
const texts = (d: DeckImpl) =>
  d.getData().slides.map((s) =>
    s.elements.flatMap((e: SlideElement) => (e.type === 'text' || e.type === 'shape') && e.richText ? [richTextToPlainText(e.richText)] : [])
  );

const shapeOf = (d: DeckImpl) =>
  d.getData().slides.map((s) =>
    s.elements.map((e) => ({
      type: e.type,
      frame: [e.x, e.y, e.w, e.h, e.rotation ?? 0].map((n) => Math.round(n * 10) / 10),
      fill: 'fill' in e ? JSON.stringify(e.fill) : undefined,
      cells: e.type === 'table' ? e.cells.map((r) => r.length) : undefined,
      grouped: e.groupId !== undefined,
      bind: e.type === 'line' ? [e.startBind?.anchor, e.endBind?.anchor] : undefined,
    }))
  );

describe('PPTX corpus through the editor', () => {
  it.each(files)('%s keeps its slides, element kinds and text through export', async (f) => {
    const deck = new DeckImpl();
    const { loss } = await importFromPptx(load(f), deck);
    expect(loss.format).toBe('pptx');
    const out = await exportToPptx(deck);
    const again = new DeckImpl();
    await importFromPptx(out.bytes, again);
    expect(again.getData().slides).toHaveLength(deck.getData().slides.length);
    expect(kinds(again)).toEqual(kinds(deck));
    expect(texts(again)).toEqual(texts(deck));
    expect(shapeOf(again)).toEqual(shapeOf(deck));
    // exporting the re-imported deck once more changes nothing the reader sees
    const twice = await exportToPptx(again);
    const third = new DeckImpl();
    await importFromPptx(twice.bytes, third);
    expect(kinds(third)).toEqual(kinds(again));
  });

  it('refuses a file that is not a presentation and leaves the deck as it was', async () => {
    const deck = new DeckImpl();
    const before = JSON.stringify(deck.getData());
    await expect(importFromPptx(load('../pptx/manifest.json'), deck)).rejects.toThrow();
    expect(JSON.stringify(deck.getData())).toBe(before);
  });

  it('reports what a file holds that the editor cannot edit', async () => {
    const report = await inspectPptx(load('transitions-animations.pptx'));
    expect(report.format).toBe('pptx');
    expect(report.items.length).toBeGreaterThan(0);
    const full = await inspectPptx(load('kitchen-sink.pptx'), { detail: 'full' });
    expect(full.items.length).toBeGreaterThanOrEqual((await inspectPptx(load('kitchen-sink.pptx'))).items.length);
  });
});
