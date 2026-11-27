import { readFileSync } from 'node:fs';
import path from 'node:path';
import { detectPptxLoss } from '@b32nio/spindle-io/pptx';
import { PPTX_FIXTURES_DIR } from '../../../../scripts/pptx-fidelity';
import { readParts } from '../../../../scripts/fidelity-core/ooxml-package';
import { SUPPORTED_PRESETS } from '../../../io/src/pptx/slide-scan';
import { SHAPE_PRESETS, shapePrst } from '../scene/shape-presets';

// The preset deck holds every ST_ShapeType value once: the editor's preset names are checked against it, and the
// loss detector's supported set against the editor's library.
const bytes = new Uint8Array(readFileSync(path.join(PPTX_FIXTURES_DIR, 'shapes-presets.pptx')));

/** Predate K5 and are not OOXML names (`rtTriangle`, `rightArrow` and `leftArrow` are): the names are checked through `shapePrst`. */
const LEGACY_NAMES = new Set(['rightTriangle', 'arrowRight', 'arrowLeft']);

describe('preset names against the OOXML preset deck', () => {
  it('every OOXML-named preset is a real ST_ShapeType value', async () => {
    const parts = await readParts(bytes);
    const xml = Object.entries(parts.xml)
      .filter(([name]) => name.startsWith('ppt/slides/slide'))
      .map(([, x]) => x)
      .join('');
    const deck = new Set([...xml.matchAll(/<a:prstGeom prst="([^"]+)"/g)].map((m) => m[1]));
    expect(deck.size).toBe(187);
    expect(SHAPE_PRESETS.filter((p) => !deck.has(shapePrst(p)))).toEqual([]);
    expect([...LEGACY_NAMES].every((p) => deck.has(shapePrst(p as never)))).toBe(true);
  });

  it('the detector\'s supported set is exactly the library plus the line geometries', () => {
    // both spellings of the three legacy names: the editor's own and the OOXML one
    const expected = new Set<string>([...SHAPE_PRESETS, ...SHAPE_PRESETS.map((p) => shapePrst(p)), 'line', 'straightConnector1']);
    expect([...SUPPORTED_PRESETS].sort()).toEqual([...expected].sort());
  });

  it('the detector reports exactly the presets outside the library as unsupported', async () => {
    const report = await detectPptxLoss(bytes, { detail: 'full' });
    const item = report.items.find((i) => i.id === 'pptx.shapes.unsupportedGeometry')!;
    const where = new Set(item.where);
    for (const p of SHAPE_PRESETS) expect(where.has(p)).toBe(false);
    // 187 presets in the deck; the library covers 74 distinct prst names (arrowRight is rightArrow's twin), so the rest are
    // unsupported, less the line and connector presets the detector counts elsewhere
    const distinct = new Set(SHAPE_PRESETS.map((p) => shapePrst(p))).size;
    expect(item.count).toBeLessThanOrEqual(187 - distinct);
    expect(item.count).toBeGreaterThan(187 - distinct - 25);
  });
});
