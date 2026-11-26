import { readFileSync } from 'node:fs';
import path from 'node:path';
import { detectPptxLoss } from '@b32nio/spindle-io/pptx';
import { readParts } from '../../../../scripts/fidelity-core/ooxml-package';
import { CONSTRUCTS } from '../../../../scripts/pptx-fidelity/constructs';
import { PPTX_FIXTURES_DIR, loadPptxManifest } from '../../../../scripts/pptx-fidelity';

/**
 * The manifest claims (verified against the package text by the construct probes) are the ground truth of what each
 * deck contains. The XML-based detector must agree with them: for every loss id below, a deck reports the id exactly
 * when it claims one of the constructs the id covers. The ids not listed (fonts, masters' colour map, equations in
 * other namespaces...) have no one-to-one claim and are tested in the io package with synthetic decks.
 */
const COVERS: Record<string, string[]> = {
  'pptx.charts': ['charts.bar', 'charts.line', 'charts.pie', 'charts.scatter'],
  'pptx.smartArt': ['smartart.diagram'],
  'pptx.media': ['media.audio', 'media.video'],
  'pptx.embeddedObjects': ['objects.ole'],
  'pptx.transitions': ['transitions.basic'],
  'pptx.animations': ['animations.timing'],
  'pptx.comments': ['comments.legacy', 'comments.modern'],
  'pptx.macros': ['macros.vba'],
  'pptx.hiddenSlides': ['slides.hidden'],
  'pptx.notes': ['notes.slides'],
  'pptx.images.unsupportedFormat': ['images.emf', 'images.wmf'],
  'pptx.shapes.custGeom': ['shapes.custGeom'],
  'pptx.masters.multiple': ['masters.multiple'],
  'pptx.masters.handoutMaster': ['masters.handoutMaster'],
  'pptx.masters.notesMaster': ['notes.master'],
  'pptx.theme.colorMap': ['theme.colorMap'],
  'pptx.theme.fontScheme': ['theme.fontScheme'],
  'pptx.text.bullets.custom': ['text.bulletsCustom'],
  'pptx.shapes.connectors.elbowCurved': ['connectors.elbow', 'connectors.curved'],
  'pptx.tables.styles': ['tables.styleRef', 'tables.headerRow'],
  'pptx.theme.effectScheme': ['theme.effectScheme'],
  'pptx.text.autofit': ['text.autofit', 'text.autofitShape'],
  'pptx.text.bullets.picture': ['text.bulletsPicture'],
  'pptx.text.tabStops': ['text.tabStops'],
  'pptx.text.columns': ['text.columns'],
  'pptx.text.vertical': ['text.vertical'],
  'pptx.text.fields': ['text.fields'],
  'pptx.text.wordArt': ['objects.wordArt'],
  'pptx.text.rtl': ['text.rtl'],
  'pptx.text.spacing': ['text.charSpacing', 'text.caps'],
  'pptx.text.baseline': ['text.baseline'],
  'pptx.text.lang': ['text.lang'],
  'pptx.shapes.effects': [
    'effects.shadow',
    'effects.glow',
    'effects.reflection',
    'effects.softEdge',
    'effects.blur',
  ],
  'pptx.shapes.gradientFill': ['fills.gradient'],
  'pptx.shapes.patternFill': ['fills.pattern'],
  'pptx.shapes.groups.nested': ['groups.nested'],
  'pptx.shapes.scene3d': ['effects.scene3d', 'effects.sp3d'],
  'pptx.images.crop': ['images.crop'],
  'pptx.images.recolor': ['images.recolor'],
  'pptx.images.linked': ['images.linked'],
  'pptx.images.svg': ['images.svg'],
  'pptx.tables.merge': ['tables.merge'],
  'pptx.tables.banding': ['tables.banding'],
  'pptx.sections': ['sections.list'],
  'pptx.customShows': ['sections.customShows'],
  'pptx.ink': ['objects.ink'],
  'pptx.equations': ['objects.equations'],
  'pptx.actions': ['objects.actions'],
  'pptx.encrypted': ['edge.encrypted'],
  'pptx.legacyBinary': ['edge.legacyBinary'],
};

/**
 * Differences that are deliberate: the detector's definition is wider ("detector") or narrower ("claim") than the probe's,
 * for the stated reason. Each is asserted to hold, so a stale entry fails.
 */
const KNOWN_DIFFERENCES: Record<string, 'detector' | 'claim'> = {
  // a bullet character other than the round one, not only bullet colour and size, counts as a custom bullet (not claimed)
  'text-basic:pptx.text.bullets.custom': 'detector',
  'kitchen-sink:pptx.text.bullets.custom': 'detector',
  // preset connector geometries on ordinary shapes (the preset sampler) and connectors in the lines deck
  'shapes-presets:pptx.shapes.connectors.elbowCurved': 'detector',
  'fills-lines:pptx.shapes.connectors.elbowCurved': 'detector',
  // gradient fills in table cells count too (the fills probe looks at shape properties only)
  'tables:pptx.shapes.gradientFill': 'detector',
  // only the themes a slide master points at count; the claim looks at every theme part (here the notes master's)
  'libreoffice-deck:pptx.theme.fontScheme': 'claim',
};

describe('the detector agrees with the manifest claims', () => {
  const manifest = loadPptxManifest();
  const committed = manifest.fixtures.filter((f) => !f.generatedAtRuntime);

  it.each(committed.map((f) => [f.id, f.filename] as const))('%s', async (id, filename) => {
    const bytes = new Uint8Array(readFileSync(path.join(PPTX_FIXTURES_DIR, filename)));
    const parts = await readParts(bytes);
    const report = await detectPptxLoss(bytes, { detail: 'full' });
    const counts = new Map(report.items.map((i) => [i.id, i.count]));
    const disagreements: string[] = [];
    for (const [lossId, claims] of Object.entries(COVERS)) {
      if (`${id}:${lossId}` in KNOWN_DIFFERENCES) continue;
      const claimed = claims.reduce(
        (n, c) => n + (CONSTRUCTS.find((k) => k.id === c)?.count(parts) ?? 0),
        0
      );
      if (claimed > 0 !== (counts.get(lossId) ?? 0) > 0)
        disagreements.push(`${lossId} (claims ${claimed}, detector ${counts.get(lossId) ?? 0})`);
    }
    expect(disagreements).toEqual([]);
    for (const [known, side] of Object.entries(KNOWN_DIFFERENCES))
      if (known.startsWith(`${id}:`)) {
        const found = counts.get(known.slice(id.length + 1)) ?? 0;
        if (side === 'detector') expect(found).toBeGreaterThan(0);
        else expect(found).toBe(0);
      }
  });

  it('covers only registered constructs and known loss ids', () => {
    const registry = new Set(CONSTRUCTS.map((c) => c.id));
    for (const claims of Object.values(COVERS)) for (const c of claims) expect(registry.has(c)).toBe(true);
  });
});
