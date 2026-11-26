import { spawnSync } from 'node:child_process';
import {
  existsSync,
  mkdtempSync,
  readFileSync,
  readdirSync,
  rmSync,
  statSync,
  writeFileSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { unverifiedClaims } from '../../../../scripts/fidelity-core/claims';
import { CONSTRUCTS } from '../../../../scripts/pptx-fidelity/constructs';
import {
  PPTX_FIXTURES_DIR,
  injectPptxMatrix,
  loadPptxManifest,
  uncoveredPptxConstructs,
  verifyPptxClaims,
  pptxMatrix,
} from '../../../../scripts/pptx-fidelity';
import {
  isWellFormed,
  readParts,
  relationshipTargets,
} from '../../../../scripts/fidelity-core/ooxml-package';
import { maxGroupDepth, relationshipCycles } from '../../../../scripts/pptx-fidelity/structure';

const MAX_BYTES = 300 * 1024;
const manifest = loadPptxManifest();
const committed = manifest.fixtures.filter((f) => !f.generatedAtRuntime);
const runtime = manifest.fixtures.filter((f) => f.generatedAtRuntime);
const parts = (file: string) =>
  readParts(new Uint8Array(readFileSync(path.join(PPTX_FIXTURES_DIR, file))));
const notAPackage = (claims: string[]) =>
  claims.some((c) => ['edge.truncatedZip', 'edge.encrypted', 'edge.legacyBinary'].includes(c));
const count = (p: Awaited<ReturnType<typeof parts>>, id: string) =>
  CONSTRUCTS.find((c) => c.id === id)?.count(p);

describe('PPTX corpus', () => {
  it('has the planned shape', () => {
    expect(committed.length).toBeGreaterThanOrEqual(21);
    expect(runtime.map((f) => f.id)).toEqual(['edge-1000-slides']);
    for (const id of [
      'edge-truncated-zip',
      'edge-no-presentation',
      'edge-slide-missing',
      'edge-bad-rels',
      'edge-invalid-xml',
      'edge-deep-groups',
      'edge-relationship-cycles',
      'edge-encrypted',
      'edge-legacy-ppt',
      'text-basic',
      'text-rich',
      'layouts-masters',
      'themes',
      'shapes-presets',
      'shapes-custgeom',
      'fills-lines',
      'effects',
      'sizes-16x9',
      'sizes-4x3',
      'sizes-a4',
      'sizes-custom',
      'sizes-portrait',
    ]) {
      expect(committed.map((f) => f.id)).toContain(id);
    }
  });

  it('every manifest claim is present in the package XML', async () => {
    expect(await verifyPptxClaims()).toEqual([]);
  });

  it('every construct in the registry is claimed by at least one fixture', () => {
    expect(uncoveredPptxConstructs(manifest)).toEqual([]);
  });

  it('records the licence and the producer of every fixture', () => {
    for (const f of manifest.fixtures) {
      expect(f.license).toBe(
        'Original content authored for this repository (MIT, same as the repo)'
      );
      expect(f.source.tool.length).toBeGreaterThan(0);
      if (/python-pptx|LibreOffice/.test(f.source.tool))
        expect(f.source.version).toMatch(/\d+\.\d+/);
    }
  });

  it('lists every committed file and no more, none over 300 KB', () => {
    const onDisk = readdirSync(PPTX_FIXTURES_DIR)
      .filter((n) => /\.(pptx|pptm|ppt)$/.test(n))
      .sort();
    expect(onDisk).toEqual(committed.map((f) => f.filename).sort());
    for (const f of committed)
      expect(statSync(path.join(PPTX_FIXTURES_DIR, f.filename)).size).toBeLessThanOrEqual(
        MAX_BYTES
      );
    for (const f of runtime)
      expect(existsSync(path.join(PPTX_FIXTURES_DIR, f.filename))).toBe(false);
  });

  it('points every fixture at a generator that exists', () => {
    for (const f of manifest.fixtures) {
      expect(f.generator).toBeDefined();
      expect(existsSync(path.join(PPTX_FIXTURES_DIR, f.generator as string))).toBe(true);
    }
  });

  it.each(committed.map((f) => [f.id, f] as const))(
    '%s: is a package (or deliberately not), with normalised core properties and sound structure',
    async (_id, f) => {
      const p = await parts(f.filename);
      if (notAPackage(f.constructs)) {
        expect(p.kind).not.toBe('zip');
        return;
      }
      expect(p.kind).toBe('zip');
      expect(p.names).toContain('[Content_Types].xml');
      const core = p.xml['docProps/core.xml'];
      if (core) {
        expect(core).not.toMatch(/<dc:creator>[^<]+/);
        expect(core).not.toMatch(/<cp:lastModifiedBy>[^<]+/);
        expect(core).toContain('2026-01-15T09:00:00Z');
        expect(
          [...core.matchAll(/<dcterms:(?:created|modified)\b[^>]*>([^<]*)</g)].every(
            (m) => m[1] === '2026-01-15T09:00:00Z'
          )
        ).toBe(true);
      }
      expect(p.xml['docProps/app.xml'] ?? '').not.toMatch(/<(?:Company|Manager)>[^<]+/);
      if (f.constructs.some((c) => c.startsWith('edge.'))) return; // damaged on purpose
      for (const [name, xml] of Object.entries(p.xml))
        expect({ name, ok: isWellFormed(xml) }).toEqual({ name, ok: true });
      for (const [name, xml] of Object.entries(p.xml).filter(([n]) => n.endsWith('.rels'))) {
        const ids = relationshipTargets(name, xml).map((r) => r.id);
        expect(new Set(ids).size).toBe(ids.length);
        for (const r of relationshipTargets(name, xml))
          if (r.target !== null)
            expect({ name, target: r.target, exists: p.names.includes(r.target) }).toEqual({
              name,
              target: r.target,
              exists: true,
            });
      }
      const types = p.xml['[Content_Types].xml'];
      const defaults = new Set(
        [...types.matchAll(/<Default\b[^>]*Extension="([^"]+)"/g)].map((m) => m[1].toLowerCase())
      );
      const overrides = new Set(
        [...types.matchAll(/<Override\b[^>]*PartName="\/([^"]+)"/g)].map((m) => m[1])
      );
      for (const n of p.names.filter((x) => x !== '[Content_Types].xml'))
        expect({
          n,
          typed: overrides.has(n) || defaults.has(n.split('.').pop()?.toLowerCase() ?? ''),
        }).toEqual({ n, typed: true });
      // ids: slides, masters and layouts each have unique ids, masters and layouts share one id space
      const pres = p.xml['ppt/presentation.xml'];
      const slideIds = [...pres.matchAll(/<p:sldId\b[^>]*\bid="(\d+)"/g)].map((m) => m[1]);
      expect(new Set(slideIds).size).toBe(slideIds.length);
      const layoutIds = Object.entries(p.xml)
        .filter(([n]) => /^ppt\/slideMasters\/slideMaster\d+\.xml$/.test(n))
        .flatMap(([, x]) => [...x.matchAll(/<p:sldLayoutId\b[^>]*\bid="(\d+)"/g)].map((m) => m[1]));
      const masterIds = [...pres.matchAll(/<p:sldMasterId\b[^>]*\bid="(\d+)"/g)].map((m) => m[1]);
      const shared = [...layoutIds, ...masterIds];
      expect(new Set(shared).size).toBe(shared.length);
      // every slide shape id is unique within its slide
      for (const [name, xml] of Object.entries(p.xml).filter(([n]) =>
        /^ppt\/slides\/slide\d+\.xml$/.test(n)
      )) {
        const ids = [...xml.matchAll(/<p:cNvPr\b[^>]*\bid="(\d+)"/g)].map((m) => m[1]);
        expect({ name, unique: new Set(ids).size === ids.length }).toEqual({ name, unique: true });
      }
    }
  );

  it('keeps the preset-geometry deck complete: all 187 ST_ShapeType values, once each as a shape', async () => {
    const p = await parts('shapes-presets.pptx');
    expect(count(p, 'shapes.presetKinds')).toBe(187);
  });

  it('keeps the deep-group fixture 60 levels deep and the cycle fixture cyclic', async () => {
    const deep = await parts('edge-deep-groups.pptx');
    expect(maxGroupDepth(deep.xml['ppt/slides/slide1.xml'])).toBe(60);
    expect(relationshipCycles(await parts('edge-relationship-cycles.pptx'))).toBe(2);
    expect(relationshipCycles(await parts('text-basic.pptx'))).toBe(0);
  });

  it('classifies each slide-size deck as its own size', async () => {
    for (const [file, id] of [
      ['sizes-4x3.pptx', 'sizes.4x3'],
      ['sizes-a4.pptx', 'sizes.a4'],
      ['sizes-custom.pptx', 'sizes.custom'],
      ['sizes-portrait.pptx', 'sizes.portrait'],
      ['sizes-16x9.pptx', 'sizes.16x9'],
    ] as const) {
      const p = await parts(file);
      const sizes = CONSTRUCTS.filter((c) => c.id.startsWith('sizes.') && c.count(p) > 0).map(
        (c) => c.id
      );
      expect({ file, sizes }).toEqual({ file, sizes: [id] });
    }
  });

  it('keeps the README coverage matrix in sync (PPTX_FIDELITY_UPDATE=1 rewrites it)', () => {
    const file = path.join(PPTX_FIXTURES_DIR, 'README.md');
    const current = readFileSync(file, 'utf8');
    const next = injectPptxMatrix(current, pptxMatrix(manifest));
    if (process.env.PPTX_FIDELITY_UPDATE === '1' && next !== current) writeFileSync(file, next);
    expect(current === next || process.env.PPTX_FIDELITY_UPDATE === '1').toBe(true);
  });
});

describe('run-time generators', () => {
  // Without python3 the tests fail on CI and are skipped elsewhere (never a silent pass).
  const python = spawnSync('python3', ['--version']).status === 0;
  const itPy = python || process.env.CI ? it : it.skip;
  const dir = mkdtempSync(path.join(tmpdir(), 'pptx-runtime-'));
  afterAll(() => rmSync(dir, { recursive: true, force: true }));

  itPy(
    'edge-1000-slides: the generator writes a deck that holds its claims',
    async () => {
      const entry = runtime.find((f) => f.id === 'edge-1000-slides');
      const script = path.join(PPTX_FIXTURES_DIR, entry?.generator as string);
      const run = spawnSync('python3', [script, '--out', dir, 'edge-1000-slides'], {
        encoding: 'utf8',
        timeout: 120_000,
      });
      expect(run.status).toBe(0);
      const file = path.join(dir, entry?.filename as string);
      expect(statSync(file).size).toBeGreaterThan(MAX_BYTES);
      expect(statSync(file).size).toBeLessThanOrEqual(MAX_BYTES * 5);
      const p = await readParts(new Uint8Array(readFileSync(file)));
      expect(unverifiedClaims(entry?.constructs ?? [], CONSTRUCTS, p)).toEqual([]);
      expect(p.names.filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))).toHaveLength(1000);
    },
    180_000
  );

  itPy('edge-1000-slides: the generator can write a smaller deck with the same shape', async () => {
    const script = path.join(PPTX_FIXTURES_DIR, 'generators', 'build_runtime.py');
    const code = `import sys; sys.path.insert(0, ${JSON.stringify(path.dirname(script))}); import build_runtime as B; B.many_slides(sys.argv[1], count=7)`;
    const out = path.join(dir, 'many-small.pptx');
    expect(spawnSync('python3', ['-I', '-c', code, out]).status).toBe(0);
    const p = await readParts(new Uint8Array(readFileSync(out)));
    expect(unverifiedClaims(['slides.multiple', 'text.basic'], CONSTRUCTS, p)).toEqual([]);
    expect(p.xml['ppt/presentation.xml'].match(/<p:sldId\b/g)).toHaveLength(7);
  });
});

describe('generators', () => {
  const python = spawnSync('python3', ['--version']).status === 0;
  const itPy = python || process.env.CI ? it : it.skip;
  const dir = mkdtempSync(path.join(tmpdir(), 'pptx-rebuild-'));
  afterAll(() => rmSync(dir, { recursive: true, force: true }));

  itPy(
    'rebuild the hand-authored and edge files byte for byte',
    () => {
      const script = path.join(PPTX_FIXTURES_DIR, 'generators', 'build_all.py');
      const run = spawnSync('python3', [script, '--out', dir, '--stdlib'], {
        encoding: 'utf8',
        timeout: 120_000,
      });
      expect(run.status).toBe(0);
      const rebuilt = readdirSync(dir).sort();
      expect(rebuilt.length).toBeGreaterThan(15);
      for (const name of rebuilt) {
        expect({
          name,
          same: readFileSync(path.join(dir, name)).equals(
            readFileSync(path.join(PPTX_FIXTURES_DIR, name))
          ),
        }).toEqual({ name, same: true });
      }
      const programmatic = committed
        .filter((f) => f.source.tool.startsWith('programmatic'))
        .map((f) => f.filename)
        .sort();
      expect(rebuilt).toEqual(programmatic);
    },
    180_000
  );
});
