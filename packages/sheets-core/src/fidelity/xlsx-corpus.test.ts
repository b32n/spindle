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
import { unverifiedClaims } from '../../../../scripts/xlsx-fidelity/claims';
import { CONSTRUCTS } from '../../../../scripts/xlsx-fidelity/constructs';
import {
  XLSX_FIXTURES_DIR,
  injectXlsxMatrix,
  loadXlsxManifest,
  uncoveredXlsxConstructs,
  verifyXlsxClaims,
  xlsxMatrix,
} from '../../../../scripts/xlsx-fidelity';
import {
  isWellFormed,
  readParts,
  relationshipTargets,
} from '../../../../scripts/xlsx-fidelity/ooxml-package';

const MAX_BYTES = 300 * 1024;
const manifest = loadXlsxManifest();
const committed = manifest.fixtures.filter((f) => !f.generatedAtRuntime);
const runtime = manifest.fixtures.filter((f) => f.generatedAtRuntime);
const parts = (file: string) =>
  readParts(new Uint8Array(readFileSync(path.join(XLSX_FIXTURES_DIR, file))));
const notAPackage = (claims: string[]) =>
  claims.some((c) => ['edge.truncatedZip', 'edge.encrypted', 'edge.legacyBinary'].includes(c));

describe('XLSX corpus', () => {
  it('has the planned shape', () => {
    expect(committed.length).toBeGreaterThanOrEqual(26);
    expect(runtime.map((f) => f.id).sort()).toEqual([
      'edge-huge-dimension',
      'edge-zip-bomb',
      'perf-100k-x-10',
    ]);
    for (const id of [
      'edge-truncated-zip',
      'edge-no-workbook',
      'edge-bad-rels',
      'edge-xml-entities',
      'edge-encrypted',
      'edge-legacy-xls',
      'edge-duplicate-sheet-names',
      'edge-index-out-of-range',
      'edge-empty-workbook',
    ]) {
      expect(committed.map((f) => f.id)).toContain(id);
    }
    for (const id of [
      'openpyxl-image',
      'openpyxl-chart',
      'openpyxl-comment',
      'formulas-core',
      'number-format-probe',
      'formula-syntax',
    ]) {
      expect(committed.map((f) => f.id)).toContain(id);
    }
  });

  it('every manifest claim is present in the package XML', async () => {
    expect(await verifyXlsxClaims()).toEqual([]);
  });

  it('every construct in the registry is claimed by at least one fixture', () => {
    expect(uncoveredXlsxConstructs(manifest)).toEqual([]);
  });

  it('records the licence and the producer of every fixture', () => {
    for (const f of manifest.fixtures) {
      expect(f.license).toMatch(
        /^Original content authored for this repository \(MIT, same as the repo\)$/
      );
      expect(f.source.tool.length).toBeGreaterThan(0);
      if (/openpyxl|xlsxwriter|LibreOffice/.test(f.source.tool))
        expect(f.source.version).toMatch(/\d+\.\d+/);
    }
  });

  it('lists every committed file and no more, none over 300 KB', () => {
    const onDisk = readdirSync(XLSX_FIXTURES_DIR)
      .filter((n) => /\.(xlsx|xlsm|xls)$/.test(n))
      .sort();
    expect(onDisk).toEqual(committed.map((f) => f.filename).sort());
    for (const f of committed)
      expect(statSync(path.join(XLSX_FIXTURES_DIR, f.filename)).size).toBeLessThanOrEqual(
        MAX_BYTES
      );
    for (const f of runtime)
      expect(existsSync(path.join(XLSX_FIXTURES_DIR, f.filename))).toBe(false);
  });

  it('points every fixture at a generator that exists', () => {
    for (const f of manifest.fixtures) {
      expect(f.generator).toBeDefined();
      expect(existsSync(path.join(XLSX_FIXTURES_DIR, f.generator as string))).toBe(true);
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
      const sheetNames = [
        ...(p.xml['xl/workbook.xml'] ?? '').matchAll(/<sheet\b[^>]*\bname="([^"]*)"/g),
      ].map((m) => m[1].toLowerCase());
      expect(new Set(sheetNames).size).toBe(sheetNames.length);
      for (const n of sheetNames) expect(n.length).toBeLessThanOrEqual(31);
    }
  );

  it('keeps the README coverage matrix in sync (XLSX_FIDELITY_UPDATE=1 rewrites it)', () => {
    const file = path.join(XLSX_FIXTURES_DIR, 'README.md');
    const current = readFileSync(file, 'utf8');
    const next = injectXlsxMatrix(current, xlsxMatrix(manifest));
    if (process.env.XLSX_FIDELITY_UPDATE === '1' && next !== current) writeFileSync(file, next);
    expect(current === next || process.env.XLSX_FIDELITY_UPDATE === '1').toBe(true);
  });
});

describe('run-time generators', () => {
  const python = spawnSync('python3', ['--version']).status === 0;
  const dir = mkdtempSync(path.join(tmpdir(), 'xlsx-runtime-'));
  afterAll(() => rmSync(dir, { recursive: true, force: true }));

  it.each([['edge-huge-dimension'], ['edge-zip-bomb']])(
    '%s: the generator writes a file that holds its claims',
    async (id) => {
      if (!python) return;
      const entry = runtime.find((f) => f.id === id);
      const script = path.join(XLSX_FIXTURES_DIR, entry?.generator as string);
      const run = spawnSync('python3', [script, '--out', dir, id], {
        encoding: 'utf8',
        timeout: 120_000,
      });
      expect(run.status).toBe(0);
      const file = path.join(dir, entry?.filename as string);
      expect(statSync(file).size).toBeLessThanOrEqual(MAX_BYTES * 2);
      const p = await readParts(new Uint8Array(readFileSync(file)));
      expect(unverifiedClaims(entry?.constructs ?? [], CONSTRUCTS, p)).toEqual([]);
      if (id === 'edge-zip-bomb') expect(p.oversized).toContain('xl/worksheets/sheet1.xml');
    },
    180_000
  );

  it('perf-100k-x-10: the generator can write a smaller sheet with the same shape', async () => {
    if (!python) return;
    const script = path.join(XLSX_FIXTURES_DIR, 'generators', 'build_runtime.py');
    const code = `import sys; sys.path.insert(0, ${JSON.stringify(path.dirname(script))}); import build_runtime as B; B.perf(sys.argv[1], rows=200, cols=10)`;
    const out = path.join(dir, 'perf-small.xlsx');
    expect(spawnSync('python3', ['-I', '-c', code, out]).status).toBe(0);
    const p = await readParts(new Uint8Array(readFileSync(out)));
    expect(unverifiedClaims(['cells.numbers', 'cells.inlineStrings'], CONSTRUCTS, p)).toEqual([]);
    expect(p.xml['xl/worksheets/sheet1.xml']).toContain('<dimension ref="A1:J200"/>');
  });
});
