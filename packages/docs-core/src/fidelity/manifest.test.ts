import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { ManifestError, discoverFixtures, loadManifest, validateManifest } from '../../../../scripts/docx-fidelity/manifest';

const entry = (over: Record<string, unknown> = {}) => ({
  id: 'one',
  filename: 'one.docx',
  title: 'One',
  source: { tool: 'test', os: 'linux' },
  license: 'original',
  constructs: ['headings'],
  notes: 'n',
  ...over,
});

describe('manifest', () => {
  let dir: string;
  beforeEach(() => {
    dir = mkdtempSync(path.join(tmpdir(), 'docx-manifest-'));
    writeFileSync(path.join(dir, 'one.docx'), 'x');
    writeFileSync(path.join(dir, 'two.docx'), 'x');
  });
  afterEach(() => rmSync(dir, { recursive: true, force: true }));
  const write = (m: unknown) => writeFileSync(path.join(dir, 'manifest.json'), JSON.stringify(m));

  it('accepts a valid manifest', () => {
    expect(validateManifest({ version: 1, fixtures: [entry(), entry({ id: 'two', filename: 'two.docx', expectImportError: true })] }, dir)).toEqual([]);
  });

  it('rejects a manifest that is not an object with fixtures', () => {
    expect(validateManifest(null, dir)).toHaveLength(1);
    expect(validateManifest({ fixtures: 3 }, dir)).toHaveLength(1);
  });

  it.each([
    [{ version: 2, fixtures: [] }, /unsupported manifest version/],
    [{ version: 1, fixtures: [entry({ title: '' })] }, /"title" must be a non-empty string/],
    [{ version: 1, fixtures: [entry({ id: 'Bad_Id' })] }, /kebab-case/],
    [{ version: 1, fixtures: [entry(), entry()] }, /duplicate id/],
    [{ version: 1, fixtures: [entry({ filename: '../x.docx' })] }, /plain \*\.docx name/],
    [{ version: 1, fixtures: [entry({ filename: 'missing.docx' })] }, /does not exist/],
    [{ version: 1, fixtures: [entry(), entry({ id: 'other' })] }, /duplicate filename/],
    [{ version: 1, fixtures: [entry({ source: { tool: 'x' } })] }, /"source" needs/],
    [{ version: 1, fixtures: [entry({ constructs: [] })] }, /"constructs" must be a non-empty array/],
    [{ version: 1, fixtures: [entry({ constructs: ['nope'] })] }, /unknown construct "nope"/],
    [{ version: 1, fixtures: [entry({ expectImportError: 'yes' })] }, /expectImportError/],
    [{ version: 1, fixtures: [entry()] }, /two\.docx is in the fixtures directory but not in the manifest/],
  ])('reports problems: %#', (manifest, pattern) => {
    expect(validateManifest(manifest, dir).join('\n')).toMatch(pattern);
  });

  it('does not require a fixtures directory to exist', () => {
    expect(validateManifest({ version: 1, fixtures: [] }, path.join(dir, 'nope'))).toEqual([]);
  });

  it('loads and sorts fixtures by id, optionally filtered', () => {
    write({ version: 1, fixtures: [entry({ id: 'two', filename: 'two.docx' }), entry()] });
    expect(loadManifest(dir).fixtures).toHaveLength(2);
    expect(discoverFixtures(dir).map((f) => f.id)).toEqual(['one', 'two']);
    expect(discoverFixtures(dir, []).map((f) => f.id)).toEqual(['one', 'two']);
    expect(discoverFixtures(dir, ['two']).map((f) => f.id)).toEqual(['two']);
    expect(() => discoverFixtures(dir, ['three'])).toThrow(/unknown fixture id\(s\): three/);
  });

  it('throws a ManifestError listing every problem', () => {
    write({ version: 1, fixtures: [entry({ title: '' }), entry({ id: 'two', filename: 'gone.docx' })] });
    try {
      loadManifest(dir);
      throw new Error('should have thrown');
    } catch (e) {
      expect(e).toBeInstanceOf(ManifestError);
      expect((e as ManifestError).problems.length).toBeGreaterThanOrEqual(2);
      expect((e as Error).message).toContain('Invalid DOCX fixture manifest');
    }
  });

  it('throws a ManifestError when the manifest is missing or not JSON', () => {
    expect(() => loadManifest(dir)).toThrow(ManifestError);
    mkdirSync(path.join(dir, 'sub'));
    writeFileSync(path.join(dir, 'sub', 'manifest.json'), '{nope');
    expect(() => loadManifest(path.join(dir, 'sub'))).toThrow(/cannot read/);
  });
});
