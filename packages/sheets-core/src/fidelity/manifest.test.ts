import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import {
  ManifestError,
  loadManifest,
  validateManifest,
} from '../../../../scripts/xlsx-fidelity/manifest';
import type { ManifestOptions } from '../../../../scripts/xlsx-fidelity/types';

const opts: ManifestOptions = {
  extensions: ['.xlsx', '.xls'],
  constructIds: ['a.one', 'a.two'],
  label: 'Test',
};
const entry = (over: Record<string, unknown> = {}) => ({
  id: 'one',
  filename: 'one.xlsx',
  title: 'One',
  source: { tool: 'test', os: 'linux' },
  license: 'original',
  constructs: ['a.one'],
  notes: 'n',
  ...over,
});

describe('manifest', () => {
  let dir: string;
  beforeEach(() => {
    dir = mkdtempSync(path.join(tmpdir(), 'xlsx-manifest-'));
    writeFileSync(path.join(dir, 'one.xlsx'), 'x');
    writeFileSync(path.join(dir, 'two.xls'), 'x');
  });
  afterEach(() => rmSync(dir, { recursive: true, force: true }));
  const write = (m: unknown) => writeFileSync(path.join(dir, 'manifest.json'), JSON.stringify(m));
  const valid = {
    version: 1,
    fixtures: [
      entry(),
      entry({
        id: 'two',
        filename: 'two.xls',
        expectImportError: true,
        source: { tool: 't', version: '1.2', os: 'o' },
      }),
    ],
  };

  it('accepts a valid manifest, with optional version and flags', () => {
    expect(validateManifest(valid, dir, opts)).toEqual([]);
  });

  it('accepts a run-time fixture that is not on disk and names a generator', () => {
    const m = {
      version: 1,
      fixtures: [
        ...valid.fixtures,
        entry({
          id: 'big',
          filename: 'big.xlsx',
          generatedAtRuntime: true,
          generator: 'generators/x.py',
        }),
      ],
    };
    expect(validateManifest(m, dir, opts)).toEqual([]);
  });

  it('rejects a manifest that is not an object with fixtures', () => {
    expect(validateManifest(null, dir, opts)).toHaveLength(1);
    expect(validateManifest('x', dir, opts)).toHaveLength(1);
    expect(validateManifest({ fixtures: 3 }, dir, opts)).toHaveLength(1);
  });

  it.each([
    [{ version: 2, fixtures: [] }, /unsupported manifest version/],
    [{ version: 1, fixtures: [entry({ title: '' })] }, /"title" must be a non-empty string/],
    [
      { version: 1, fixtures: [entry({ license: undefined })] },
      /"license" must be a non-empty string/,
    ],
    [{ version: 1, fixtures: [entry({ id: 'Bad_Id' })] }, /kebab-case/],
    [{ version: 1, fixtures: [entry(), entry()] }, /duplicate id/],
    [{ version: 1, fixtures: [entry({ filename: '../x.xlsx' })] }, /plain name ending in/],
    [
      { version: 1, fixtures: [entry({ filename: 'one.docx' })] },
      /plain name ending in \.xlsx \/ \.xls/,
    ],
    [{ version: 1, fixtures: [entry({ filename: 'ONE.XLSX' })] }, /does not exist/],
    [{ version: 1, fixtures: [entry({ filename: 'missing.xlsx' })] }, /does not exist/],
    [{ version: 1, fixtures: [entry(), entry({ id: 'other' })] }, /duplicate filename/],
    [
      { version: 1, fixtures: [entry({ generatedAtRuntime: true, generator: 'g.py' })] },
      /must not be committed/,
    ],
    [
      { version: 1, fixtures: [entry({ filename: 'big.xlsx', generatedAtRuntime: true })] },
      /needs a "generator"/,
    ],
    [
      {
        version: 1,
        fixtures: [entry({ filename: 'big.xlsx', generatedAtRuntime: true, generator: '' })],
      },
      /needs a "generator"/,
    ],
    [{ version: 1, fixtures: [entry({ source: { tool: 'x' } })] }, /"source" needs/],
    [{ version: 1, fixtures: [entry({ source: undefined })] }, /"source" needs/],
    [
      { version: 1, fixtures: [entry({ source: { tool: 'x', os: 'y', version: '' } })] },
      /source\.version/,
    ],
    [
      { version: 1, fixtures: [entry({ source: { tool: 'x', os: 'y', version: 3 } })] },
      /source\.version/,
    ],
    [
      { version: 1, fixtures: [entry({ constructs: [] })] },
      /"constructs" must be a non-empty array/,
    ],
    [
      { version: 1, fixtures: [entry({ constructs: undefined })] },
      /"constructs" must be a non-empty array/,
    ],
    [{ version: 1, fixtures: [entry({ constructs: ['nope'] })] }, /unknown construct "nope"/],
    [
      { version: 1, fixtures: [entry({ expectImportError: 'yes' })] },
      /"expectImportError" must be a boolean/,
    ],
    [
      { version: 1, fixtures: [entry({ generatedAtRuntime: 1 })] },
      /"generatedAtRuntime" must be a boolean/,
    ],
    [{ version: 1, fixtures: [entry({ id: 5, filename: 5 })] }, /"id" must be a non-empty string/],
    [
      { version: 1, fixtures: [entry()] },
      /two\.xls is in the fixtures directory but not in the manifest/,
    ],
    [{ version: 1, fixtures: [null] }, /"id" must be a non-empty string/],
  ])('reports problems: %#', (manifest, pattern) => {
    expect(validateManifest(manifest, dir, opts).join('\n')).toMatch(pattern);
  });

  it('does not require a fixtures directory to exist', () => {
    expect(validateManifest({ version: 1, fixtures: [] }, path.join(dir, 'nope'), opts)).toEqual(
      []
    );
  });

  it('loads a valid manifest', () => {
    write(valid);
    expect(loadManifest(dir, opts).fixtures).toHaveLength(2);
  });

  it('throws a ManifestError listing every problem', () => {
    write({
      version: 1,
      fixtures: [entry({ title: '' }), entry({ id: 'two', filename: 'gone.xlsx' })],
    });
    try {
      loadManifest(dir, opts);
      throw new Error('should have thrown');
    } catch (e) {
      expect(e).toBeInstanceOf(ManifestError);
      expect((e as ManifestError).problems.length).toBeGreaterThanOrEqual(2);
      expect((e as Error).message).toContain('Invalid Test manifest');
    }
  });

  it('throws a ManifestError when the manifest is missing or not JSON', () => {
    expect(() => loadManifest(dir, opts)).toThrow(ManifestError);
    mkdirSync(path.join(dir, 'sub'));
    writeFileSync(path.join(dir, 'sub', 'manifest.json'), '{ nope');
    expect(() => loadManifest(path.join(dir, 'sub'), opts)).toThrow(/cannot read/);
    expect(new ManifestError(['x']).message).toContain('Invalid fixture manifest');
  });
});
