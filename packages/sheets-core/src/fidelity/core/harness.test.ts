import { existsSync, mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import {
  describeOutcome,
  goldensDirOf,
  importExpectationOf,
  runHarness,
  runOne,
  selectFixtures,
  withTimeout,
  type HarnessAdapter,
} from '../../../../../scripts/fidelity-core/harness';
import { snapshotPath, stableStringify } from '../../../../../scripts/fidelity-core/snapshot';
import type {
  Manifest,
  ManifestEntry,
  ManifestOptions,
} from '../../../../../scripts/fidelity-core/types';

const options: ManifestOptions = { extensions: ['.dat'], constructIds: ['k'], label: 'test' };
const entry = (id: string, extra: Partial<ManifestEntry> = {}): ManifestEntry => ({
  id,
  filename: `${id}.dat`,
  title: id,
  source: { tool: 't', os: 'linux' },
  license: 'l',
  constructs: ['k'],
  notes: 'n',
  ...extra,
});

interface R {
  text: string;
  error: string | null;
}
const adapter: HarnessAdapter<R> = {
  kinds: ['model', 'extra'],
  run: async (bytes) => {
    const text = Buffer.from(bytes).toString('utf8');
    return { text, error: text === 'boom' ? 'it broke' : null };
  },
  crashed: (message) => ({ text: '', error: `crashed: ${message}` }),
  snapshot: (r, kind) => ({ kind, text: r.text, error: r.error }),
  importError: (r) => r.error,
};

let dir: string;
const writeManifest = (fixtures: ManifestEntry[]) =>
  writeFileSync(
    path.join(dir, 'manifest.json'),
    JSON.stringify({ version: 1, fixtures } satisfies Manifest)
  );
beforeEach(() => {
  dir = mkdtempSync(path.join(tmpdir(), 'fidelity-harness-'));
});
afterEach(() => rmSync(dir, { recursive: true, force: true }));

describe('withTimeout', () => {
  it('resolves with the value and clears its timer', async () => {
    await expect(withTimeout(Promise.resolve(5), 1000, 'x')).resolves.toBe(5);
  });
  it('passes a rejection through', async () => {
    await expect(withTimeout(Promise.reject(new Error('no')), 1000, 'x')).rejects.toThrow('no');
  });
  it('rejects with a named message when the work does not finish', async () => {
    await expect(withTimeout(new Promise(() => undefined), 5, 'thing')).rejects.toThrow(
      'thing did not finish within 5 ms'
    );
  });
});

describe('selectFixtures', () => {
  const manifest: Manifest = {
    version: 1,
    fixtures: [
      entry('a'),
      entry('b'),
      entry('rt', { generatedAtRuntime: true, generator: 'g.py' }),
    ],
  };
  it('returns the committed fixtures, not the run-time ones', () => {
    expect(selectFixtures(manifest).map((f) => f.id)).toEqual(['a', 'b']);
    expect(selectFixtures(manifest, []).map((f) => f.id)).toEqual(['a', 'b']);
  });
  it('narrows to the requested ids and rejects unknown ones (run-time ids are not selectable)', () => {
    expect(selectFixtures(manifest, ['b']).map((f) => f.id)).toEqual(['b']);
    expect(() => selectFixtures(manifest, ['nope'])).toThrow(
      'unknown fixture id(s): nope (known: a, b)'
    );
    expect(() => selectFixtures(manifest, ['rt'])).toThrow('unknown fixture id(s): rt');
  });
});

describe('importExpectationOf', () => {
  it.each([
    [{}, null, 'as-expected'],
    [{}, 'err', 'unexpected-error'],
    [{ expectImportError: true }, 'err', 'as-expected'],
    [{ expectImportError: true }, null, 'unexpected-success'],
    [{ expectImportError: false }, 'err', 'unexpected-error'],
  ] as const)('%j with %j is %s', (extra, error, expected) => {
    expect(importExpectationOf(entry('x', extra), error)).toBe(expected);
  });
});

describe('runOne', () => {
  it('returns what the adapter returns', async () => {
    expect(await runOne(adapter, entry('a'), new TextEncoder().encode('hi'), 1000)).toEqual({
      text: 'hi',
      error: null,
    });
  });
  it('turns an adapter exception and a timeout into the adapter crash result', async () => {
    const throwing: HarnessAdapter<R> = {
      ...adapter,
      run: () => Promise.reject(new Error('kaput')),
    };
    expect(await runOne(throwing, entry('a'), new Uint8Array(), 1000)).toEqual({
      text: '',
      error: 'crashed: kaput',
    });
    const throwingString: HarnessAdapter<R> = { ...adapter, run: () => Promise.reject('plain') };
    expect((await runOne(throwingString, entry('a'), new Uint8Array(), 1000)).error).toBe(
      'crashed: plain'
    );
    const hanging: HarnessAdapter<R> = { ...adapter, run: () => new Promise(() => undefined) };
    expect((await runOne(hanging, entry('a'), new Uint8Array(), 5)).error).toBe(
      'crashed: fixture a did not finish within 5 ms'
    );
  });
});

describe('runHarness', () => {
  it('writes goldens in update mode, then matches them, and flags drift', async () => {
    writeManifest([entry('a'), entry('b', { expectImportError: true })]);
    writeFileSync(path.join(dir, 'a.dat'), 'alpha');
    writeFileSync(path.join(dir, 'b.dat'), 'boom');
    const base = { adapter, fixturesDir: dir, manifestOptions: options };

    const missing = await runHarness(base);
    expect(missing.ok).toBe(false);
    expect(missing.fixtures[0].status).toEqual({ model: 'missing', extra: 'missing' });
    expect(describeOutcome(missing.fixtures[0], adapter.kinds)).toEqual([
      'a: model snapshot missing',
      'a: extra snapshot missing',
    ]);

    const created = await runHarness({ ...base, update: true });
    expect(created.ok).toBe(true);
    expect(created.fixtures[0].status.model).toBe('created');
    expect(existsSync(snapshotPath(goldensDirOf(dir), 'b', 'extra'))).toBe(true);

    const matched = await runHarness(base);
    expect(matched.ok).toBe(true);
    expect(matched.fixtures.map((f) => f.importExpectation)).toEqual([
      'as-expected',
      'as-expected',
    ]);

    writeFileSync(path.join(dir, 'a.dat'), 'changed');
    const drift = await runHarness(base);
    expect(drift.ok).toBe(false);
    expect(describeOutcome(drift.fixtures[0], adapter.kinds)).toEqual([
      'a: model snapshot mismatch',
      'a: extra snapshot mismatch',
    ]);
    expect((await runHarness({ ...base, only: ['b'] })).ok).toBe(true);
  });

  it('records an unexpected import error without failing, or fails under strictImport', async () => {
    writeManifest([entry('a'), entry('b', { expectImportError: true })]);
    writeFileSync(path.join(dir, 'a.dat'), 'boom');
    writeFileSync(path.join(dir, 'b.dat'), 'fine');
    const base = { adapter, fixturesDir: dir, manifestOptions: options, update: true };
    await runHarness(base);

    const lenient = await runHarness({ ...base, update: false });
    expect(lenient.ok).toBe(true);
    expect(lenient.fixtures.map((f) => f.importExpectation)).toEqual([
      'unexpected-error',
      'unexpected-success',
    ]);
    expect(lenient.fixtures.every((f) => f.problems.length === 0)).toBe(true);

    const strict = await runHarness({ ...base, update: false, strictImport: true });
    expect(strict.ok).toBe(false);
    expect(strict.fixtures[0].problems).toEqual(['the import threw unexpectedly: it broke']);
    expect(strict.fixtures[1].problems).toEqual([
      'manifest says the import should throw, but it succeeded',
    ]);
    expect(describeOutcome(strict.fixtures[0], adapter.kinds)).toEqual([
      'a: the import threw unexpectedly: it broke',
    ]);
  });

  it('reports orphan goldens, and deletes them in update mode', async () => {
    writeManifest([entry('a')]);
    writeFileSync(path.join(dir, 'a.dat'), 'x');
    const base = { adapter, fixturesDir: dir, manifestOptions: options };
    await runHarness({ ...base, update: true });
    const goldens = goldensDirOf(dir);
    mkdirSync(goldens, { recursive: true });
    writeFileSync(path.join(goldens, 'old.model.json'), '{}');
    const found = await runHarness(base);
    expect(found.orphanGoldens).toEqual(['old.model.json']);
    expect(found.ok).toBe(false);
    const fixed = await runHarness({ ...base, update: true });
    expect(fixed).toMatchObject({ ok: true, orphanGoldens: [] });
    expect(existsSync(path.join(goldens, 'old.model.json'))).toBe(false);
    // orphans are judged against the whole manifest even for a subset run
    expect((await runHarness({ ...base, only: ['a'] })).orphanGoldens).toEqual([]);
  });

  it('turns a run that exceeds the time limit into the crash result and keeps going', async () => {
    writeManifest([entry('a')]);
    writeFileSync(path.join(dir, 'a.dat'), 'x');
    const slow: HarnessAdapter<R> = { ...adapter, run: () => new Promise(() => undefined) };
    const run = await runHarness({
      adapter: slow,
      fixturesDir: dir,
      manifestOptions: options,
      update: true,
      timeoutMs: 5,
    });
    expect(run.fixtures[0].result.error).toContain('did not finish within 5 ms');
    expect(stableStringify(run.fixtures[0].result)).toContain('crashed');
  });
});
