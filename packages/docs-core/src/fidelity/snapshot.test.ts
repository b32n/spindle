import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { compareSnapshot, findOrphanGoldens, removeGoldens, snapshotPath, stableStringify } from '../../../../scripts/docx-fidelity/snapshot';

describe('stableStringify', () => {
  it('sorts keys recursively, pretty-prints and ends with a newline', () => {
    const out = stableStringify({ b: 1, a: { d: [{ z: 1, y: undefined, x: 2 }], c: null } });
    expect(out).toBe('{\n  "a": {\n    "c": null,\n    "d": [\n      {\n        "x": 2,\n        "z": 1\n      }\n    ]\n  },\n  "b": 1\n}\n');
  });

  it('is independent of key insertion order', () => {
    expect(stableStringify({ a: 1, b: 2 })).toBe(stableStringify({ b: 2, a: 1 }));
  });

  it('maps a bare undefined to null', () => {
    expect(stableStringify(undefined)).toBe('null\n');
  });
});

describe('golden files', () => {
  let dir: string;
  beforeEach(() => {
    dir = mkdtempSync(path.join(tmpdir(), 'docx-golden-'));
  });
  afterEach(() => rmSync(dir, { recursive: true, force: true }));

  it('names snapshots <id>.<kind>.json', () => {
    expect(snapshotPath('/g', 'abc', 'loss')).toBe(path.join('/g', 'abc.loss.json'));
  });

  it('reports missing, then mismatch, then match', () => {
    const file = path.join(dir, 'a.model.json');
    expect(compareSnapshot(file, { x: 1 }, false)).toBe('missing');
    expect(existsSync(file)).toBe(false);
    writeFileSync(file, stableStringify({ x: 2 }));
    expect(compareSnapshot(file, { x: 1 }, false)).toBe('mismatch');
    expect(readFileSync(file, 'utf8')).toBe(stableStringify({ x: 2 }));
    expect(compareSnapshot(file, { x: 2 }, false)).toBe('match');
  });

  it('creates and updates goldens only in update mode', () => {
    const file = path.join(dir, 'nested', 'b.roundtrip.json');
    expect(compareSnapshot(file, { x: 1 }, true)).toBe('created');
    expect(compareSnapshot(file, { x: 1 }, true)).toBe('match');
    expect(compareSnapshot(file, { x: 2 }, true)).toBe('updated');
    expect(readFileSync(file, 'utf8')).toBe(stableStringify({ x: 2 }));
  });

  it('finds and removes orphaned goldens', () => {
    for (const f of ['keep.model.json', 'keep.loss.json', 'gone.model.json', 'gone.roundtrip.json', 'notes.txt']) writeFileSync(path.join(dir, f), '{}\n');
    expect(findOrphanGoldens(dir, ['keep'])).toEqual(['gone.model.json', 'gone.roundtrip.json']);
    removeGoldens(dir, ['gone.model.json', 'gone.roundtrip.json']);
    expect(findOrphanGoldens(dir, ['keep'])).toEqual([]);
    expect(existsSync(path.join(dir, 'notes.txt'))).toBe(true);
    expect(findOrphanGoldens(path.join(dir, 'does-not-exist'), [])).toEqual([]);
  });
});
