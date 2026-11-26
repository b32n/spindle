import { mkdtempSync, readFileSync, rmSync, writeFileSync, existsSync, mkdirSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import {
  compareSnapshot,
  findOrphanGoldens,
  removeGoldens,
  snapshotPath,
  stableStringify,
} from '../../../../../scripts/fidelity-core/snapshot';

let dir: string;
beforeEach(() => {
  dir = mkdtempSync(path.join(tmpdir(), 'fidelity-snapshot-'));
});
afterEach(() => rmSync(dir, { recursive: true, force: true }));

describe('stableStringify', () => {
  it('sorts keys recursively, drops undefined, ends with a newline', () => {
    expect(stableStringify({ b: 1, a: { d: [3, { z: 1, y: undefined, x: null }], c: 'x' } })).toBe(
      '{\n  "a": {\n    "c": "x",\n    "d": [\n      3,\n      {\n        "x": null,\n        "z": 1\n      }\n    ]\n  },\n  "b": 1\n}\n'
    );
  });

  it('turns a bare undefined into null and keeps primitives', () => {
    expect(stableStringify(undefined)).toBe('null\n');
    expect(stableStringify('s')).toBe('"s"\n');
    expect(stableStringify([undefined])).toBe('[\n  null\n]\n');
  });

  it('is independent of key insertion order', () => {
    expect(stableStringify({ a: 1, b: 2 })).toBe(stableStringify({ b: 2, a: 1 }));
  });
});

describe('snapshotPath', () => {
  it('names a golden <id>.<kind>.json', () => {
    expect(snapshotPath('/g', 'cell-types', 'values')).toBe(
      path.join('/g', 'cell-types.values.json')
    );
  });
});

describe('compareSnapshot', () => {
  it('reports missing, then created in update mode, then match, mismatch, updated', () => {
    const file = path.join(dir, 'nested', 'a.model.json');
    expect(compareSnapshot(file, { a: 1 }, false)).toBe('missing');
    expect(existsSync(file)).toBe(false);
    expect(compareSnapshot(file, { a: 1 }, true)).toBe('created');
    expect(readFileSync(file, 'utf8')).toBe('{\n  "a": 1\n}\n');
    expect(compareSnapshot(file, { a: 1 }, false)).toBe('match');
    expect(compareSnapshot(file, { a: 2 }, false)).toBe('mismatch');
    expect(readFileSync(file, 'utf8')).toContain('1');
    expect(compareSnapshot(file, { a: 2 }, true)).toBe('updated');
    expect(compareSnapshot(file, { a: 2 }, true)).toBe('match');
  });
});

describe('orphan goldens', () => {
  it('finds goldens of the given kinds whose fixture is unknown, ignoring other files', () => {
    expect(findOrphanGoldens(path.join(dir, 'missing'), ['a'], ['model'])).toEqual([]);
    mkdirSync(path.join(dir, 'g'));
    for (const f of [
      'a.model.json',
      'a.loss.json',
      'gone.model.json',
      'gone.values.json',
      'gone.other.json',
      'README.md',
      'x.modelXjson',
    ])
      writeFileSync(path.join(dir, 'g', f), '{}');
    expect(findOrphanGoldens(path.join(dir, 'g'), ['a'], ['model', 'loss', 'values'])).toEqual([
      'gone.model.json',
      'gone.values.json',
    ]);
    expect(
      findOrphanGoldens(path.join(dir, 'g'), new Set(['a', 'gone']), ['model', 'loss', 'values'])
    ).toEqual([]);
  });

  it('escapes kind names that contain regular-expression characters', () => {
    mkdirSync(path.join(dir, 'g'));
    writeFileSync(path.join(dir, 'g', 'a.m+.json'), '{}');
    writeFileSync(path.join(dir, 'g', 'a.mm.json'), '{}');
    expect(findOrphanGoldens(path.join(dir, 'g'), [], ['m+'])).toEqual(['a.m+.json']);
  });

  it('removes the given goldens', () => {
    writeFileSync(path.join(dir, 'x.model.json'), '{}');
    removeGoldens(dir, ['x.model.json']);
    expect(existsSync(path.join(dir, 'x.model.json'))).toBe(false);
  });
});
