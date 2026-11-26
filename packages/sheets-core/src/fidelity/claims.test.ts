import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import JSZip from 'jszip';
import {
  detectConstructs,
  unverifiedClaims,
  verifyClaims,
} from '../../../../scripts/xlsx-fidelity/claims';
import type { OoxmlParts } from '../../../../scripts/xlsx-fidelity/ooxml-package';
import type { Construct, Manifest } from '../../../../scripts/xlsx-fidelity/types';

const registry: Array<Construct<OoxmlParts>> = [
  { id: 'has.a', label: 'A', count: (p) => (p.names.includes('a.xml') ? 2 : 0) },
  { id: 'has.b', label: 'B', count: (p) => (p.names.includes('b.xml') ? 1 : 0) },
];
const parts = (names: string[]): OoxmlParts => ({
  kind: 'zip',
  raw: new Uint8Array(),
  names,
  xml: {},
  sizes: {},
  oversized: [],
});

describe('claims', () => {
  it('reports claimed constructs whose probe finds nothing, including unknown ids', () => {
    expect(unverifiedClaims(['has.a', 'has.b', 'unknown'], registry, parts(['a.xml']))).toEqual([
      'has.b',
      'unknown',
    ]);
    expect(unverifiedClaims([], registry, parts([]))).toEqual([]);
  });

  it('lists the constructs a package actually contains', () => {
    expect(detectConstructs(registry, parts(['a.xml', 'b.xml']))).toEqual(['has.a', 'has.b']);
    expect(detectConstructs(registry, parts([]))).toEqual([]);
  });

  describe('verifyClaims', () => {
    let dir: string;
    beforeEach(() => (dir = mkdtempSync(path.join(tmpdir(), 'xlsx-claims-'))));
    afterEach(() => rmSync(dir, { recursive: true, force: true }));
    const entry = (id: string, constructs: string[], extra: Record<string, unknown> = {}) => ({
      id,
      filename: `${id}.xlsx`,
      title: id,
      source: { tool: 't', os: 'o' },
      license: 'l',
      notes: 'n',
      constructs,
      ...extra,
    });

    it('checks committed files, skips run-time ones and names every false claim', async () => {
      const z = new JSZip();
      z.file('a.xml', '<a/>');
      writeFileSync(path.join(dir, 'ok.xlsx'), await z.generateAsync({ type: 'uint8array' }));
      writeFileSync(path.join(dir, 'bad.xlsx'), await z.generateAsync({ type: 'uint8array' }));
      const manifest = {
        version: 1,
        fixtures: [
          entry('ok', ['has.a']),
          entry('bad', ['has.a', 'has.b']),
          entry('later', ['has.b'], { generatedAtRuntime: true }),
        ],
      } as unknown as Manifest;
      expect(await verifyClaims(dir, manifest, registry)).toEqual([
        'bad: claims "has.b" but the file does not contain it',
      ]);
    });
  });
});
