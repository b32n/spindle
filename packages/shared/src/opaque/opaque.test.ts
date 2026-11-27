import * as Y from 'yjs';
import {
  createOpaqueStore,
  DEFAULT_OPAQUE_LIMITS,
  decodeOpaqueEntry,
  droppedLossItem,
  encodeOpaqueEntry,
  hashBytes,
  isOpaqueHash,
  keptLossItem,
  readOpaqueParts,
  writeOpaqueParts,
  type OpaqueStore,
  type OpaqueYMap,
} from './index';
import { opaqueInternals } from './store';

const bytes = (...n: number[]) => new Uint8Array(n);
const CT = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';

describe('hashBytes', () => {
  it('is the SHA-256 hex digest (known vectors)', async () => {
    expect(await hashBytes(new Uint8Array())).toBe(
      'e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855'
    );
    expect(await hashBytes(new TextEncoder().encode('abc'))).toBe(
      'ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad'
    );
  });
  it('is stable and content-only', async () => {
    expect(await hashBytes(bytes(1, 2, 3))).toBe(await hashBytes(bytes(1, 2, 3)));
    expect(await hashBytes(bytes(1, 2, 3))).not.toBe(await hashBytes(bytes(1, 2, 4)));
  });
  it('isOpaqueHash accepts only 64 lower-case hex chars', async () => {
    expect(isOpaqueHash(await hashBytes(bytes(1)))).toBe(true);
    for (const bad of ['', 'ABC', 'g'.repeat(64), 'a'.repeat(63), 5, null, undefined])
      expect(isOpaqueHash(bad)).toBe(false);
  });
});

describe('without SubtleCrypto (non-secure context)', () => {
  const real = Object.getOwnPropertyDescriptor(globalThis, 'crypto')!;
  const setCrypto = (value: unknown) =>
    Object.defineProperty(globalThis, 'crypto', { value, configurable: true, writable: true });
  afterEach(() => Object.defineProperty(globalThis, 'crypto', real));

  it.each([
    ['no crypto object', undefined],
    ['crypto without subtle', {}],
  ])('hashBytes and put fail with a clear error (%s)', async (_n, value) => {
    setCrypto(value);
    await expect(hashBytes(bytes(1))).rejects.toThrow(/SubtleCrypto unavailable.*secure context/);
    await expect(
      createOpaqueStore().put({ bytes: bytes(1), contentType: CT, role: 'r' })
    ).rejects.toThrow('SubtleCrypto unavailable');
  });

  it('an oversized part is still refused (not thrown on) because it is never hashed', async () => {
    setCrypto(undefined);
    const s = createOpaqueStore({ maxPartBytes: 1 });
    expect(await s.put({ bytes: bytes(1, 2), contentType: CT, role: 'r' })).toMatchObject({
      ok: false,
      reason: 'too-large',
    });
  });
});

describe('createOpaqueStore', () => {
  it('stores a part once and dedupes by content, across names and roles', async () => {
    const s = createOpaqueStore();
    const a = await s.put({
      bytes: bytes(1, 2, 3),
      contentType: CT,
      role: 'chart',
      partName: 'xl/charts/chart1.xml',
    });
    const b = await s.put({
      bytes: bytes(1, 2, 3),
      contentType: CT,
      role: 'chart',
      partName: 'xl/charts/chart2.xml',
    });
    const c = await s.put({ bytes: bytes(9), contentType: 'image/png', role: 'media' });
    expect(a).toMatchObject({
      ok: true,
      deduped: false,
      ref: { role: 'chart', partName: 'xl/charts/chart1.xml' },
    });
    expect(b).toMatchObject({ ok: true, deduped: true, ref: { partName: 'xl/charts/chart2.xml' } });
    expect(a.ok && b.ok && a.ref.hash === b.ref.hash).toBe(true);
    expect(c.ok && 'partName' in c.ref).toBe(false);
    expect(s.size).toBe(2);
    expect(s.totalBytes).toBe(4);
    expect(s.list().map((p) => p.contentType)).toEqual([CT, 'image/png']);
    expect(s.has((a as { ref: { hash: string } }).ref.hash)).toBe(true);
    expect(s.get('nope')).toBeUndefined();
  });

  it('copies bytes in, so later mutation of the caller buffer cannot change the stored part', async () => {
    const s = createOpaqueStore();
    const src = bytes(1, 2, 3);
    const r = await s.put({ bytes: src, contentType: CT, role: 'x' });
    src[0] = 99;
    expect(s.get((r as { ref: { hash: string } }).ref.hash)!.bytes).toEqual(bytes(1, 2, 3));
  });

  it('refuses (never truncates or throws) a part over the per-part cap, without hashing it, and records it', async () => {
    const s = createOpaqueStore({ maxPartBytes: 4 });
    const r = await s.put({
      bytes: new Uint8Array(5),
      contentType: CT,
      role: 'media',
      partName: 'ppt/media/big.mp4',
    });
    expect(r).toEqual({
      ok: false,
      reason: 'too-large',
      size: 5,
      role: 'media',
      partName: 'ppt/media/big.mp4',
    });
    expect(s.size).toBe(0);
    expect(s.dropped).toEqual([
      { reason: 'too-large', size: 5, role: 'media', partName: 'ppt/media/big.mp4' },
    ]);
    expect((await s.put({ bytes: new Uint8Array(4), contentType: CT, role: 'media' })).ok).toBe(
      true
    ); // exactly at the cap is fine
  });

  it('refuses a part that would push the store over the total cap, but still dedupes an existing one', async () => {
    const s = createOpaqueStore({ maxPartBytes: 10, maxTotalBytes: 6 });
    expect((await s.put({ bytes: bytes(1, 2, 3, 4), contentType: CT, role: 'a' })).ok).toBe(true);
    const full = await s.put({ bytes: bytes(5, 6, 7), contentType: CT, role: 'b' });
    expect(full).toMatchObject({ ok: false, reason: 'store-full', size: 3 });
    expect(await s.put({ bytes: bytes(1, 2, 3, 4), contentType: CT, role: 'a' })).toMatchObject({
      ok: true,
      deduped: true,
    });
    expect(s.totalBytes).toBe(4);
    expect(s.dropped).toHaveLength(1);
  });

  it('uses the documented defaults and accepts partial overrides', () => {
    expect(createOpaqueStore().limits).toEqual(DEFAULT_OPAQUE_LIMITS);
    expect(createOpaqueStore({ maxPartBytes: 1 }).limits).toEqual({
      ...DEFAULT_OPAQUE_LIMITS,
      maxPartBytes: 1,
    });
  });

  it('get and list return copies: mutating them cannot corrupt the stored part', async () => {
    const s = createOpaqueStore();
    const r = await s.put({ bytes: bytes(1, 2, 3), contentType: CT, role: 'x' });
    const hash = (r as { ref: { hash: string } }).ref.hash;
    const got = s.get(hash)!;
    got.bytes[0] = 99;
    got.contentType = 'changed';
    s.list()[0].bytes[1] = 99;
    expect(s.get(hash)).toEqual({ hash, contentType: CT, bytes: bytes(1, 2, 3) });
    expect(await hashBytes(s.get(hash)!.bytes)).toBe(hash);
  });

  it('reports a content-type conflict when equal bytes arrive under another type, keeping the first type', async () => {
    const s = createOpaqueStore();
    const a = await s.put({ bytes: bytes(4, 5), contentType: CT, role: 'chart' });
    const same = await s.put({ bytes: bytes(4, 5), contentType: CT, role: 'chart' });
    const other = await s.put({ bytes: bytes(4, 5), contentType: 'image/png', role: 'media' });
    expect(a).toMatchObject({ deduped: false, contentTypeConflict: false });
    expect(same).toMatchObject({ deduped: true, contentTypeConflict: false });
    expect(other).toMatchObject({ ok: true, deduped: true, contentTypeConflict: true });
    expect(s.list()).toHaveLength(1);
    expect(s.list()[0].contentType).toBe(CT);
  });

  it('does not expose a way to insert an unverified part', () => {
    expect('putVerified' in createOpaqueStore()).toBe(false);
  });

  it('internal accessor refuses foreign stores, applies caps, is idempotent and does not copy on read', () => {
    const foreign = { ...createOpaqueStore() };
    expect(() => opaqueInternals(foreign)).toThrow(TypeError);
    const s = createOpaqueStore({ maxPartBytes: 3, maxTotalBytes: 4 });
    const i = opaqueInternals(s);
    const p = (n: number, len: number) => ({
      hash: String(n).padStart(64, '0'),
      contentType: CT,
      bytes: new Uint8Array(len),
    });
    expect(i.insertVerified(p(1, 3))).toBe(true);
    expect(i.insertVerified(p(1, 3))).toBe(true);
    expect(i.insertVerified(p(2, 4))).toBe(false); // too large
    expect(i.insertVerified(p(3, 2))).toBe(false); // over the total
    expect(i.refusal(4)).toBe('too-large');
    expect(i.refusal(2)).toBe('store-full');
    expect(i.refusal(1)).toBeNull();
    expect(s.size).toBe(1);
    expect(s.totalBytes).toBe(3);
    expect([...i.entries()]).toHaveLength(1);
  });
});

describe('entry encoding', () => {
  it('round-trips content type (including non-ASCII, empty) and bytes', () => {
    for (const contentType of [CT, '', 'é/ü+日本']) {
      const part = { hash: 'h', contentType, bytes: bytes(0, 255, 7) };
      expect(decodeOpaqueEntry('h', encodeOpaqueEntry(part))).toEqual(part);
    }
    expect(
      decodeOpaqueEntry(
        'h',
        encodeOpaqueEntry({ hash: 'h', contentType: CT, bytes: new Uint8Array() })
      )!.bytes
    ).toHaveLength(0);
  });
  it('rejects malformed values', () => {
    expect(decodeOpaqueEntry('h', 'str')).toBeUndefined();
    expect(decodeOpaqueEntry('h', bytes(0))).toBeUndefined();
    expect(decodeOpaqueEntry('h', bytes(0, 5, 65))).toBeUndefined(); // declared length beyond the value
    expect(decodeOpaqueEntry('h', bytes(0, 1, 0xff))).toBeUndefined(); // invalid UTF-8 content type
  });
  it('refuses a content type that cannot be length-prefixed', () => {
    expect(() =>
      encodeOpaqueEntry({ hash: 'h', contentType: 'x'.repeat(0x10000), bytes: bytes(1) })
    ).toThrow(RangeError);
  });
});

async function filled(...parts: number[][]): Promise<OpaqueStore> {
  const s = createOpaqueStore();
  for (const p of parts) await s.put({ bytes: new Uint8Array(p), contentType: CT, role: 'r' });
  return s;
}

describe('Y map helper', () => {
  it('writes each part once into a Y.Map and reads it back into a fresh store', async () => {
    const doc = new Y.Doc();
    const map = doc.getMap<Uint8Array>('opaqueParts');
    const s = await filled([1, 2], [3]);
    expect(writeOpaqueParts(s, map as unknown as OpaqueYMap)).toBe(2);
    expect(writeOpaqueParts(s, map as unknown as OpaqueYMap)).toBe(0);
    expect(map.size).toBe(2);

    const peer = new Y.Doc();
    Y.applyUpdate(peer, Y.encodeStateAsUpdate(doc));
    const out = createOpaqueStore();
    expect(
      await readOpaqueParts(peer.getMap<Uint8Array>('opaqueParts') as unknown as OpaqueYMap, out)
    ).toEqual({ loaded: 2, rejected: [] });
    expect(
      out
        .list()
        .map((p) => [...p.bytes])
        .sort()
    ).toEqual([[1, 2], [3]]);
    expect(out.list()[0].contentType).toBe(CT);
  });

  it('two docs adding the same and different parts concurrently converge on the union, one entry per hash', async () => {
    const a = new Y.Doc();
    const b = new Y.Doc();
    const ma = a.getMap<Uint8Array>('opaqueParts');
    const mb = b.getMap<Uint8Array>('opaqueParts');
    writeOpaqueParts(await filled([1], [2]), ma as unknown as OpaqueYMap);
    writeOpaqueParts(await filled([2], [3]), mb as unknown as OpaqueYMap);
    Y.applyUpdate(a, Y.encodeStateAsUpdate(b));
    Y.applyUpdate(b, Y.encodeStateAsUpdate(a));
    expect([...ma.keys()].sort()).toEqual([...mb.keys()].sort());
    expect(ma.size).toBe(3);
    const sa = createOpaqueStore();
    const sb = createOpaqueStore();
    await readOpaqueParts(ma as unknown as OpaqueYMap, sa);
    await readOpaqueParts(mb as unknown as OpaqueYMap, sb);
    expect(
      sa
        .list()
        .map((p) => p.hash)
        .sort()
    ).toEqual(
      sb
        .list()
        .map((p) => p.hash)
        .sort()
    );
  });

  it('an update carries a shared part once', async () => {
    const big = new Uint8Array(50_000).fill(7);
    const doc = new Y.Doc();
    const map = doc.getMap<Uint8Array>('opaqueParts');
    const s = createOpaqueStore();
    for (let i = 0; i < 5; i++) await s.put({ bytes: big, contentType: CT, role: `r${i}` });
    writeOpaqueParts(s, map as unknown as OpaqueYMap);
    expect(Y.encodeStateAsUpdate(doc).byteLength).toBeLessThan(big.byteLength * 2);
  });

  it('skips untrusted entries: bad key, malformed value, wrong hash, over the caps', async () => {
    const good = await filled([1, 2, 3]);
    const goodPart = good.list()[0];
    const entries = new Map<string, Uint8Array>([
      [goodPart.hash, encodeOpaqueEntry(goodPart)],
      ['not-a-hash', encodeOpaqueEntry(goodPart)],
      ['a'.repeat(64), bytes(0)],
      [await hashBytes(bytes(9)), encodeOpaqueEntry({ ...goodPart, bytes: bytes(8) })], // bytes do not hash to the key
    ]);
    const map: OpaqueYMap = {
      has: (k) => entries.has(k),
      get: (k) => entries.get(k),
      set: (k, v) => entries.set(k, v),
      forEach: (cb) => entries.forEach(cb),
    };
    const out = createOpaqueStore();
    const res = await readOpaqueParts(map, out);
    expect(res.loaded).toBe(1);
    expect(res.rejected.sort()).toEqual(
      ['a'.repeat(64), await hashBytes(bytes(9)), 'not-a-hash'].sort()
    );
    const tiny = createOpaqueStore({ maxPartBytes: 2 });
    expect(await readOpaqueParts(map, tiny)).toMatchObject({ loaded: 0 });
    expect(tiny.size).toBe(0);
  });
});

describe('readOpaqueParts size pre-checks', () => {
  const mapOf = (entries: Map<string, Uint8Array>): OpaqueYMap => ({
    has: (k) => entries.has(k),
    get: (k) => entries.get(k),
    set: (k, v) => entries.set(k, v),
    forEach: (cb) => entries.forEach(cb),
  });
  const entryFor = async (data: Uint8Array, contentType = CT) =>
    [await hashBytes(data), encodeOpaqueEntry({ hash: '', contentType, bytes: data })] as const;

  it('rejects an oversized peer entry without copying or hashing it', async () => {
    const small = await entryFor(bytes(1));
    const bigKey = 'b'.repeat(64);
    const big = encodeOpaqueEntry({ hash: bigKey, contentType: CT, bytes: new Uint8Array(100) });
    const sliceSpy = jest.spyOn(big, 'slice');
    const digest = jest.spyOn(globalThis.crypto.subtle, 'digest');
    try {
      const store = createOpaqueStore({ maxPartBytes: 10 });
      const res = await readOpaqueParts(mapOf(new Map([[bigKey, big], small])), store);
      expect(res).toEqual({ loaded: 1, rejected: [bigKey] });
      expect(digest).toHaveBeenCalledTimes(1); // the small entry only
      expect(sliceSpy).not.toHaveBeenCalled();
    } finally {
      digest.mockRestore();
    }
  });

  it('measures the part bytes, not the content type: a long content type with a small part is fine, exactly at the cap is fine', async () => {
    const e = await entryFor(new Uint8Array(10), 'x/'.padEnd(5000, 'y'));
    expect(
      await readOpaqueParts(mapOf(new Map([e])), createOpaqueStore({ maxPartBytes: 10 }))
    ).toEqual({ loaded: 1, rejected: [] });
    expect(
      await readOpaqueParts(mapOf(new Map([e])), createOpaqueStore({ maxPartBytes: 9 }))
    ).toMatchObject({ loaded: 0 });
    expect(decodeOpaqueEntry('h', e[1], (n) => n <= 10)).toBeDefined();
    expect(decodeOpaqueEntry('h', e[1], (n) => n <= 9)).toBeUndefined();
  });

  it('rejects an entry that would overflow the store before hashing, but accepts one the store already holds', async () => {
    const a = await entryFor(bytes(1, 2, 3, 4));
    const b = await entryFor(bytes(5, 6, 7));
    const sliceSpy = jest.spyOn(b[1], 'slice');
    const store = createOpaqueStore({ maxTotalBytes: 5 });
    const digest = jest.spyOn(globalThis.crypto.subtle, 'digest');
    try {
      expect(await readOpaqueParts(mapOf(new Map([a, b])), store)).toEqual({
        loaded: 1,
        rejected: [b[0]],
      });
      expect(digest).toHaveBeenCalledTimes(1);
      expect(sliceSpy).not.toHaveBeenCalled();
      expect(await readOpaqueParts(mapOf(new Map([a])), store)).toEqual({
        loaded: 1,
        rejected: [],
      }); // already stored
    } finally {
      digest.mockRestore();
    }
  });

  it('propagates a missing SubtleCrypto as an error rather than rejecting the peer entries', async () => {
    const real = Object.getOwnPropertyDescriptor(globalThis, 'crypto')!;
    const e = await entryFor(bytes(1));
    Object.defineProperty(globalThis, 'crypto', {
      value: undefined,
      configurable: true,
      writable: true,
    });
    try {
      await expect(readOpaqueParts(mapOf(new Map([e])), createOpaqueStore())).rejects.toThrow(
        'SubtleCrypto unavailable'
      );
    } finally {
      Object.defineProperty(globalThis, 'crypto', real);
    }
  });

  it('refuses a store not made by createOpaqueStore', async () => {
    await expect(readOpaqueParts(mapOf(new Map()), { ...createOpaqueStore() })).rejects.toThrow(
      TypeError
    );
    expect(() => writeOpaqueParts({ ...createOpaqueStore() }, mapOf(new Map()))).toThrow(TypeError);
  });
});

describe('loss items', () => {
  it('keptLossItem marks state kept, with a default and an overridable effect', () => {
    expect(keptLossItem({ id: 'pptx.charts', label: 'chart(s)', count: 2 })).toEqual({
      id: 'pptx.charts',
      label: 'chart(s)',
      count: 2,
      effect: 'kept: stored unchanged and saved back, not editable',
      state: 'kept',
    });
    expect(
      keptLossItem({ id: 'a', label: 'a', count: 1, where: ['S1'], effect: 'kept: x' })
    ).toMatchObject({ where: ['S1'], effect: 'kept: x', state: 'kept' });
  });
  it('droppedLossItem reports refused parts as lost: too large (never silent), capping where at 5', () => {
    const d = (n: number) => ({
      reason: 'too-large' as const,
      size: 9,
      role: 'media',
      partName: `m${n}`,
    });
    const item = droppedLossItem('pptx.media', 'audio/video', [1, 2, 3, 4, 5, 6, 7].map(d));
    expect(item).toMatchObject({
      id: 'pptx.media',
      count: 7,
      state: 'lost',
      effect: 'lost: dropped, too large to keep',
      where: ['m1', 'm2', 'm3', 'm4', 'm5'],
    });
  });
  it('droppedLossItem distinguishes a full store and omits where when parts are unnamed', () => {
    const item = droppedLossItem('x', 'x', [{ reason: 'store-full', size: 1, role: 'r' }]);
    expect(item.effect).toContain('over the size limit');
    expect('where' in item).toBe(false);
  });
});
