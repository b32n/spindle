import * as Y from 'yjs';
import {
  createOpaqueStore, DEFAULT_OPAQUE_LIMITS, decodeOpaqueEntry, droppedLossItem, encodeOpaqueEntry, hashBytes,
  isOpaqueHash, keptLossItem, readOpaqueParts, writeOpaqueParts, type OpaqueStore, type OpaqueYMap,
} from './index';

const bytes = (...n: number[]) => new Uint8Array(n);
const CT = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';

describe('hashBytes', () => {
  it('is the SHA-256 hex digest (known vectors)', async () => {
    expect(await hashBytes(new Uint8Array())).toBe('e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855');
    expect(await hashBytes(new TextEncoder().encode('abc'))).toBe('ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad');
  });
  it('is stable and content-only', async () => {
    expect(await hashBytes(bytes(1, 2, 3))).toBe(await hashBytes(bytes(1, 2, 3)));
    expect(await hashBytes(bytes(1, 2, 3))).not.toBe(await hashBytes(bytes(1, 2, 4)));
  });
  it('isOpaqueHash accepts only 64 lower-case hex chars', async () => {
    expect(isOpaqueHash(await hashBytes(bytes(1)))).toBe(true);
    for (const bad of ['', 'ABC', 'g'.repeat(64), 'a'.repeat(63), 5, null, undefined]) expect(isOpaqueHash(bad)).toBe(false);
  });
});

describe('createOpaqueStore', () => {
  it('stores a part once and dedupes by content, across names and roles', async () => {
    const s = createOpaqueStore();
    const a = await s.put({ bytes: bytes(1, 2, 3), contentType: CT, role: 'chart', partName: 'xl/charts/chart1.xml' });
    const b = await s.put({ bytes: bytes(1, 2, 3), contentType: CT, role: 'chart', partName: 'xl/charts/chart2.xml' });
    const c = await s.put({ bytes: bytes(9), contentType: 'image/png', role: 'media' });
    expect(a).toMatchObject({ ok: true, deduped: false, ref: { role: 'chart', partName: 'xl/charts/chart1.xml' } });
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
    const r = await s.put({ bytes: new Uint8Array(5), contentType: CT, role: 'media', partName: 'ppt/media/big.mp4' });
    expect(r).toEqual({ ok: false, reason: 'too-large', size: 5, role: 'media', partName: 'ppt/media/big.mp4' });
    expect(s.size).toBe(0);
    expect(s.dropped).toEqual([{ reason: 'too-large', size: 5, role: 'media', partName: 'ppt/media/big.mp4' }]);
    expect((await s.put({ bytes: new Uint8Array(4), contentType: CT, role: 'media' })).ok).toBe(true); // exactly at the cap is fine
  });

  it('refuses a part that would push the store over the total cap, but still dedupes an existing one', async () => {
    const s = createOpaqueStore({ maxPartBytes: 10, maxTotalBytes: 6 });
    expect((await s.put({ bytes: bytes(1, 2, 3, 4), contentType: CT, role: 'a' })).ok).toBe(true);
    const full = await s.put({ bytes: bytes(5, 6, 7), contentType: CT, role: 'b' });
    expect(full).toMatchObject({ ok: false, reason: 'store-full', size: 3 });
    expect(await s.put({ bytes: bytes(1, 2, 3, 4), contentType: CT, role: 'a' })).toMatchObject({ ok: true, deduped: true });
    expect(s.totalBytes).toBe(4);
    expect(s.dropped).toHaveLength(1);
  });

  it('uses the documented defaults and accepts partial overrides', () => {
    expect(createOpaqueStore().limits).toEqual(DEFAULT_OPAQUE_LIMITS);
    expect(createOpaqueStore({ maxPartBytes: 1 }).limits).toEqual({ ...DEFAULT_OPAQUE_LIMITS, maxPartBytes: 1 });
  });

  it('putVerified honours the caps and is idempotent', async () => {
    const s = createOpaqueStore({ maxPartBytes: 3, maxTotalBytes: 4 });
    const p = (n: number, len: number) => ({ hash: String(n).padStart(64, '0'), contentType: CT, bytes: new Uint8Array(len) });
    expect(s.putVerified(p(1, 3))).toBe(true);
    expect(s.putVerified(p(1, 3))).toBe(true);
    expect(s.putVerified(p(2, 4))).toBe(false); // too large
    expect(s.putVerified(p(3, 2))).toBe(false); // over the total
    expect(s.size).toBe(1);
    expect(s.totalBytes).toBe(3);
  });
});

describe('entry encoding', () => {
  it('round-trips content type (including non-ASCII, empty) and bytes', () => {
    for (const contentType of [CT, '', 'é/ü+日本']) {
      const part = { hash: 'h', contentType, bytes: bytes(0, 255, 7) };
      expect(decodeOpaqueEntry('h', encodeOpaqueEntry(part))).toEqual(part);
    }
    expect(decodeOpaqueEntry('h', encodeOpaqueEntry({ hash: 'h', contentType: CT, bytes: new Uint8Array() }))!.bytes).toHaveLength(0);
  });
  it('rejects malformed values', () => {
    expect(decodeOpaqueEntry('h', 'str')).toBeUndefined();
    expect(decodeOpaqueEntry('h', bytes(0))).toBeUndefined();
    expect(decodeOpaqueEntry('h', bytes(0, 5, 65))).toBeUndefined(); // declared length beyond the value
    expect(decodeOpaqueEntry('h', bytes(0, 1, 0xff))).toBeUndefined(); // invalid UTF-8 content type
  });
  it('refuses a content type that cannot be length-prefixed', () => {
    expect(() => encodeOpaqueEntry({ hash: 'h', contentType: 'x'.repeat(0x10000), bytes: bytes(1) })).toThrow(RangeError);
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
    expect(await readOpaqueParts(peer.getMap<Uint8Array>('opaqueParts') as unknown as OpaqueYMap, out)).toEqual({ loaded: 2, rejected: [] });
    expect(out.list().map((p) => [...p.bytes]).sort()).toEqual([[1, 2], [3]]);
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
    expect(sa.list().map((p) => p.hash).sort()).toEqual(sb.list().map((p) => p.hash).sort());
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
    const map: OpaqueYMap = { has: (k) => entries.has(k), get: (k) => entries.get(k), set: (k, v) => entries.set(k, v), forEach: (cb) => entries.forEach(cb) };
    const out = createOpaqueStore();
    const res = await readOpaqueParts(map, out);
    expect(res.loaded).toBe(1);
    expect(res.rejected.sort()).toEqual(['a'.repeat(64), await hashBytes(bytes(9)), 'not-a-hash'].sort());
    const tiny = createOpaqueStore({ maxPartBytes: 2 });
    expect(await readOpaqueParts(map, tiny)).toMatchObject({ loaded: 0 });
    expect(tiny.size).toBe(0);
  });
});

describe('loss items', () => {
  it('keptLossItem marks state kept, with a default and an overridable effect', () => {
    expect(keptLossItem({ id: 'pptx.charts', label: 'chart(s)', count: 2 })).toEqual({
      id: 'pptx.charts', label: 'chart(s)', count: 2, effect: 'kept: stored unchanged and saved back, not editable', state: 'kept',
    });
    expect(keptLossItem({ id: 'a', label: 'a', count: 1, where: ['S1'], effect: 'kept: x' })).toMatchObject({ where: ['S1'], effect: 'kept: x', state: 'kept' });
  });
  it('droppedLossItem reports refused parts as lost: too large (never silent), capping where at 5', () => {
    const d = (n: number) => ({ reason: 'too-large' as const, size: 9, role: 'media', partName: `m${n}` });
    const item = droppedLossItem('pptx.media', 'audio/video', [1, 2, 3, 4, 5, 6, 7].map(d));
    expect(item).toMatchObject({ id: 'pptx.media', count: 7, state: 'lost', effect: 'lost: dropped, too large to keep', where: ['m1', 'm2', 'm3', 'm4', 'm5'] });
  });
  it('droppedLossItem distinguishes a full store and omits where when parts are unnamed', () => {
    const item = droppedLossItem('x', 'x', [{ reason: 'store-full', size: 1, role: 'r' }]);
    expect(item.effect).toContain('over the size limit');
    expect('where' in item).toBe(false);
  });
});
