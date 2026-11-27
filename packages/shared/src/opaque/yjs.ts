import { hashBytes, isOpaqueHash } from './hash';
import type { OpaquePart, OpaqueStore } from './types';

/**
 * The slice of `Y.Map<Uint8Array>` the helper needs, so `shared` does not depend on Yjs.
 * Each core keeps one such map (`opaqueParts`, hash → encoded part) in its collab document.
 * Keying by hash means a part is carried once per document however many elements refer to it,
 * and two peers adding the same part converge on one entry.
 */
export interface OpaqueYMap {
  has(key: string): boolean;
  get(key: string): Uint8Array | undefined;
  set(key: string, value: Uint8Array): unknown;
  forEach(callback: (value: Uint8Array, key: string) => void): void;
}

const enc = new TextEncoder();
const dec = new TextDecoder('utf-8', { fatal: true });

/** Value layout: u16 big-endian content-type length, the UTF-8 content type, then the part bytes. */
export function encodeOpaqueEntry(part: OpaquePart): Uint8Array {
  const ct = enc.encode(part.contentType);
  if (ct.byteLength > 0xffff) throw new RangeError('content type too long');
  const out = new Uint8Array(2 + ct.byteLength + part.bytes.byteLength);
  out[0] = ct.byteLength >> 8;
  out[1] = ct.byteLength & 0xff;
  out.set(ct, 2);
  out.set(part.bytes, 2 + ct.byteLength);
  return out;
}

/** Inverse of `encodeOpaqueEntry`; `undefined` for a malformed value. The hash is not checked here (see `readOpaqueParts`). */
export function decodeOpaqueEntry(hash: string, value: unknown): OpaquePart | undefined {
  if (!(value instanceof Uint8Array) || value.byteLength < 2) return undefined;
  const ctLen = (value[0] << 8) | value[1];
  if (value.byteLength < 2 + ctLen) return undefined;
  try {
    return { hash, contentType: dec.decode(value.subarray(2, 2 + ctLen)), bytes: value.slice(2 + ctLen) };
  } catch {
    return undefined;
  }
}

/** Add every part of the store that the map lacks (call inside the caller's Y transaction). Returns how many were written. */
export function writeOpaqueParts(store: OpaqueStore, map: OpaqueYMap): number {
  let written = 0;
  for (const part of store.list()) {
    if (map.has(part.hash)) continue;
    map.set(part.hash, encodeOpaqueEntry(part));
    written++;
  }
  return written;
}

export interface OpaqueReadResult {
  /** Entries whose key, encoding and SHA-256 are valid, and that fit the store's caps. */
  loaded: number;
  /** Keys of entries that were malformed, whose bytes do not hash to the key, or that did not fit. They are skipped, never thrown on. */
  rejected: string[];
}

/** Load every valid part of the map into the store (a remote peer's data is untrusted: each entry is re-hashed). */
export async function readOpaqueParts(map: OpaqueYMap, store: OpaqueStore): Promise<OpaqueReadResult> {
  const entries: Array<[string, Uint8Array]> = [];
  map.forEach((v, k) => entries.push([k, v]));
  const result: OpaqueReadResult = { loaded: 0, rejected: [] };
  for (const [key, value] of entries) {
    const part = isOpaqueHash(key) ? decodeOpaqueEntry(key, value) : undefined;
    if (!part || (await hashBytes(part.bytes)) !== key || !store.putVerified(part)) result.rejected.push(key);
    else result.loaded++;
  }
  return result;
}
