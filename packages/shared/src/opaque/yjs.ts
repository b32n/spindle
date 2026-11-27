import { hashBytes, isOpaqueHash } from './hash';
import { opaqueInternals } from './store';
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

/**
 * Inverse of `encodeOpaqueEntry`; `undefined` for a malformed value, or one whose part length `accepts` refuses
 * (judged from the header before anything is copied). The hash is not checked here (see `readOpaqueParts`).
 */
export function decodeOpaqueEntry(
  hash: string,
  value: unknown,
  accepts: (partBytes: number) => boolean = () => true
): OpaquePart | undefined {
  if (!(value instanceof Uint8Array) || value.byteLength < 2) return undefined;
  const ctLen = (value[0] << 8) | value[1];
  if (value.byteLength < 2 + ctLen || !accepts(value.byteLength - 2 - ctLen)) return undefined;
  try {
    return {
      hash,
      contentType: dec.decode(value.subarray(2, 2 + ctLen)),
      bytes: value.slice(2 + ctLen),
    };
  } catch {
    return undefined;
  }
}

/** Add every part of the store that the map lacks (call inside the caller's Y transaction). Returns how many were written. */
export function writeOpaqueParts(store: OpaqueStore, map: OpaqueYMap): number {
  let written = 0;
  for (const part of opaqueInternals(store).entries()) {
    if (map.has(part.hash)) continue;
    map.set(part.hash, encodeOpaqueEntry(part));
    written++;
  }
  return written;
}

export interface OpaqueReadResult {
  /** Entries whose key, encoding and SHA-256 are valid and that fit the store's caps, plus those under keys the store already held. */
  loaded: number;
  /** Keys of entries that were malformed, whose bytes do not hash to the key, or that did not fit. They are skipped, never thrown on. */
  rejected: string[];
}

/**
 * Load every valid part of the map into the store. A remote peer's data is untrusted: an entry is rejected from its header, before any
 * copy or hash, when its part is over the per-part cap or would not fit the store, and otherwise re-hashed against its key.
 * An entry under a key the store already holds is counted as loaded without being read.
 * Rejects (does not skip) only when SubtleCrypto is unavailable, which is the host's problem, not the peer's.
 */
export async function readOpaqueParts(
  map: OpaqueYMap,
  store: OpaqueStore
): Promise<OpaqueReadResult> {
  const trusted = opaqueInternals(store);
  const entries: Array<[string, Uint8Array]> = [];
  map.forEach((v, k) => entries.push([k, v]));
  const result: OpaqueReadResult = { loaded: 0, rejected: [] };
  for (const [key, value] of entries) {
    if (isOpaqueHash(key) && store.has(key)) {
      // Already held (and verified when it went in): the peer's copy is neither sliced nor hashed, whatever its size.
      result.loaded++;
      continue;
    }
    const part = isOpaqueHash(key)
      ? decodeOpaqueEntry(key, value, (n) => !trusted.refusal(n))
      : undefined;
    if (!part || (await hashBytes(part.bytes)) !== key || !trusted.insertVerified(part))
      result.rejected.push(key);
    else result.loaded++;
  }
  return result;
}
