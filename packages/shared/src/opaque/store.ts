import { hashBytes } from './hash';
import type {
  OpaqueDropped,
  OpaqueLimits,
  OpaquePart,
  OpaquePutResult,
  OpaqueStore,
} from './types';

/** 25 MiB per part, 200 MiB per store: embedded workbooks and media can be large, and every byte also lives in the collab document. */
export const DEFAULT_OPAQUE_LIMITS: OpaqueLimits = {
  maxPartBytes: 25 * 1024 * 1024,
  maxTotalBytes: 200 * 1024 * 1024,
};

/** Internal trusted access for `./yjs` (not exported from the package index): lets a verified peer entry in without re-hashing, and reads without copying. */
export interface OpaqueStoreInternals {
  /** The refusal reason a part of this size would get, or null. */
  refusal(size: number): 'too-large' | 'store-full' | null;
  /** Insert a part whose hash the caller has already verified and whose bytes the caller owns (no copy is made). False when it would not fit. */
  insertVerified(part: OpaquePart): boolean;
  /** The stored parts themselves, not copies; callers must not mutate them. */
  entries(): IterableIterator<OpaquePart>;
}

const internals = new WeakMap<OpaqueStore, OpaqueStoreInternals>();

export function opaqueInternals(store: OpaqueStore): OpaqueStoreInternals {
  const found = internals.get(store);
  if (!found) throw new TypeError('not an opaque store created by createOpaqueStore');
  return found;
}

const copyOf = (part: OpaquePart): OpaquePart => ({ ...part, bytes: part.bytes.slice() });

export function createOpaqueStore(limits: Partial<OpaqueLimits> = {}): OpaqueStore {
  const lim: OpaqueLimits = { ...DEFAULT_OPAQUE_LIMITS, ...limits };
  const parts = new Map<string, OpaquePart>();
  const dropped: OpaqueDropped[] = [];
  let total = 0;

  const fits = (size: number): 'too-large' | 'store-full' | null =>
    size > lim.maxPartBytes ? 'too-large' : total + size > lim.maxTotalBytes ? 'store-full' : null;

  const store: OpaqueStore = {
    limits: lim,
    async put({ bytes, contentType, role, partName }): Promise<OpaquePutResult> {
      const refuse = (reason: 'too-large' | 'store-full') => {
        const d: OpaqueDropped = {
          reason,
          size: bytes.byteLength,
          role,
          ...(partName !== undefined && { partName }),
        };
        dropped.push(d);
        return { ok: false as const, ...d };
      };
      // Cheap size check first so an oversized part is never hashed.
      if (bytes.byteLength > lim.maxPartBytes) return refuse('too-large');
      const hash = await hashBytes(bytes);
      const ref = { hash, role, ...(partName !== undefined && { partName }) };
      const existing = parts.get(hash);
      if (existing)
        return {
          ok: true,
          ref,
          deduped: true,
          contentTypeConflict: existing.contentType !== contentType,
        };
      const over = fits(bytes.byteLength);
      if (over) return refuse(over);
      parts.set(hash, { hash, contentType, bytes: bytes.slice() });
      total += bytes.byteLength;
      return { ok: true, ref, deduped: false, contentTypeConflict: false };
    },
    get: (hash) => {
      const part = parts.get(hash);
      return part && copyOf(part);
    },
    has: (hash) => parts.has(hash),
    list: () => [...parts.values()].map(copyOf),
    get size() {
      return parts.size;
    },
    get totalBytes() {
      return total;
    },
    dropped,
  };
  internals.set(store, {
    refusal: fits,
    insertVerified(part) {
      if (parts.has(part.hash)) return true;
      if (fits(part.bytes.byteLength)) return false;
      parts.set(part.hash, part);
      total += part.bytes.byteLength;
      return true;
    },
    entries: () => parts.values(),
  });
  return store;
}
