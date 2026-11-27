import { hashBytes } from './hash';
import type { OpaqueDropped, OpaqueLimits, OpaquePart, OpaquePutResult, OpaqueStore } from './types';

/** 25 MiB per part, 200 MiB per store: embedded workbooks and media can be large, and every byte also lives in the collab document. */
export const DEFAULT_OPAQUE_LIMITS: OpaqueLimits = { maxPartBytes: 25 * 1024 * 1024, maxTotalBytes: 200 * 1024 * 1024 };

export function createOpaqueStore(limits: Partial<OpaqueLimits> = {}): OpaqueStore {
  const lim: OpaqueLimits = { ...DEFAULT_OPAQUE_LIMITS, ...limits };
  const parts = new Map<string, OpaquePart>();
  const dropped: OpaqueDropped[] = [];
  let total = 0;

  const fits = (size: number): 'too-large' | 'store-full' | null =>
    size > lim.maxPartBytes ? 'too-large' : total + size > lim.maxTotalBytes ? 'store-full' : null;

  return {
    limits: lim,
    async put({ bytes, contentType, role, partName }): Promise<OpaquePutResult> {
      const refuse = (reason: 'too-large' | 'store-full') => {
        const d: OpaqueDropped = { reason, size: bytes.byteLength, role, ...(partName !== undefined && { partName }) };
        dropped.push(d);
        return { ok: false as const, ...d };
      };
      // Cheap size check first so an oversized part is never hashed.
      if (bytes.byteLength > lim.maxPartBytes) return refuse('too-large');
      const hash = await hashBytes(bytes);
      const ref = { hash, role, ...(partName !== undefined && { partName }) };
      if (parts.has(hash)) return { ok: true, ref, deduped: true };
      const over = fits(bytes.byteLength);
      if (over) return refuse(over);
      parts.set(hash, { hash, contentType, bytes: bytes.slice() });
      total += bytes.byteLength;
      return { ok: true, ref, deduped: false };
    },
    putVerified(part) {
      if (parts.has(part.hash)) return true;
      if (fits(part.bytes.byteLength)) return false;
      parts.set(part.hash, { ...part, bytes: part.bytes.slice() });
      total += part.bytes.byteLength;
      return true;
    },
    get: (hash) => parts.get(hash),
    has: (hash) => parts.has(hash),
    list: () => [...parts.values()],
    get size() {
      return parts.size;
    },
    get totalBytes() {
      return total;
    },
    dropped,
  };
}
