/**
 * Lower-case hex SHA-256 (SubtleCrypto: no network, no dependency).
 * SubtleCrypto exists only in secure contexts (https, localhost, Node, workers of those): on a page served over
 * plain http from another origin this throws `Error('SubtleCrypto unavailable: …')` instead of a bare TypeError.
 */
export async function hashBytes(bytes: Uint8Array): Promise<string> {
  const subtle = globalThis.crypto?.subtle;
  if (!subtle)
    throw new Error(
      'SubtleCrypto unavailable: opaque parts are hashed with SHA-256, which needs a secure context (https or localhost)'
    );
  const digest = await subtle.digest('SHA-256', bytes as BufferSource);
  let out = '';
  for (const b of new Uint8Array(digest)) out += b.toString(16).padStart(2, '0');
  return out;
}

/** True for a string shaped like a `hashBytes` result. */
export function isOpaqueHash(value: unknown): value is string {
  return typeof value === 'string' && /^[0-9a-f]{64}$/.test(value);
}
