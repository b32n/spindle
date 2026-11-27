/** Lower-case hex SHA-256 (SubtleCrypto: no network, no dependency). */
export async function hashBytes(bytes: Uint8Array): Promise<string> {
  const digest = await globalThis.crypto.subtle.digest('SHA-256', bytes as BufferSource);
  let out = '';
  for (const b of new Uint8Array(digest)) out += b.toString(16).padStart(2, '0');
  return out;
}

/** True for a string shaped like a `hashBytes` result. */
export function isOpaqueHash(value: unknown): value is string {
  return typeof value === 'string' && /^[0-9a-f]{64}$/.test(value);
}
