/**
 * An unsupported package part (chart, media, custom XML…) kept byte for byte.
 * `hash` is the lower-case hex SHA-256 of `bytes` alone, so equal bytes are stored once
 * whatever their name or content type.
 */
export interface OpaquePart {
  hash: string;
  /** MIME / OOXML content type the part is registered under, e.g. `application/vnd.openxmlformats-officedocument.drawingml.chart+xml`. */
  contentType: string;
  bytes: Uint8Array;
}

/**
 * How a document model points at a stored part. `partName` is the name in the source package
 * (informational: exporters pick a collision-free name and rewrite relationships); `role` says
 * what the part is for, e.g. `chart`, `media`, `customXml`.
 */
export interface OpaqueRef {
  hash: string;
  partName?: string;
  role: string;
}

/** Size caps. A part above `maxPartBytes`, or one that would push the store past `maxTotalBytes`, is refused (never truncated). */
export interface OpaqueLimits {
  maxPartBytes: number;
  maxTotalBytes: number;
}

export type OpaquePutResult =
  | { ok: true; ref: OpaqueRef; deduped: boolean }
  | { ok: false; reason: 'too-large' | 'store-full'; size: number; role: string; partName?: string };

/** A part the store refused, kept so the importer can report it instead of dropping it silently. */
export interface OpaqueDropped {
  reason: 'too-large' | 'store-full';
  size: number;
  role: string;
  partName?: string;
}

/** Content-addressed part store. In memory; the per-core Y helper (`./yjs`) mirrors it into a document. */
export interface OpaqueStore {
  readonly limits: OpaqueLimits;
  /** Store a part (hashing it) unless it is over a cap. Never throws on size; the refusal is the result and is also recorded in `dropped`. */
  put(input: { bytes: Uint8Array; contentType: string; role: string; partName?: string }): Promise<OpaquePutResult>;
  /** Insert an already-hashed part (e.g. read from a Y map); the hash is not re-verified. Applies the same caps. */
  putVerified(part: OpaquePart): boolean;
  get(hash: string): OpaquePart | undefined;
  has(hash: string): boolean;
  /** Parts in insertion order. */
  list(): OpaquePart[];
  /** Number of distinct parts. */
  readonly size: number;
  /** Sum of the distinct parts' byte lengths. */
  readonly totalBytes: number;
  /** Everything `put` refused so far. */
  readonly dropped: readonly OpaqueDropped[];
}
