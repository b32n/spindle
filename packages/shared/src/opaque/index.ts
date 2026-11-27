// Opaque part store (`@b32nio/spindle-shared/opaque`): unsupported package parts kept byte for byte,
// content-addressed and size-capped, plus a duck-typed Y map helper and loss-item builders.
export * from './types';
export { hashBytes, isOpaqueHash } from './hash';
export { createOpaqueStore, DEFAULT_OPAQUE_LIMITS } from './store';
export { encodeOpaqueEntry, decodeOpaqueEntry, writeOpaqueParts, readOpaqueParts } from './yjs';
export type { OpaqueYMap, OpaqueReadResult } from './yjs';
export { keptLossItem, droppedLossItem } from './loss';
