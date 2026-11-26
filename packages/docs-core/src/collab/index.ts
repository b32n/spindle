// @b32nio/spindle-docs-core/collab — Yjs-backed collaboration binding.

export {
  getYDocFields,
  hydrateYDocFromData,
  serializeYDocToData,
  COLLAB_SCHEMA_VERSION,
  CollabSchemaVersionError,
  readSchemaVersion,
  assertSchemaVersion,
  type YDocFields,
} from './y-schema';

export {
  attachCollabToYDoc,
  type CollabHandle,
  type AttachCollabOptions,
} from './binding';

export { getCollabBlockers, CollabBlockedError, type CollabBlocker } from './blockers';
