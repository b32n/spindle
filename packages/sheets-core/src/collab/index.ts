// @b32nio/spindle-sheets-core/collab — Yjs-backed collaboration binding.

export {
  getWorkbookYTypes,
  getSheetYTypes,
  createSheetYMap,
  ensureSheetYMap,
  hydrateYDocFromData,
  serializeYDocToData,
  COLLAB_SCHEMA_VERSION,
  CollabSchemaVersionError,
  readSchemaVersion,
  assertSchemaVersion,
  type WorkbookYTypes,
  type SheetYTypes,
} from './y-schema';

export {
  attachCollabToWorkbook,
  type CollabHandle,
  type AttachCollabOptions,
} from './binding';

export { getCollabBlockers, CollabBlockedError, type CollabBlocker } from './blockers';
