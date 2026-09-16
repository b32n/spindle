// @b32nio/spindle-sheets-core/collab — Yjs-backed collaboration binding.

export {
  getWorkbookYTypes,
  getSheetYTypes,
  createSheetYMap,
  ensureSheetYMap,
  hydrateYDocFromData,
  serializeYDocToData,
  type WorkbookYTypes,
  type SheetYTypes,
} from './y-schema';

export {
  attachCollabToWorkbook,
  type CollabHandle,
  type AttachCollabOptions,
} from './binding';
