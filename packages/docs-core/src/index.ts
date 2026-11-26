// @b32nio/spindle-docs-core
// Core document editor engine

// Types
export * from './types';

// Comments
export * from './comments';

// Document
export { DocumentImpl, createDocument, createSection } from './document';

// Collaboration: what a document needs to be attachable, and why attaching can fail
export { getCollabBlockers, CollabBlockedError, type CollabBlocker } from './collab/blockers';
export { CollabSchemaVersionError, COLLAB_SCHEMA_VERSION } from './collab/y-schema';

// Blocks
export * from './blocks';

// Export
export * from './export';

// Selection
export {
  createCursorPosition,
  createCollapsedSelection,
  createSelection,
  isSelectionCollapsed,
  getSelectionStart,
  getSelectionEnd,
  getBlockTextLength,
  getBlockText,
  findRunAtOffset,
  runOffsetToBlockOffset,
} from './selection';

// History
export { DocumentHistory, type HistoryConfig } from './history';

// Style Pools
export { TextStylePoolImpl, ParagraphStylePoolImpl } from './style-pool';

// Event Emitter
export { DocumentEventEmitter } from './event-emitter';

// ProseMirror integration
export * from './prosemirror';

