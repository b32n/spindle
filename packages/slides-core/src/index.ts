// @b32nio/spindle-slides-core
// Core presentation engine (no Yjs — import '@b32nio/spindle-slides-core/collab' for that).

// Deck-layer types + events
export * from './types';

// Comments
export * from './comments';

// Scene layer
export * from './scene/types';
export * from './scene/elements';
export * from './scene/fractional-index';
export * from './scene/geometry';
export * from './scene/connector';
export * from './scene/shape-presets';
export * from './scene/custom-geometry';
export * from './scene/adjustment-handles';
export * from './scene/table';
export * from './scene/transform';
export * from './scene/snapping';
export * from './scene/align';
export * from './scene/z-order';
export * from './scene/group';
export * from './scene/group-tree';
export * from './scene/crop';
export { opaqueLabel } from './opaque/label';
export { pruneTiming, timingTargets } from './opaque/timing';

// Rich text model + schema
export * from './text/model';
export * from './text/body';
export * from './text/paragraph';
export * from './text/run';
export { slidesSchema } from './text/schema';

// Theme + layouts
export * from './theme/types';
export * from './theme/resolve';
export * from './theme/fill';
export * from './theme/builtin';
export { buildPlaceholderElement, resolveLevelStyle } from './theme/materialize';
export * from './slide-size';

// Engine
export { DeckImpl } from './deck';
export type { NewElementSpec, AddSlideOptions } from './deck';
export { DeckHistory } from './history';
export type { DeckSnapshot } from './history';
export { normalizeDeckData } from './serialization';
export { getCollabBlockers, CollabBlockedError } from './collab/blockers';
export type { CollabBlocker } from './collab/blockers';

// Utilities re-exported for convenience
export { generateId } from './utils/id';
export { EventEmitter } from '@b32nio/spindle-shared';

// PPTX import / export (the reader and writer in `@b32nio/spindle-io/pptx` load on first use)
export { importFromPptx, exportToPptx, inspectPptx } from './export/pptx';
export type { ImportFromPptxOptions, ExportToPptxOptions, ExportToPptxResult } from './export/pptx';
