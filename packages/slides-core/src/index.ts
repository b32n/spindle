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
export * from './scene/table';
export * from './scene/transform';
export * from './scene/snapping';
export * from './scene/align';
export * from './scene/z-order';
export * from './scene/group';

// Rich text model + schema
export * from './text/model';
export { slidesSchema } from './text/schema';

// Theme + layouts
export * from './theme/types';
export * from './theme/resolve';
export * from './theme/builtin';
export { buildPlaceholderElement } from './theme/materialize';

// Engine
export { DeckImpl } from './deck';
export type { NewElementSpec, AddSlideOptions } from './deck';
export { DeckHistory } from './history';
export type { DeckSnapshot } from './history';
export { normalizeDeckData } from './serialization';

// Utilities re-exported for convenience
export { generateId } from './utils/id';
export { EventEmitter } from '@b32nio/spindle-shared';
