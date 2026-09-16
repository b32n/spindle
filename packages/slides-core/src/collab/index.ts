// Collaboration entry (@b32nio/spindle-slides-core/collab).
//
// This subpath is the only place that imports Yjs, keeping it out of the base
// bundle. DeckImpl.attachCollab() lazy-imports the binding, so consumers who
// never call it never load Yjs.

export { attachCollabToDeck, LOCAL_ORIGIN } from './binding';
export type { CollabHandle, AttachCollabOptions } from './binding';
export {
  getDeckYTypes,
  hydrateYDocFromData,
  serializeYDocToData,
  createElementYMap,
  yMapToElement,
} from './y-schema';
export type { DeckYTypes } from './y-schema';
