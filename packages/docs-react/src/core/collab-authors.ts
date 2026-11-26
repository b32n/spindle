/**
 * Who the authors of tracked changes are in a shared document: the name a suggestion is written under, and the colours
 * people in the room are drawn in (the colour of their cursor).
 */

import { collabAuthorColors } from '@b32nio/spindle-docs-core';

/** What of a collaboration handle presence is read from (its awareness). */
type Presence = Parameters<typeof collabAuthorColors>[0];

/** The author of a suggestion: the name the user joined the room under, else the current user's name. */
export function suggestionAuthor(handle: { identity: { displayName?: string } } | null | undefined, fallback: string): string {
  // An empty name would write `w:author=""` and match the next empty-name author, so it is never used.
  return handle?.identity.displayName || fallback || 'Unknown';
}

/** True when two maps of author colours are the same. */
export function sameAuthorColors(a: ReadonlyMap<string, string>, b: ReadonlyMap<string, string>): boolean {
  return a.size === b.size && [...b].every(([name, color]) => a.get(name) === color);
}

/**
 * Tell `onChange` the colours of the people in the room now and whenever the room changes; returns what stops it.
 * The caller decides whether the colours changed (see `sameAuthorColors`): awareness also changes when a cursor moves.
 */
export function watchAuthorColors(handle: Presence, onChange: (colors: Map<string, string>) => void): () => void {
  const update = () => onChange(collabAuthorColors(handle));
  update();
  handle.awareness.on('change', update);
  return () => handle.awareness.off('change', update);
}
