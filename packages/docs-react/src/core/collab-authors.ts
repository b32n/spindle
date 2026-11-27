/**
 * Who the authors of tracked changes are in a shared document: the name a suggestion is written under, and the colours
 * people in the room are drawn in (the colour of their cursor).
 */

import { assignAuthorColors, collabAuthorColors } from '@b32nio/spindle-docs-core';

/** What of a collaboration handle presence is read from (its awareness). */
type Presence = Parameters<typeof collabAuthorColors>[0];

/** The author of a suggestion: the name the user joined the room under, else the current user's name. */
export function suggestionAuthor(handle: { identity: { displayName?: string } } | null | undefined, fallback: string): string {
  // A blank display name would write suggestions with no author (`w:author=""`, which the next blank author would share).
  return handle?.identity.displayName?.trim() || fallback.trim() || 'Unknown';
}

/**
 * The authors of the changes with the colour the markup shows them in: the colours are shared out among every author of the
 * document (`palette`), those who are in the room keep their cursor colour (`presence`). An author who left the room has
 * no cursor colour any more and gets one from the palette, so a colour can change when somebody leaves.
 */
export function authorsWithColors<T extends { name: string; color: string }>(counts: readonly T[], palette: readonly string[], presence: ReadonlyMap<string, string>): T[] {
  const colors = assignAuthorColors(palette, presence);
  return counts.map((a) => ({ ...a, color: colors.get(a.name) ?? a.color }));
}

/** True when two maps of author colours are the same. */
export function sameAuthorColors(a: ReadonlyMap<string, string>, b: ReadonlyMap<string, string>): boolean {
  return a.size === b.size && [...b].every(([name, color]) => a.get(name) === color);
}

/** The colours as one string that does not depend on the order the map was built in, so a new map with the same colours is the same key. */
export function authorColorsKey(colors: ReadonlyMap<string, string> | undefined): string {
  return colors ? JSON.stringify([...colors.entries()].sort()) : '';
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
