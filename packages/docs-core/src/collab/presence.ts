// Who is in a shared document, as far as the colours of tracked changes are concerned.

import type { CollabHandle } from './binding';

/**
 * The colour each person in the room is shown in (the colour of their cursor), by the name their changes are written under.
 * An author who has left the room has no entry: their changes are drawn in a colour the document's authors share out.
 */
export function collabAuthorColors(handle: Pick<CollabHandle, 'awareness'>): Map<string, string> {
  const colors = new Map<string, string>();
  handle.awareness.getStates().forEach((state) => {
    const user = (state as { user?: { name?: unknown; color?: unknown } }).user;
    if (user && typeof user.name === 'string' && typeof user.color === 'string') colors.set(user.name, user.color);
  });
  return colors;
}
