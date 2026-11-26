// Who is in a shared document, as far as the colours of tracked changes are concerned.

import type { CollabHandle } from './binding';

const HEX = /^#[0-9a-f]{6}$/i;

/**
 * The colour each person in the room is shown in (the colour of their cursor), by the name their changes are written under.
 * Several people under one name are one author, in the smallest of their colours. An author who has left the room has no entry: their changes are drawn in a colour the document's authors share out.
 */
export function collabAuthorColors(handle: Pick<CollabHandle, 'awareness'>): Map<string, string> {
  const colors = new Map<string, string>();
  handle.awareness.getStates().forEach((state) => {
    const user = (state as { user?: { name?: unknown; color?: unknown } }).user;
    if (!user || typeof user.name !== 'string' || typeof user.color !== 'string') return;
    // Two people under one name can have two colours; every client picks the same one (the smallest), not the one met last.
    const known = colors.get(user.name);
    const better = known === undefined || (HEX.test(user.color) && (!HEX.test(known) || user.color.toLowerCase() < known.toLowerCase()));
    if (better) colors.set(user.name, user.color);
  });
  return colors;
}
