// Who is in a shared document, as far as the colours of tracked changes are concerned.

import type { CollabHandle } from './binding';
import { HEX_COLOR } from '../revisions';

/** Compared in lower case; two spellings of one colour are told apart by the raw text, so every client picks the same one. */
const smaller = (a: string, b: string): boolean => {
  const [x, y] = [a.toLowerCase(), b.toLowerCase()];
  return x < y || (x === y && a < b);
};

/**
 * The colour each person in the room is shown in (the colour of their cursor), by the name their changes are written under.
 * The name is trimmed, as the author written on a suggestion is. The colour strings are what the clients announced (not checked
 * to be colours, apart from preferring a six-digit hex colour over any other): a caller that uses one as a CSS value must validate it
 * itself (`assignAuthorColors` does). Several people under one name are one author, in the smallest of their colours. An author who has left the room has no entry: their changes are drawn in a colour the document's authors share out.
 */
export function collabAuthorColors(handle: Pick<CollabHandle, 'awareness'>): Map<string, string> {
  const colors = new Map<string, string>();
  handle.awareness.getStates().forEach((state) => {
    const user = (state as { user?: { name?: unknown; color?: unknown } } | null | undefined)?.user;
    if (!user || typeof user.name !== 'string' || typeof user.color !== 'string') return;
    const name = user.name.trim();
    if (!name) return;
    // Two people under one name can have two colours; every client picks the same one (the smallest), not the one met last.
    const known = colors.get(name);
    // Valid hex colours beat invalid ones; between two of a kind the smaller (compared in lower case) wins, so the order the room lists people in does not matter.
    const valid = HEX_COLOR.test(user.color);
    const knownValid = known !== undefined && HEX_COLOR.test(known);
    const better = known === undefined || (valid && !knownValid) || (valid === knownValid && smaller(user.color, known));
    if (better) colors.set(name, user.color);
  });
  return colors;
}
