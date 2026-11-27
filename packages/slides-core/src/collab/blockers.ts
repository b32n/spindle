// What the collaboration (Y.Doc) layout cannot hold.
//
// A deck that contains such a construct must not be attached to a room: the shared copy would silently lose it.
// `getCollabBlockers` lists those constructs so a host can tell the user why collaboration is unavailable for this
// deck, and `DeckImpl.attachCollab` refuses with the same list. The rule for new deck constructs is that they are
// representable in the Y.Doc (an additive Y key) before they ship; a construct that is temporarily model-only is added
// here as a check and removed again when its Y representation lands. This file imports no Yjs (it is part of the base
// bundle).

import type { DeckData } from '../types';

/** One kind of content that collaboration cannot carry, with how many places in the deck have it. */
export interface CollabBlocker {
  /** Stable id, e.g. `collab.unsupportedChart`. */
  id: string;
  /** A sentence a host can show to the user. */
  message: string;
  /** How many places in the deck are affected. */
  count: number;
}

/** Thrown by `attachCollab` when the deck has content collaboration cannot hold. */
export class CollabBlockedError extends Error {
  readonly blockers: CollabBlocker[];
  constructor(blockers: CollabBlocker[]) {
    super(
      'Collaboration is not available for this deck: ' +
        blockers.map((b) => `${b.message} (${b.count})`).join('; '),
    );
    this.name = 'CollabBlockedError';
    this.blockers = blockers;
  }
}

/**
 * The content of a deck that collaboration cannot hold (empty when the deck can be attached).
 *
 * Every construct the deck model has today is carried by the Y layout, so the list is empty for now. Each later change
 * that adds a construct either stores it in additive Y keys in the same change or registers a blocker here.
 */
export function getCollabBlockers(_deck: Pick<DeckData, 'slides'>): CollabBlocker[] {
  return [];
}
