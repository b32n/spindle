// useComments — a thin React wrapper over the deck's SlidesCommentStore. Re-
// renders on `commentChange` (local or remote) and threads new comments under
// the provider's currentUser.

import { useEffect, useReducer } from 'react';
import type { CommentAuthor, ElementCommentAnchor, SlidesCommentThread } from '@b32nio/spindle-slides-core';
import { useDeckContext } from '../context/DeckContext';

export interface UseComments {
  threads: SlidesCommentThread[];
  currentUser?: CommentAuthor;
  mentionableUsers: CommentAuthor[];
  threadsForElement(elementId: string): SlidesCommentThread[];
  isOrphaned(threadId: string): boolean;
  addThread(anchor: ElementCommentAnchor, body: string, mentions?: string[]): void;
  addReply(threadId: string, body: string, mentions?: string[]): void;
  resolve(threadId: string): void;
  reopen(threadId: string): void;
  deleteThread(threadId: string): void;
  deleteComment(threadId: string, commentId: string): void;
  editComment(threadId: string, commentId: string, body: string): void;
}

export function useComments(): UseComments {
  const { deck, currentUser, mentionableUsers } = useDeckContext();
  const [, force] = useReducer((n) => n + 1, 0);
  // Orphan status depends on element existence, so refresh on element/deck
  // changes too — not just comment mutations.
  useEffect(() => {
    const offs = [
      deck.on('commentChange', force as () => void),
      deck.on('elementDelete', force as () => void),
      deck.on('elementAdd', force as () => void),
      deck.on('deckChange', force as () => void),
    ];
    return () => offs.forEach((o) => o());
  }, [deck]);

  const store = deck.getComments();
  return {
    threads: store.getThreads(),
    currentUser,
    mentionableUsers,
    threadsForElement: (id) => store.getThreadsForElement(id),
    isOrphaned: (tid) => deck.isThreadOrphaned(tid),
    addThread: (anchor, body, mentions) => {
      if (currentUser) store.addThread(anchor, body, currentUser, mentions);
    },
    addReply: (tid, body, mentions) => {
      if (currentUser) store.addReply(tid, body, currentUser, mentions);
    },
    resolve: (tid) => {
      if (currentUser) store.resolveThread(tid, currentUser);
    },
    reopen: (tid) => {
      if (currentUser) store.reopenThread(tid, currentUser);
    },
    deleteThread: (tid) => store.deleteThread(tid),
    deleteComment: (tid, cid) => store.deleteComment(tid, cid),
    editComment: (tid, cid, body) => store.editComment(tid, cid, body),
  };
}
