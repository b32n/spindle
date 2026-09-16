// Comment threads for a deck.
//
// A thread is anchored to a single element by {slideId, elementId}. If that
// element is deleted the thread is "orphaned" — kept (so nothing is silently
// lost) but no longer resolves to an element; the panel lists it under "no
// longer attached" (the sheets rule).
//
// The store owns the in-memory thread map and notifies a single listener after
// every mutation with a semantic CommentMutationEvent. DeckImpl wires that
// listener to emit `commentChange` (UI) + `commentEvent` (host hook, local
// actions only); the collab binding mirrors threads into the Y.Doc.

import type { Comment } from '@b32nio/spindle-shared';
import { generateId } from './utils/id';
import type { ElementCommentAnchor, SlidesCommentThread } from './types';

export type { Comment, CommentThread, CommentStatus } from '@b32nio/spindle-shared';
export type { ElementCommentAnchor, SlidesCommentThread } from './types';

/** Identity of whoever is creating a comment or resolving a thread. */
export interface CommentAuthor {
  id: string;
  name: string;
}

/** A semantic description of a single comment mutation (local actions only). */
export type CommentMutationEvent =
  | { type: 'thread-created'; threadId: string; anchor: ElementCommentAnchor; comment: Comment; mentions: string[] }
  | { type: 'reply-added'; threadId: string; anchor: ElementCommentAnchor; comment: Comment; mentions: string[] }
  | { type: 'comment-edited'; threadId: string; anchor: ElementCommentAnchor; commentId: string }
  | { type: 'comment-deleted'; threadId: string; anchor: ElementCommentAnchor; commentId: string }
  | { type: 'thread-deleted'; threadId: string; anchor: ElementCommentAnchor }
  | { type: 'thread-resolved'; threadId: string; anchor: ElementCommentAnchor; by: CommentAuthor }
  | { type: 'thread-reopened'; threadId: string; anchor: ElementCommentAnchor; by: CommentAuthor };

export type SlidesCommentEvent = CommentMutationEvent;

function now(): string {
  return new Date().toISOString();
}

export class SlidesCommentStore {
  private threads: Map<string, SlidesCommentThread> = new Map();
  private changeListener: ((event: CommentMutationEvent) => void) | undefined;

  /** @internal Wired by DeckImpl to mirror + emit on every mutation. */
  __setChangeListener(listener: ((event: CommentMutationEvent) => void) | undefined): void {
    this.changeListener = listener;
  }

  private notify(event: CommentMutationEvent): void {
    this.changeListener?.(event);
  }

  // --- Queries -------------------------------------------------------------

  getThreads(): SlidesCommentThread[] {
    return [...this.threads.values()];
  }

  getThread(threadId: string): SlidesCommentThread | undefined {
    return this.threads.get(threadId);
  }

  getThreadsForElement(elementId: string): SlidesCommentThread[] {
    return this.getThreads().filter((t) => t.anchor.elementId === elementId);
  }

  /** True when the element has at least one open (unresolved) thread. */
  hasOpenThread(elementId: string): boolean {
    for (const t of this.threads.values()) {
      if (t.anchor.elementId === elementId && t.status === 'open') return true;
    }
    return false;
  }

  // --- Mutations -----------------------------------------------------------

  addThread(anchor: ElementCommentAnchor, body: string, author: CommentAuthor, mentions: string[] = []): SlidesCommentThread {
    const timestamp = now();
    const comment: Comment = { id: generateId(), authorId: author.id, authorName: author.name, body, createdAt: timestamp };
    if (mentions.length > 0) comment.mentions = [...mentions];
    const thread: SlidesCommentThread = { id: generateId(), anchor: { ...anchor }, status: 'open', createdAt: timestamp, comments: [comment] };
    this.threads.set(thread.id, thread);
    this.notify({ type: 'thread-created', threadId: thread.id, anchor: thread.anchor, comment, mentions });
    return thread;
  }

  addReply(threadId: string, body: string, author: CommentAuthor, mentions: string[] = []): Comment | undefined {
    const thread = this.threads.get(threadId);
    if (!thread) return undefined;
    const comment: Comment = { id: generateId(), authorId: author.id, authorName: author.name, body, createdAt: now() };
    if (mentions.length > 0) comment.mentions = [...mentions];
    thread.comments.push(comment);
    this.notify({ type: 'reply-added', threadId, anchor: thread.anchor, comment, mentions });
    return comment;
  }

  editComment(threadId: string, commentId: string, body: string): void {
    const thread = this.threads.get(threadId);
    if (!thread) return;
    const comment = thread.comments.find((c) => c.id === commentId);
    if (!comment) return;
    comment.body = body;
    comment.editedAt = now();
    this.notify({ type: 'comment-edited', threadId, anchor: thread.anchor, commentId });
  }

  deleteComment(threadId: string, commentId: string): void {
    const thread = this.threads.get(threadId);
    if (!thread) return;
    const idx = thread.comments.findIndex((c) => c.id === commentId);
    if (idx === -1) return;
    const anchor = thread.anchor;
    thread.comments.splice(idx, 1);
    if (thread.comments.length === 0) this.threads.delete(threadId);
    this.notify({ type: 'comment-deleted', threadId, anchor, commentId });
  }

  deleteThread(threadId: string): void {
    const thread = this.threads.get(threadId);
    if (!thread) return;
    const anchor = thread.anchor;
    this.threads.delete(threadId);
    this.notify({ type: 'thread-deleted', threadId, anchor });
  }

  resolveThread(threadId: string, by: CommentAuthor): void {
    const thread = this.threads.get(threadId);
    if (!thread || thread.status === 'resolved') return;
    thread.status = 'resolved';
    thread.resolvedBy = by.id;
    thread.resolvedAt = now();
    this.notify({ type: 'thread-resolved', threadId, anchor: thread.anchor, by });
  }

  reopenThread(threadId: string, by: CommentAuthor): void {
    const thread = this.threads.get(threadId);
    if (!thread || thread.status === 'open') return;
    thread.status = 'open';
    delete thread.resolvedBy;
    delete thread.resolvedAt;
    this.notify({ type: 'thread-reopened', threadId, anchor: thread.anchor, by });
  }

  // --- Serialization -------------------------------------------------------

  toJSON(): SlidesCommentThread[] {
    return this.getThreads().map((t) => ({ ...t, anchor: { ...t.anchor }, comments: t.comments.map((c) => ({ ...c })) }));
  }

  /** Replace all threads from serialized data. Does NOT notify listeners. */
  loadJSON(threads: SlidesCommentThread[] | undefined): void {
    this.threads.clear();
    if (!threads) return;
    for (const t of threads) {
      this.threads.set(t.id, { ...t, anchor: { ...t.anchor }, comments: t.comments.map((c) => ({ ...c })) });
    }
  }
}
