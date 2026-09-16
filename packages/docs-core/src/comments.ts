// Comment threads for a document.
//
// A thread is anchored to a span of text by a `comment` ProseMirror mark that
// carries the thread's id; that mark lives in the document body, so it is
// position-tracked and collab-synced (via y-prosemirror) for free. This store
// holds the thread CONTENT — comments, replies, status — keyed by thread id.
//
// The store owns an in-memory thread map and notifies a single listener after
// every mutation, handing it a semantic event. DocumentImpl wires that listener
// to mirror threads into the Y.Doc, trigger UI re-renders, and surface a
// `commentEvent` the host app can hook for notifications.

import type { Comment, CommentThread } from '@b32nio/spindle-shared';

// Re-export the shared comment vocabulary so consumers of docs-core get the
// full set of types without depending on @b32nio/spindle-shared directly.
export type { Comment, CommentThread, CommentStatus } from '@b32nio/spindle-shared';

let idCounter = 0;
function genId(prefix: string): string {
  return `${prefix}_${Date.now()}_${++idCounter}_${Math.random().toString(36).slice(2, 8)}`;
}

/** A short snapshot of the commented text — for sidebar display + orphan UI. */
export interface DocsCommentAnchor {
  quote: string;
}

export type DocsCommentThread = CommentThread<DocsCommentAnchor>;

/** Identity of whoever is creating a comment or resolving a thread. */
export interface CommentAuthor {
  id: string;
  name: string;
}

/**
 * A semantic description of a single comment mutation. Emitted for the local
 * user's own actions (not for threads arriving from collaborators).
 */
export type DocsCommentEvent =
  | { type: 'thread-created'; threadId: string; anchor: DocsCommentAnchor; comment: Comment; mentions: string[] }
  | { type: 'reply-added'; threadId: string; anchor: DocsCommentAnchor; comment: Comment; mentions: string[] }
  | { type: 'comment-edited'; threadId: string; anchor: DocsCommentAnchor; commentId: string }
  | { type: 'comment-deleted'; threadId: string; anchor: DocsCommentAnchor; commentId: string }
  | { type: 'thread-deleted'; threadId: string; anchor: DocsCommentAnchor }
  | { type: 'thread-resolved'; threadId: string; anchor: DocsCommentAnchor; by: CommentAuthor }
  | { type: 'thread-reopened'; threadId: string; anchor: DocsCommentAnchor; by: CommentAuthor };

function now(): string {
  return new Date().toISOString();
}

export class DocsCommentStore {
  private threads: Map<string, DocsCommentThread> = new Map();
  private changeListener: ((event: DocsCommentEvent) => void) | undefined;

  /** @internal Wired by DocumentImpl to mirror + emit on every mutation. */
  __setChangeListener(listener: ((event: DocsCommentEvent) => void) | undefined): void {
    this.changeListener = listener;
  }

  private notify(event: DocsCommentEvent): void {
    this.changeListener?.(event);
  }

  // --- Queries -------------------------------------------------------------

  getThreads(): DocsCommentThread[] {
    return [...this.threads.values()];
  }

  getThread(threadId: string): DocsCommentThread | undefined {
    return this.threads.get(threadId);
  }

  // --- Mutations -----------------------------------------------------------

  /**
   * Create a thread seeded with its root comment. The id is minted here; the
   * caller writes it into the `comment` mark covering the commented text.
   */
  addThread(
    anchor: DocsCommentAnchor,
    body: string,
    author: CommentAuthor,
    mentions: string[] = [],
  ): DocsCommentThread {
    const timestamp = now();
    const comment: Comment = {
      id: genId('cmt'),
      authorId: author.id,
      authorName: author.name,
      body,
      createdAt: timestamp,
    };
    if (mentions.length > 0) comment.mentions = [...mentions];
    const thread: DocsCommentThread = {
      id: genId('thr'),
      anchor: { ...anchor },
      status: 'open',
      createdAt: timestamp,
      comments: [comment],
    };
    this.threads.set(thread.id, thread);
    this.notify({ type: 'thread-created', threadId: thread.id, anchor: thread.anchor, comment, mentions });
    return thread;
  }

  /** Append a reply to an existing thread. */
  addReply(
    threadId: string,
    body: string,
    author: CommentAuthor,
    mentions: string[] = [],
  ): Comment | undefined {
    const thread = this.threads.get(threadId);
    if (!thread) return undefined;
    const comment: Comment = {
      id: genId('cmt'),
      authorId: author.id,
      authorName: author.name,
      body,
      createdAt: now(),
    };
    if (mentions.length > 0) comment.mentions = [...mentions];
    thread.comments.push(comment);
    this.notify({ type: 'reply-added', threadId, anchor: thread.anchor, comment, mentions });
    return comment;
  }

  /** Edit the body of a single comment. */
  editComment(threadId: string, commentId: string, body: string): void {
    const thread = this.threads.get(threadId);
    if (!thread) return;
    const comment = thread.comments.find((c) => c.id === commentId);
    if (!comment) return;
    comment.body = body;
    comment.editedAt = now();
    this.notify({ type: 'comment-edited', threadId, anchor: thread.anchor, commentId });
  }

  /** Remove a comment. Removing the last comment removes the whole thread. */
  deleteComment(threadId: string, commentId: string): void {
    const thread = this.threads.get(threadId);
    if (!thread) return;
    const idx = thread.comments.findIndex((c) => c.id === commentId);
    if (idx === -1) return;
    const anchor = thread.anchor;
    thread.comments.splice(idx, 1);
    if (thread.comments.length === 0) {
      this.threads.delete(threadId);
    }
    this.notify({ type: 'comment-deleted', threadId, anchor, commentId });
  }

  /** Remove a thread and all its comments. */
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

  /** Snapshot the threads as plain JSON (for getData / Y.Doc mirroring). */
  toJSON(): DocsCommentThread[] {
    return this.getThreads().map((t) => ({
      ...t,
      anchor: { ...t.anchor },
      comments: t.comments.map((c) => ({ ...c })),
    }));
  }

  /** Replace all threads from serialized data. Does NOT notify listeners. */
  loadJSON(threads: DocsCommentThread[] | undefined): void {
    this.threads.clear();
    if (!threads) return;
    for (const t of threads) {
      this.threads.set(t.id, {
        ...t,
        anchor: { ...t.anchor },
        comments: t.comments.map((c) => ({ ...c })),
      });
    }
  }
}
