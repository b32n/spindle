// Comment threads for a sheet.
//
// A thread is anchored to a single cell by its STABLE rowId/colId, so it stays
// attached to the right cell across row/column inserts, deletes and sorts. If
// the anchored row or column is itself deleted the thread is "orphaned" — it is
// kept (so nothing is silently lost) but no longer resolves to a cell.
//
// The store owns the in-memory thread map and notifies a single listener after
// every mutation, handing it a semantic CommentMutationEvent. WorkbookImpl
// wires that listener to mirror threads into the Y.Doc, emit a `commentChange`
// event for the UI, and emit a `commentEvent` the host app can hook for
// notifications.

import type { Comment, CommentThread } from '@b32nio/spindle-shared';
import { generateId } from './utils/id';

// Re-export the shared comment vocabulary so consumers of sheets-core get the
// full set of types without depending on @b32nio/spindle-shared directly.
export type { Comment, CommentThread, CommentStatus } from '@b32nio/spindle-shared';

/** Stable-ID pointer to the cell a thread is anchored to. */
export interface CellCommentAnchor {
  rowId: string;
  colId: string;
}

export type SheetCommentThread = CommentThread<CellCommentAnchor>;

/** Identity of whoever is creating a comment or resolving a thread. */
export interface CommentAuthor {
  id: string;
  name: string;
}

/**
 * A semantic description of a single comment mutation. Emitted for the local
 * user's own actions (not for threads arriving from collaborators).
 */
export type CommentMutationEvent =
  | { type: 'thread-created'; threadId: string; anchor: CellCommentAnchor; comment: Comment; mentions: string[] }
  | { type: 'reply-added'; threadId: string; anchor: CellCommentAnchor; comment: Comment; mentions: string[] }
  | { type: 'comment-edited'; threadId: string; anchor: CellCommentAnchor; commentId: string }
  | { type: 'comment-deleted'; threadId: string; anchor: CellCommentAnchor; commentId: string }
  | { type: 'thread-deleted'; threadId: string; anchor: CellCommentAnchor }
  | { type: 'thread-resolved'; threadId: string; anchor: CellCommentAnchor; by: CommentAuthor }
  | { type: 'thread-reopened'; threadId: string; anchor: CellCommentAnchor; by: CommentAuthor };

/** A CommentMutationEvent enriched with the sheet it occurred on. */
export type SheetCommentEvent = CommentMutationEvent & { sheetId: string };

function now(): string {
  return new Date().toISOString();
}

export class CommentStore {
  private threads: Map<string, SheetCommentThread> = new Map();
  private changeListener: ((event: CommentMutationEvent) => void) | undefined;

  /** @internal Wired by WorkbookImpl to mirror + emit on every mutation. */
  __setChangeListener(listener: ((event: CommentMutationEvent) => void) | undefined): void {
    this.changeListener = listener;
  }

  private notify(event: CommentMutationEvent): void {
    this.changeListener?.(event);
  }

  // --- Queries -------------------------------------------------------------

  /** All threads, in insertion order. */
  getThreads(): SheetCommentThread[] {
    return [...this.threads.values()];
  }

  getThread(threadId: string): SheetCommentThread | undefined {
    return this.threads.get(threadId);
  }

  getThreadsForCell(rowId: string, colId: string): SheetCommentThread[] {
    return this.getThreads().filter(
      (t) => t.anchor.rowId === rowId && t.anchor.colId === colId,
    );
  }

  /** True when the cell has at least one open (unresolved) thread. */
  hasOpenThread(rowId: string, colId: string): boolean {
    for (const t of this.threads.values()) {
      if (t.anchor.rowId === rowId && t.anchor.colId === colId && t.status === 'open') {
        return true;
      }
    }
    return false;
  }

  // --- Mutations -----------------------------------------------------------

  /** Create a new thread anchored to a cell, seeded with its root comment. */
  addThread(
    anchor: CellCommentAnchor,
    body: string,
    author: CommentAuthor,
    mentions: string[] = [],
  ): SheetCommentThread {
    const timestamp = now();
    const comment: Comment = {
      id: generateId(),
      authorId: author.id,
      authorName: author.name,
      body,
      createdAt: timestamp,
    };
    if (mentions.length > 0) comment.mentions = [...mentions];
    const thread: SheetCommentThread = {
      id: generateId(),
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
      id: generateId(),
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
  toJSON(): SheetCommentThread[] {
    return this.getThreads().map((t) => ({
      ...t,
      anchor: { ...t.anchor },
      comments: t.comments.map((c) => ({ ...c })),
    }));
  }

  /** Replace all threads from serialized data. Does NOT notify listeners. */
  loadJSON(threads: SheetCommentThread[] | undefined): void {
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
