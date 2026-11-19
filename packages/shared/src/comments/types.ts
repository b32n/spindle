// Editor-agnostic comment-thread model.
//
// A thread is anchored to some piece of content (the `anchor` is supplied by
// each editor — a cell for sheets, a range/mark for docs) and holds an ordered
// list of comments: comments[0] is the root, the rest are replies.

export type CommentStatus = 'open' | 'resolved';

export interface Comment {
  id: string;
  /** Stable identity of the author (e.g. a CollabIdentity.userId). */
  authorId: string;
  /**
   * Display name snapshotted at write time, so a thread still renders
   * correctly when the author is offline or unknown to the current peer.
   */
  authorName: string;
  /** Author's initials when the source records them (DOCX comments do). */
  authorInitials?: string;
  body: string;
  /** ISO-8601 timestamp. */
  createdAt: string;
  /** ISO-8601 timestamp, set when the body is edited. */
  editedAt?: string;
  /** IDs of users @-mentioned in the body. */
  mentions?: string[];
}

export interface CommentThread<Anchor = unknown> {
  id: string;
  /** Editor-specific pointer to the commented content. */
  anchor: Anchor;
  status: CommentStatus;
  /** Ordered comments — [0] is the root, the rest are replies. */
  comments: Comment[];
  /** ISO-8601 timestamp of the root comment. */
  createdAt: string;
  resolvedBy?: string;
  resolvedAt?: string;
}
