import React, { memo, useRef, useState } from 'react';
import { X, Check, RotateCcw, Trash2, MessageSquare } from 'lucide-react';
import { columnIndexToLabel } from '@b32nio/spindle-sheets-core';
import type { Comment, CommentAuthor, SheetCommentThread } from '@b32nio/spindle-sheets-core';
import { useComments } from '../hooks/useComments';
import { useWorkbook } from '../context/WorkbookContext';

// ============================================================================
// Helpers
// ============================================================================

const FONT = '"Inter", "SF Pro Display", -apple-system, BlinkMacSystemFont, sans-serif';

const AVATAR_COLORS = [
  '#6366f1', '#0ea5e9', '#10b981', '#f59e0b',
  '#ef4444', '#ec4899', '#8b5cf6', '#14b8a6',
];

function authorColor(id: string): string {
  let hash = 0;
  for (let i = 0; i < id.length; i++) hash = (hash * 31 + id.charCodeAt(i)) >>> 0;
  return AVATAR_COLORS[hash % AVATAR_COLORS.length];
}

function initials(name: string): string {
  const parts = name.trim().split(/\s+/).filter(Boolean);
  if (parts.length === 0) return '?';
  if (parts.length === 1) return parts[0][0].toUpperCase();
  return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
}

function relativeTime(iso: string): string {
  const diff = Date.now() - new Date(iso).getTime();
  const s = Math.floor(diff / 1000);
  if (s < 45) return 'just now';
  const m = Math.floor(s / 60);
  if (m < 60) return `${m}m ago`;
  const h = Math.floor(m / 60);
  if (h < 24) return `${h}h ago`;
  const d = Math.floor(h / 24);
  if (d < 7) return `${d}d ago`;
  return new Date(iso).toLocaleDateString();
}

const mentionStyle: React.CSSProperties = {
  color: '#6366f1',
  fontWeight: 600,
};

/** Render a comment body, highlighting @Name spans for known mentioned users. */
function renderBody(body: string, mentionNames: string[]): React.ReactNode {
  if (mentionNames.length === 0) return body;
  // Longest names first so "@John Smith" wins over "@John".
  const escaped = mentionNames
    .slice()
    .sort((a, b) => b.length - a.length)
    .map((n) => n.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'));
  const re = new RegExp('@(?:' + escaped.join('|') + ')', 'g');
  const parts: React.ReactNode[] = [];
  let last = 0;
  let m: RegExpExecArray | null;
  while ((m = re.exec(body)) !== null) {
    if (m.index > last) parts.push(body.slice(last, m.index));
    parts.push(
      <span key={m.index} style={mentionStyle}>
        {m[0]}
      </span>,
    );
    last = m.index + m[0].length;
  }
  if (last < body.length) parts.push(body.slice(last));
  return parts;
}

// ============================================================================
// Avatar
// ============================================================================

const Avatar = memo(function Avatar({ name, id, size = 26 }: { name: string; id: string; size?: number }) {
  return (
    <div
      style={{
        width: size,
        height: size,
        borderRadius: '50%',
        backgroundColor: authorColor(id),
        color: '#fff',
        fontSize: size * 0.42,
        fontWeight: 600,
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        flexShrink: 0,
        userSelect: 'none',
      }}
    >
      {initials(name)}
    </div>
  );
});

// ============================================================================
// Composer
// ============================================================================

interface ComposerProps {
  placeholder: string;
  submitLabel: string;
  autoFocus?: boolean;
  onSubmit: (body: string, mentions: string[]) => void;
}

function Composer({ placeholder, submitLabel, autoFocus, onSubmit }: ComposerProps) {
  const { mentionableUsers } = useWorkbook();
  const [value, setValue] = useState('');
  const [focused, setFocused] = useState(false);
  // Users picked from the @-autocomplete; reconciled against the body on submit.
  const [picked, setPicked] = useState<CommentAuthor[]>([]);
  const [mentionQuery, setMentionQuery] = useState<string | null>(null);
  const [highlightIdx, setHighlightIdx] = useState(0);
  const textareaRef = useRef<HTMLTextAreaElement>(null);

  const matches =
    mentionQuery !== null
      ? mentionableUsers
          .filter((u) => u.name.toLowerCase().includes(mentionQuery.toLowerCase()))
          .slice(0, 6)
      : [];
  const dropdownOpen = mentionQuery !== null && matches.length > 0;

  const submit = (): void => {
    const body = value.trim();
    if (!body) return;
    const mentions = [
      ...new Set(picked.filter((u) => body.includes('@' + u.name)).map((u) => u.id)),
    ];
    onSubmit(body, mentions);
    setValue('');
    setPicked([]);
    setMentionQuery(null);
  };

  // Track whether the caret sits inside an `@token` partial.
  const syncMentionQuery = (text: string, caret: number): void => {
    const m = text.slice(0, caret).match(/(?:^|\s)@(\w*)$/);
    setMentionQuery(m ? m[1] : null);
    setHighlightIdx(0);
  };

  const insertMention = (user: CommentAuthor): void => {
    const ta = textareaRef.current;
    const caret = ta?.selectionStart ?? value.length;
    const m = value.slice(0, caret).match(/@(\w*)$/);
    if (!m) return;
    const atIndex = caret - m[0].length;
    const next = value.slice(0, atIndex) + '@' + user.name + ' ' + value.slice(caret);
    setValue(next);
    setPicked((p) => (p.some((u) => u.id === user.id) ? p : [...p, user]));
    setMentionQuery(null);
    const newCaret = atIndex + user.name.length + 2;
    requestAnimationFrame(() => {
      ta?.focus();
      ta?.setSelectionRange(newCaret, newCaret);
    });
  };

  const handleKeyDown = (e: React.KeyboardEvent<HTMLTextAreaElement>): void => {
    if (dropdownOpen) {
      if (e.key === 'ArrowDown') {
        e.preventDefault();
        setHighlightIdx((i) => Math.min(i + 1, matches.length - 1));
        return;
      }
      if (e.key === 'ArrowUp') {
        e.preventDefault();
        setHighlightIdx((i) => Math.max(i - 1, 0));
        return;
      }
      if (e.key === 'Enter' || e.key === 'Tab') {
        e.preventDefault();
        insertMention(matches[highlightIdx]);
        return;
      }
      if (e.key === 'Escape') {
        e.preventDefault();
        setMentionQuery(null);
        return;
      }
    }
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      submit();
    }
  };

  return (
    <div>
      <div style={{ position: 'relative' }}>
        <textarea
          ref={textareaRef}
          value={value}
          autoFocus={autoFocus}
          placeholder={placeholder}
          onChange={(e) => {
            setValue(e.target.value);
            syncMentionQuery(e.target.value, e.target.selectionStart ?? e.target.value.length);
          }}
          onFocus={() => setFocused(true)}
          onBlur={() => setFocused(false)}
          onKeyDown={handleKeyDown}
          rows={2}
          style={{
            width: '100%',
            resize: 'vertical',
            minHeight: 38,
            padding: '8px 10px',
            borderRadius: '8px',
            border: `1px solid ${focused ? '#6366f1' : 'rgba(15, 23, 42, 0.12)'}`,
            boxShadow: focused ? '0 0 0 3px rgba(99, 102, 241, 0.12)' : 'none',
            outline: 'none',
            fontFamily: FONT,
            fontSize: '13px',
            color: '#1e293b',
            boxSizing: 'border-box',
            transition: 'border-color 0.15s ease, box-shadow 0.15s ease',
          }}
        />
        {dropdownOpen && (
          <div
            style={{
              position: 'absolute',
              top: 'calc(100% + 4px)',
              left: 0,
              right: 0,
              background: '#ffffff',
              borderRadius: '8px',
              border: '1px solid rgba(15, 23, 42, 0.1)',
              boxShadow: '0 8px 20px -6px rgba(15, 23, 42, 0.2)',
              zIndex: 50,
              padding: 4,
              maxHeight: 180,
              overflowY: 'auto',
            }}
          >
            {matches.map((u, i) => (
              <div
                key={u.id}
                // mousedown (not click) so the textarea keeps focus
                onMouseDown={(e) => {
                  e.preventDefault();
                  insertMention(u);
                }}
                onMouseEnter={() => setHighlightIdx(i)}
                style={{
                  display: 'flex',
                  alignItems: 'center',
                  gap: 8,
                  padding: '6px 8px',
                  borderRadius: '6px',
                  cursor: 'pointer',
                  background: i === highlightIdx ? 'rgba(99, 102, 241, 0.1)' : 'transparent',
                }}
              >
                <Avatar name={u.name} id={u.id} size={22} />
                <span style={{ fontSize: '13px', color: '#1e293b' }}>{u.name}</span>
              </div>
            ))}
          </div>
        )}
      </div>
      <div style={{ display: 'flex', justifyContent: 'flex-end', marginTop: 6 }}>
        <button
          onClick={submit}
          disabled={!value.trim()}
          style={{
            padding: '6px 14px',
            borderRadius: '8px',
            border: 'none',
            backgroundColor: value.trim() ? '#6366f1' : 'rgba(99, 102, 241, 0.4)',
            color: '#fff',
            fontSize: '13px',
            fontWeight: 600,
            fontFamily: FONT,
            cursor: value.trim() ? 'pointer' : 'not-allowed',
            transition: 'background-color 0.15s ease',
          }}
        >
          {submitLabel}
        </button>
      </div>
    </div>
  );
}

// ============================================================================
// Thread card
// ============================================================================

interface ThreadCardProps {
  thread: SheetCommentThread;
  /** Cell label (A1) shown as a badge — used in the all-threads list. */
  cellLabel?: string;
  onReply: (body: string, mentions: string[]) => void;
  onResolve: () => void;
  onReopen: () => void;
  onDeleteThread: () => void;
  onDeleteComment: (commentId: string) => void;
}

const iconButton: React.CSSProperties = {
  display: 'inline-flex',
  alignItems: 'center',
  justifyContent: 'center',
  width: 26,
  height: 26,
  border: 'none',
  borderRadius: '7px',
  backgroundColor: 'transparent',
  color: '#94a3b8',
  cursor: 'pointer',
  transition: 'background-color 0.12s ease, color 0.12s ease',
};

const cellBadge: React.CSSProperties = {
  fontSize: '11px',
  fontWeight: 600,
  color: '#6366f1',
  backgroundColor: 'rgba(99, 102, 241, 0.1)',
  padding: '2px 7px',
  borderRadius: '6px',
};

function CommentRow({
  comment,
  onDelete,
}: {
  comment: Comment;
  onDelete: () => void;
}) {
  const { mentionableUsers } = useWorkbook();
  const [hovered, setHovered] = useState(false);
  const mentionNames = (comment.mentions ?? [])
    .map((id) => mentionableUsers.find((u) => u.id === id)?.name)
    .filter((n): n is string => !!n);
  return (
    <div
      onMouseEnter={() => setHovered(true)}
      onMouseLeave={() => setHovered(false)}
      style={{ display: 'flex', gap: 8 }}
    >
      <Avatar name={comment.authorName} id={comment.authorId} />
      <div style={{ flex: 1, minWidth: 0 }}>
        <div style={{ display: 'flex', alignItems: 'baseline', gap: 6 }}>
          <span style={{ fontSize: '13px', fontWeight: 600, color: '#1e293b' }}>
            {comment.authorName}
          </span>
          <span style={{ fontSize: '11px', color: '#94a3b8' }}>
            {relativeTime(comment.createdAt)}
            {comment.editedAt ? ' (edited)' : ''}
          </span>
          <button
            onClick={onDelete}
            title="Delete comment"
            style={{
              ...iconButton,
              width: 22,
              height: 22,
              marginLeft: 'auto',
              opacity: hovered ? 1 : 0,
            }}
            onMouseEnter={(e) => {
              e.currentTarget.style.backgroundColor = '#fef2f2';
              e.currentTarget.style.color = '#ef4444';
            }}
            onMouseLeave={(e) => {
              e.currentTarget.style.backgroundColor = 'transparent';
              e.currentTarget.style.color = '#94a3b8';
            }}
          >
            <Trash2 size={13} strokeWidth={2} />
          </button>
        </div>
        <div
          style={{
            fontSize: '13px',
            color: '#334155',
            whiteSpace: 'pre-wrap',
            wordBreak: 'break-word',
            marginTop: 2,
          }}
        >
          {renderBody(comment.body, mentionNames)}
        </div>
      </div>
    </div>
  );
}

const ThreadCard = memo(function ThreadCard({
  thread,
  cellLabel,
  onReply,
  onResolve,
  onReopen,
  onDeleteThread,
  onDeleteComment,
}: ThreadCardProps) {
  const resolved = thread.status === 'resolved';
  return (
    <div
      style={{
        border: '1px solid rgba(15, 23, 42, 0.08)',
        borderRadius: '12px',
        background: resolved ? 'rgba(241, 245, 249, 0.6)' : '#ffffff',
        padding: '12px',
        opacity: resolved ? 0.8 : 1,
        boxShadow: '0 1px 2px rgba(15, 23, 42, 0.04)',
      }}
    >
      {cellLabel && (
        <div style={{ marginBottom: 8 }}>
          <span style={cellBadge}>{cellLabel}</span>
        </div>
      )}
      {resolved && (
        <div
          style={{
            display: 'flex',
            alignItems: 'center',
            gap: 5,
            fontSize: '11px',
            fontWeight: 600,
            color: '#10b981',
            marginBottom: 8,
          }}
        >
          <Check size={13} strokeWidth={2.5} />
          Resolved
        </div>
      )}

      <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
        {thread.comments.map((c) => (
          <CommentRow key={c.id} comment={c} onDelete={() => onDeleteComment(c.id)} />
        ))}
      </div>

      {!resolved && (
        <div style={{ marginTop: 10 }}>
          <Composer placeholder="Reply…" submitLabel="Reply" onSubmit={onReply} />
        </div>
      )}

      <div
        style={{
          display: 'flex',
          gap: 4,
          marginTop: 10,
          paddingTop: 8,
          borderTop: '1px solid rgba(15, 23, 42, 0.06)',
        }}
      >
        {resolved ? (
          <button
            onClick={onReopen}
            title="Reopen thread"
            style={{ ...iconButton, width: 'auto', padding: '0 10px', gap: 5, color: '#475569' }}
            onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = 'rgba(99,102,241,0.08)')}
            onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'transparent')}
          >
            <RotateCcw size={14} strokeWidth={2} />
            <span style={{ fontSize: '12px', fontWeight: 500 }}>Reopen</span>
          </button>
        ) : (
          <button
            onClick={onResolve}
            title="Resolve thread"
            style={{ ...iconButton, width: 'auto', padding: '0 10px', gap: 5, color: '#475569' }}
            onMouseEnter={(e) => {
              e.currentTarget.style.backgroundColor = 'rgba(16,185,129,0.1)';
              e.currentTarget.style.color = '#10b981';
            }}
            onMouseLeave={(e) => {
              e.currentTarget.style.backgroundColor = 'transparent';
              e.currentTarget.style.color = '#475569';
            }}
          >
            <Check size={14} strokeWidth={2.5} />
            <span style={{ fontSize: '12px', fontWeight: 500 }}>Resolve</span>
          </button>
        )}
        <button
          onClick={onDeleteThread}
          title="Delete thread"
          style={{ ...iconButton, marginLeft: 'auto' }}
          onMouseEnter={(e) => {
            e.currentTarget.style.backgroundColor = '#fef2f2';
            e.currentTarget.style.color = '#ef4444';
          }}
          onMouseLeave={(e) => {
            e.currentTarget.style.backgroundColor = 'transparent';
            e.currentTarget.style.color = '#94a3b8';
          }}
        >
          <Trash2 size={14} strokeWidth={2} />
        </button>
      </div>
    </div>
  );
});

// ============================================================================
// Panel
// ============================================================================

interface SegmentedProps<T extends string> {
  options: Array<{ value: T; label: string }>;
  value: T;
  onChange: (value: T) => void;
}

function Segmented<T extends string>({ options, value, onChange }: SegmentedProps<T>) {
  return (
    <div
      style={{
        display: 'flex',
        gap: 2,
        padding: 2,
        background: 'rgba(15, 23, 42, 0.05)',
        borderRadius: '9px',
      }}
    >
      {options.map((o) => {
        const active = o.value === value;
        return (
          <button
            key={o.value}
            onClick={() => onChange(o.value)}
            style={{
              flex: 1,
              padding: '5px 10px',
              border: 'none',
              borderRadius: '7px',
              background: active ? '#ffffff' : 'transparent',
              color: active ? '#6366f1' : '#64748b',
              fontWeight: active ? 600 : 500,
              fontSize: '12px',
              fontFamily: FONT,
              cursor: 'pointer',
              boxShadow: active ? '0 1px 2px rgba(15, 23, 42, 0.1)' : 'none',
              transition: 'all 0.12s ease',
            }}
          >
            {o.label}
          </button>
        );
      })}
    </div>
  );
}

interface CommentsPanelProps {
  activeCell: { row: number; col: number } | null;
  onClose: () => void;
}

type CommentTab = 'cell' | 'all';
type CommentFilter = 'open' | 'resolved' | 'all';

export const CommentsPanel = memo(function CommentsPanel({ activeCell, onClose }: CommentsPanelProps) {
  const comments = useComments();
  const [tab, setTab] = useState<CommentTab>('cell');
  const [filter, setFilter] = useState<CommentFilter>('open');

  const cellThreads = activeCell
    ? comments.getThreadsForCell(activeCell.row, activeCell.col)
    : [];
  const cellLabel = activeCell
    ? `${columnIndexToLabel(activeCell.col)}${activeCell.row + 1}`
    : '';

  const allThreads = [...comments.threads]
    .filter((t) => filter === 'all' || t.status === filter)
    .sort((a, b) => b.createdAt.localeCompare(a.createdAt));

  const threadCallbacks = (thread: SheetCommentThread) => ({
    onReply: (body: string, mentions: string[]) => comments.addReply(thread.id, body, mentions),
    onResolve: () => comments.resolveThread(thread.id),
    onReopen: () => comments.reopenThread(thread.id),
    onDeleteThread: () => comments.deleteThread(thread.id),
    onDeleteComment: (commentId: string) => comments.deleteComment(thread.id, commentId),
  });

  return (
    <div
      style={{
        height: '100%',
        display: 'flex',
        flexDirection: 'column',
        background: 'rgba(255, 255, 255, 0.98)',
        backdropFilter: 'blur(20px)',
        WebkitBackdropFilter: 'blur(20px)',
        borderLeft: '1px solid rgba(15, 23, 42, 0.1)',
        boxShadow: '-8px 0 24px -12px rgba(15, 23, 42, 0.15)',
        fontFamily: FONT,
      }}
    >
      {/* Header */}
      <div
        style={{
          display: 'flex',
          alignItems: 'center',
          gap: 8,
          padding: '12px 14px',
          borderBottom: '1px solid rgba(15, 23, 42, 0.08)',
        }}
      >
        <MessageSquare size={17} strokeWidth={2} color="#6366f1" />
        <span style={{ fontSize: '14px', fontWeight: 700, color: '#1e293b' }}>Comments</span>
        {tab === 'cell' && activeCell && <span style={cellBadge}>{cellLabel}</span>}
        <button
          onClick={onClose}
          title="Close"
          style={{ ...iconButton, marginLeft: 'auto' }}
          onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = 'rgba(15,23,42,0.06)')}
          onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'transparent')}
        >
          <X size={16} strokeWidth={2} />
        </button>
      </div>

      {/* Tab switcher */}
      <div style={{ padding: '10px 14px 0' }}>
        <Segmented
          value={tab}
          onChange={setTab}
          options={[
            { value: 'cell', label: 'This cell' },
            { value: 'all', label: 'All' },
          ]}
        />
      </div>

      {/* Body */}
      <div style={{ flex: 1, overflowY: 'auto', padding: '14px', display: 'flex', flexDirection: 'column', gap: 12 }}>
        {tab === 'cell' ? (
          !activeCell ? (
            <div style={{ color: '#94a3b8', fontSize: '13px', textAlign: 'center', marginTop: 24 }}>
              Select a cell to add a comment.
            </div>
          ) : (
            <>
              {cellThreads.map((thread) => (
                <ThreadCard key={thread.id} thread={thread} {...threadCallbacks(thread)} />
              ))}

              {/* New-thread composer */}
              <div
                style={{
                  border: '1px dashed rgba(15, 23, 42, 0.15)',
                  borderRadius: '12px',
                  padding: '12px',
                }}
              >
                <div style={{ fontSize: '12px', fontWeight: 600, color: '#64748b', marginBottom: 8 }}>
                  {cellThreads.length > 0 ? 'Add another comment' : `Comment on ${cellLabel}`}
                </div>
                <Composer
                  placeholder="Add a comment…  Use @ to mention"
                  submitLabel="Comment"
                  autoFocus={cellThreads.length === 0}
                  onSubmit={(body, mentions) =>
                    comments.addThreadAtCell(activeCell.row, activeCell.col, body, mentions)
                  }
                />
              </div>
            </>
          )
        ) : (
          <>
            <Segmented
              value={filter}
              onChange={setFilter}
              options={[
                { value: 'open', label: 'Open' },
                { value: 'resolved', label: 'Resolved' },
                { value: 'all', label: 'All' },
              ]}
            />
            {allThreads.length === 0 ? (
              <div style={{ color: '#94a3b8', fontSize: '13px', textAlign: 'center', marginTop: 24 }}>
                {filter === 'all' ? 'No comments yet.' : `No ${filter} comments.`}
              </div>
            ) : (
              allThreads.map((thread) => (
                <ThreadCard
                  key={thread.id}
                  thread={thread}
                  cellLabel={comments.cellLabelForThread(thread) ?? '(deleted cell)'}
                  {...threadCallbacks(thread)}
                />
              ))
            )}
          </>
        )}
      </div>
    </div>
  );
});
