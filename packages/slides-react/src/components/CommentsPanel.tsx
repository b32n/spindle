// CommentsPanel — a side panel listing comment threads. Threads whose element
// still exists are grouped by slide; threads whose element was deleted appear
// under "No longer attached" (kept, never silently dropped). Clicking a thread
// selects its element and switches to its slide. New comments attach to the
// single selected element.

import React, { useState } from 'react';
import { X, Check, CornerUpLeft, Trash2 } from 'lucide-react';
import type { SlidesCommentThread } from '@b32nio/spindle-slides-core';
import { useDeck, useSelection, useActiveSlideId } from '../hooks';
import { useComments } from '../hooks/useComments';

function initials(name: string): string {
  return name.split(/\s+/).map((p) => p[0]).slice(0, 2).join('').toUpperCase();
}

function ThreadCard({ thread, orphaned, onSelect }: { thread: SlidesCommentThread; orphaned: boolean; onSelect: () => void }): React.ReactElement {
  const c = useComments();
  const [reply, setReply] = useState('');

  return (
    <div
      onClick={onSelect}
      style={{ border: '1px solid #e2e4e8', borderRadius: 8, padding: 10, marginBottom: 10, background: thread.status === 'resolved' ? '#f4f6f8' : '#fff', cursor: 'pointer', opacity: orphaned ? 0.75 : 1 }}
    >
      {thread.comments.map((comment) => (
        <div key={comment.id} style={{ display: 'flex', gap: 8, marginBottom: 8 }}>
          <div style={{ flex: 'none', width: 24, height: 24, borderRadius: '50%', background: '#2d7ff9', color: '#fff', fontSize: 10, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
            {initials(comment.authorName)}
          </div>
          <div style={{ flex: 1, minWidth: 0 }}>
            <div style={{ fontSize: 12, fontWeight: 600, color: '#2b3440' }}>{comment.authorName}</div>
            <div style={{ fontSize: 13, color: '#3e4c59', whiteSpace: 'pre-wrap', wordBreak: 'break-word' }}>{comment.body}</div>
          </div>
        </div>
      ))}
      <div style={{ display: 'flex', gap: 6, alignItems: 'center', marginTop: 6 }} onClick={(e) => e.stopPropagation()}>
        <input
          value={reply}
          onChange={(e) => setReply(e.target.value)}
          onKeyDown={(e) => { if (e.key === 'Enter' && reply.trim()) { c.addReply(thread.id, reply.trim()); setReply(''); } }}
          placeholder="Reply…"
          style={{ flex: 1, minWidth: 0, border: '1px solid #e2e4e8', borderRadius: 5, padding: '5px 8px', fontSize: 12 }}
        />
        <button title="Reply" onClick={() => { if (reply.trim()) { c.addReply(thread.id, reply.trim()); setReply(''); } }} style={iconBtn}><CornerUpLeft size={13} /></button>
        <button title={thread.status === 'resolved' ? 'Reopen' : 'Resolve'} onClick={() => (thread.status === 'resolved' ? c.reopen(thread.id) : c.resolve(thread.id))} style={iconBtn}><Check size={13} color={thread.status === 'resolved' ? '#16a34a' : '#3e4c59'} /></button>
        <button title="Delete thread" onClick={() => c.deleteThread(thread.id)} style={iconBtn}><Trash2 size={13} /></button>
      </div>
    </div>
  );
}

const iconBtn: React.CSSProperties = { display: 'inline-flex', alignItems: 'center', justifyContent: 'center', width: 26, height: 26, border: '1px solid #d5d9e0', borderRadius: 5, background: '#fff', cursor: 'pointer', flex: 'none' };

export function CommentsPanel({ onClose }: { onClose: () => void }): React.ReactElement {
  const deck = useDeck();
  const c = useComments();
  const selection = useSelection();
  const activeSlideId = useActiveSlideId();
  const [draft, setDraft] = useState('');

  const selectedId = selection.elementIds.length === 1 ? selection.elementIds[0] : null;
  const attached = c.threads.filter((t) => !c.isOrphaned(t.id));
  const orphaned = c.threads.filter((t) => c.isOrphaned(t.id));

  const selectThreadElement = (t: SlidesCommentThread) => {
    if (deck.getElement(t.anchor.elementId)) {
      deck.setActiveSlide(t.anchor.slideId);
      deck.setSelection({ slideId: t.anchor.slideId, elementIds: [t.anchor.elementId] });
    }
  };

  return (
    <aside style={{ width: 300, flex: 'none', borderLeft: '1px solid #e2e4e8', background: '#fbfcfe', display: 'flex', flexDirection: 'column', minHeight: 0 }}>
      <header style={{ display: 'flex', alignItems: 'center', padding: '10px 14px', borderBottom: '1px solid #e2e4e8' }}>
        <strong style={{ fontSize: 14 }}>Comments</strong>
        <button onClick={onClose} style={{ ...iconBtn, marginLeft: 'auto', border: 'none', background: 'transparent' }} title="Close"><X size={16} /></button>
      </header>
      <div style={{ flex: 1, overflowY: 'auto', padding: 12 }}>
        {selectedId ? (
          <div style={{ marginBottom: 14 }}>
            <textarea
              value={draft}
              onChange={(e) => setDraft(e.target.value)}
              placeholder={c.currentUser ? 'Comment on the selected element…' : 'Sign in to comment'}
              disabled={!c.currentUser}
              style={{ width: '100%', minHeight: 54, resize: 'vertical', border: '1px solid #e2e4e8', borderRadius: 6, padding: 8, fontSize: 13, fontFamily: 'inherit', boxSizing: 'border-box' }}
            />
            <button
              disabled={!draft.trim() || !c.currentUser}
              onClick={() => { c.addThread({ slideId: activeSlideId, elementId: selectedId }, draft.trim()); setDraft(''); }}
              style={{ marginTop: 6, border: 'none', background: draft.trim() && c.currentUser ? '#2d7ff9' : '#c4cad3', color: '#fff', borderRadius: 5, padding: '6px 12px', fontSize: 13, cursor: draft.trim() && c.currentUser ? 'pointer' : 'default' }}
            >
              Comment
            </button>
          </div>
        ) : (
          <p style={{ fontSize: 12, color: '#8a93a2', marginBottom: 14 }}>Select an element to add a comment.</p>
        )}

        {attached.length === 0 && orphaned.length === 0 ? (
          <p style={{ fontSize: 13, color: '#8a93a2' }}>No comments yet.</p>
        ) : null}

        {attached.map((t) => (
          <ThreadCard key={t.id} thread={t} orphaned={false} onSelect={() => selectThreadElement(t)} />
        ))}

        {orphaned.length > 0 && (
          <>
            <div style={{ fontSize: 11, textTransform: 'uppercase', letterSpacing: 0.5, color: '#8a93a2', margin: '12px 0 8px' }}>No longer attached</div>
            {orphaned.map((t) => (
              <ThreadCard key={t.id} thread={t} orphaned onSelect={() => {}} />
            ))}
          </>
        )}
      </div>
    </aside>
  );
}
