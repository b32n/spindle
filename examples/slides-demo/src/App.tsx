import { useEffect, useMemo, useState } from 'react';
import { DeckProvider, SlidesEditor } from '@b32nio/spindle-slides-react';
import { DeckImpl } from '@b32nio/spindle-slides-core';
import type { SlidesCommentEvent } from '@b32nio/spindle-slides-core';
import { InMemoryProvider, type CollabIdentity, type CollabStatus } from '@b32nio/spindle-shared';
import { WebSocketProvider } from '@b32nio/spindle-transport-websocket';
import { buildSampleDeck } from './sampleDeck';

const DEMO_USERS = [
  { id: 'alice', name: 'Alice' },
  { id: 'bob', name: 'Bob' },
];

function logCommentEvent(event: SlidesCommentEvent): void {
  console.log('[comment event]', event.type, event);
}

function makeDeck(id = 'demo-deck'): DeckImpl {
  const deck = new DeckImpl(id, 'Spindle Slides — Demo');
  deck.setData({ ...buildSampleDeck(), id });
  return deck;
}

// ── ?ws= config ───────────────────────────────────────────────────────────────

const COLORS = ['#ff6b6b', '#4ecdc4', '#ffd93d', '#6c5ce7', '#54a0ff'];
const NAMES = ['Quick Otter', 'Calm Lynx', 'Brave Owl', 'Sharp Fox', 'Eager Hare'];

interface WsConfig {
  url: string;
  roomId: string;
  identity: CollabIdentity;
}

function parseWsConfig(): WsConfig | null {
  const params = new URLSearchParams(window.location.search);
  const url = params.get('ws');
  if (!url) return null;
  const roomId = params.get('room') ?? 'slides-demo';
  const pick = <T,>(a: T[]) => a[Math.floor(Math.random() * a.length)];
  return {
    url,
    roomId,
    identity: {
      userId: `user_${Math.random().toString(36).slice(2, 8)}`,
      displayName: params.get('user') ?? pick(NAMES),
      color: params.get('color') ?? pick(COLORS),
    },
  };
}

function StatusDot({ status }: { status: CollabStatus }) {
  const color = status === 'connected' ? '#2ecc71' : status === 'connecting' ? '#f1c40f' : '#e74c3c';
  return <span style={{ width: 9, height: 9, borderRadius: '50%', background: color, display: 'inline-block', boxShadow: '0 0 0 2px rgba(255,255,255,0.4)' }} />;
}

// ── Two in-memory peers, side by side ─────────────────────────────────────────

function CollabDemo() {
  const [decks, setDecks] = useState<[DeckImpl, DeckImpl] | null>(null);

  useEffect(() => {
    let cancelled = false;
    const a = makeDeck('collab-deck');
    const b = makeDeck('collab-deck');
    const ids: CollabIdentity[] = [
      { userId: 'user_a', displayName: 'Alice', color: '#ff6b6b' },
      { userId: 'user_b', displayName: 'Bob', color: '#4ecdc4' },
    ];
    Promise.all([
      a.attachCollab(new InMemoryProvider(), ids[0], { roomId: 'collab-deck' }),
      b.attachCollab(new InMemoryProvider(), ids[1], { roomId: 'collab-deck' }),
    ]).then(() => {
      if (cancelled) { a.detachCollab(); b.detachCollab(); return; }
      setDecks([a, b]);
    });
    return () => { cancelled = true; try { a.detachCollab(); b.detachCollab(); } catch { /* ignore */ } };
  }, []);

  if (!decks) return <div style={{ padding: 16 }}>Connecting collab peers…</div>;

  return (
    <div style={{ display: 'flex', height: '100%', width: '100%' }}>
      {decks.map((deck, i) => (
        <div key={i} style={{ flex: 1, minWidth: 0, height: '100%', display: 'flex', flexDirection: 'column', borderRight: i === 0 ? '1px solid #e0e0e0' : 'none' }}>
          <div style={{ padding: '4px 12px', background: i === 0 ? '#ff6b6b' : '#4ecdc4', color: '#fff', fontWeight: 500 }}>{DEMO_USERS[i].name}</div>
          <div style={{ flex: 1, minHeight: 0 }}>
            <DeckProvider deck={deck} currentUser={DEMO_USERS[i]} mentionableUsers={DEMO_USERS} onCommentEvent={logCommentEvent}>
              <SlidesEditor />
            </DeckProvider>
          </div>
        </div>
      ))}
    </div>
  );
}

// ── Cross-tab WebSocket peer ───────────────────────────────────────────────────

function WsDemo({ config }: { config: WsConfig }) {
  const [deck, setDeck] = useState<DeckImpl | null>(null);
  const [status, setStatus] = useState<CollabStatus>('connecting');

  useEffect(() => {
    let cancelled = false;
    const d = makeDeck('collab-deck');
    const provider = new WebSocketProvider({ url: config.url });
    provider.onStatusChange?.((s) => { if (!cancelled) setStatus(s); });
    d.attachCollab(provider, config.identity, { roomId: config.roomId, persistenceKey: `slides-demo:${config.roomId}` })
      .then(() => { if (cancelled) { d.detachCollab(); return; } setDeck(d); });
    return () => { cancelled = true; try { d.detachCollab(); } catch { /* ignore */ } };
  }, [config.url, config.roomId, config.identity]);

  if (!deck) return <div style={{ padding: 16 }}>Connecting to {config.url}…</div>;
  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100%' }}>
      <div style={{ padding: '4px 12px', background: config.identity.color, color: '#fff', fontWeight: 500, display: 'flex', alignItems: 'center', gap: 8 }}>
        <StatusDot status={status} />
        <span>{config.identity.displayName} · room {config.roomId} · {status}</span>
      </div>
      <div style={{ flex: 1, minHeight: 0 }}>
        <DeckProvider deck={deck} currentUser={{ id: config.identity.userId, name: config.identity.displayName }} mentionableUsers={DEMO_USERS} onCommentEvent={logCommentEvent}>
          <SlidesEditor />
        </DeckProvider>
      </div>
    </div>
  );
}

export default function App() {
  const wsConfig = useMemo(parseWsConfig, []);
  const [twoPane, setTwoPane] = useState(false);
  const singleDeck = useMemo(() => makeDeck(), []);

  if (wsConfig) return <WsDemo config={wsConfig} />;
  if (twoPane) return <CollabDemo />;

  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100%' }}>
      <div style={{ padding: '6px 12px', borderBottom: '1px solid #e2e4e8', fontSize: 13, display: 'flex', gap: 12, alignItems: 'center' }}>
        <strong>Spindle Slides demo</strong>
        <button onClick={() => setTwoPane(true)} style={{ marginLeft: 'auto', fontSize: 12, padding: '4px 10px', borderRadius: 5, border: '1px solid #d5d9e0', background: '#fff', cursor: 'pointer' }}>
          Two-pane collab →
        </button>
        <span style={{ color: '#8a93a2', fontSize: 12 }}>or add <code>?ws=ws://localhost:1234</code> for cross-tab</span>
      </div>
      <div style={{ flex: 1, minHeight: 0 }}>
        <DeckProvider deck={singleDeck} currentUser={DEMO_USERS[0]} mentionableUsers={DEMO_USERS} onCommentEvent={logCommentEvent}>
          <SlidesEditor />
        </DeckProvider>
      </div>
    </div>
  );
}
