import { useState, useEffect } from 'react';
import { WorkbookProvider, WorkbookCanvas } from '@b32nio/spindle-sheets-react';
import { WorkbookImpl } from '@b32nio/spindle-sheets-core';
import type { CommentAuthor, SheetCommentEvent } from '@b32nio/spindle-sheets-core';
import { InMemoryProvider, type CollabIdentity, type CollabStatus } from '@b32nio/spindle-shared';
import { WebSocketProvider } from '@b32nio/spindle-transport-websocket';
import './App.css';

// Sample directory of users that can be @-mentioned in comments.
const DEMO_USERS: CommentAuthor[] = [
  { id: 'alice', name: 'Alice' },
  { id: 'bob', name: 'Bob' },
  { id: 'carol', name: 'Carol Diaz' },
  { id: 'dave', name: 'Dave Kim' },
];

// Stand-in for a host app reacting to comment activity (notifications, etc.).
function logCommentEvent(event: SheetCommentEvent): void {
  console.log('[comment event]', event.type, event);
}

// ============================================================================
// Cross-tab WebSocket demo helpers
// ============================================================================

const COLORS = ['#ff6b6b', '#4ecdc4', '#ffd93d', '#6c5ce7', '#a8e6cf', '#ff8c42', '#54a0ff', '#48dbfb'];
const ADJECTIVES = ['Quick', 'Calm', 'Brave', 'Sharp', 'Bright', 'Eager', 'Lucky', 'Witty'];
const NOUNS = ['Otter', 'Lynx', 'Owl', 'Fox', 'Heron', 'Bear', 'Hare', 'Crane'];

interface WsConfig {
  url: string;
  roomId: string;
  identity: CollabIdentity;
}

function parseWsConfig(): WsConfig | null {
  const params = new URLSearchParams(window.location.search);
  const url = params.get('ws');
  if (!url) return null;
  const roomId = params.get('room') ?? 'sheets-demo';
  const userName =
    params.get('user') ??
    `${ADJECTIVES[Math.floor(Math.random() * ADJECTIVES.length)]} ${NOUNS[Math.floor(Math.random() * NOUNS.length)]}`;
  const userColor = params.get('color') ?? COLORS[Math.floor(Math.random() * COLORS.length)];
  return {
    url,
    roomId,
    identity: {
      userId: `user_${Math.random().toString(36).slice(2, 8)}`,
      displayName: userName,
      color: userColor,
    },
  };
}

// Sample workbook data lifted out so both single-editor and collab demo
// can hydrate from the same starting state.
const sampleWorkbookData = (() => {
  const data = {
      id: 'demo-workbook',
      name: 'Demo Spreadsheet',
      activeSheetId: 'sheet_1',
      defaultRowHeight: 20,
      defaultColWidth: 100,
      stylePool: {},
      formatPool: {},
      sheets: [
        {
          id: 'sheet_1',
          name: 'Sheet1',
          cells: [
            { key: '0:0', cell: { value: 'Product' } },
            { key: '0:1', cell: { value: 'Price' } },
            { key: '0:2', cell: { value: 'Quantity' } },
            { key: '0:3', cell: { value: 'Total' } },
            { key: '1:0', cell: { value: 'Widget' } },
            { key: '1:1', cell: { value: 10 } },
            { key: '1:2', cell: { value: 5 } },
            { key: '1:3', cell: { value: 50 } },
            { key: '2:0', cell: { value: 'Gadget' } },
            { key: '2:1', cell: { value: 20 } },
            { key: '2:2', cell: { value: 3 } },
            { key: '2:3', cell: { value: 60 } },
            { key: '3:0', cell: { value: 'Thing' } },
            { key: '3:1', cell: { value: 15 } },
            { key: '3:2', cell: { value: 2 } },
            { key: '3:3', cell: { value: 30 } },
          ],
          config: {
            defaultRowHeight: 20,
            defaultColWidth: 100,
            showGridLines: true,
          },
          rowCount: 1000,
          colCount: 100,
        },
      ],
      selection: {
        ranges: [],
        activeCell: { row: 0, col: 0 },
      },
    };
  return data;
})();

function makeWorkbookFromSample(id = 'demo-workbook', name = 'Demo Spreadsheet'): WorkbookImpl {
  const wb = new WorkbookImpl(id, name);
  wb.setData({ ...sampleWorkbookData, id, name });
  // Demonstrate some formatting after load.
  wb.setCell(undefined, 1, 3, {
    ...wb.getCell(undefined, 1, 3),
    formatId: wb.getFormatPool().getOrCreate({ type: 'currency', currencyCode: 'USD' }),
  });
  wb.setCell(undefined, 2, 3, {
    ...wb.getCell(undefined, 2, 3),
    formatId: wb.getFormatPool().getOrCreate({ type: 'currency', currencyCode: 'USD' }),
  });
  return wb;
}

// Phase 2b.6 smoke-test panel. Two WorkbookImpls sharing an InMemoryProvider
// room; edits in one pane should converge into the other.
function CollabDemo({ width, height }: { width: number; height: number }) {
  const [workbooks, setWorkbooks] = useState<[WorkbookImpl, WorkbookImpl] | null>(null);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    let cancelled = false;
    let wbA: WorkbookImpl | null = null;
    let wbB: WorkbookImpl | null = null;

    (async () => {
      try {
        wbA = makeWorkbookFromSample('collab-workbook', 'Shared Workbook');
        wbB = makeWorkbookFromSample('collab-workbook', 'Shared Workbook');

        const providerA = new InMemoryProvider();
        const providerB = new InMemoryProvider();

        const identityA: CollabIdentity = {
          userId: 'user_a',
          displayName: 'Alice',
          color: '#ff6b6b',
        };
        const identityB: CollabIdentity = {
          userId: 'user_b',
          displayName: 'Bob',
          color: '#4ecdc4',
        };

        await Promise.all([
          wbA.attachCollab(providerA, identityA),
          wbB.attachCollab(providerB, identityB),
        ]);
        if (cancelled) {
          wbA.detachCollab();
          wbB.detachCollab();
          return;
        }
        setWorkbooks([wbA, wbB]);
      } catch (e) {
        if (!cancelled) setError(e instanceof Error ? e.message : String(e));
      }
    })();
    return () => {
      cancelled = true;
      try { wbA?.detachCollab(); } catch { /* ignore */ }
      try { wbB?.detachCollab(); } catch { /* ignore */ }
    };
  }, []);

  if (error) return <div style={{ padding: 16, color: 'crimson' }}>Error: {error}</div>;
  if (!workbooks) return <div style={{ padding: 16 }}>Connecting collab peers…</div>;

  const paneWidth = Math.floor(width / 2) - 1;
  const paneHeight = height - 28; // peer label strip
  const paneStyle: React.CSSProperties = {
    flex: 1,
    minWidth: 0,
    height: '100%',
    display: 'flex',
    flexDirection: 'column',
    borderRight: '1px solid #e0e0e0',
  };

  return (
    <div style={{ display: 'flex', height: '100%', width: '100%' }}>
      <div style={paneStyle}>
        <div style={{ padding: '4px 12px', background: '#ff6b6b', color: 'white', fontWeight: 500 }}>Alice</div>
        <div style={{ flex: 1, minHeight: 0 }}>
          <WorkbookProvider
            workbook={workbooks[0]}
            currentUser={DEMO_USERS[0]}
            mentionableUsers={DEMO_USERS}
            onCommentEvent={logCommentEvent}
          >
            <WorkbookCanvas width={paneWidth} height={paneHeight} />
          </WorkbookProvider>
        </div>
      </div>
      <div style={{ ...paneStyle, borderRight: 'none' }}>
        <div style={{ padding: '4px 12px', background: '#4ecdc4', color: 'white', fontWeight: 500 }}>Bob</div>
        <div style={{ flex: 1, minHeight: 0 }}>
          <WorkbookProvider
            workbook={workbooks[1]}
            currentUser={DEMO_USERS[1]}
            mentionableUsers={DEMO_USERS}
            onCommentEvent={logCommentEvent}
          >
            <WorkbookCanvas width={paneWidth} height={paneHeight} />
          </WorkbookProvider>
        </div>
      </div>
    </div>
  );
}

/**
 * Single-tab WebSocket collab mode. Activated by ?ws=ws://host/path in
 * the URL. Each open tab is one peer; multiple tabs (or browsers, or
 * machines) on the same ?ws=…&room=… all share state via the relay
 * server at /examples/collab-server.
 */
function WsDemo({
  width,
  height,
  config,
}: {
  width: number;
  height: number;
  config: WsConfig;
}) {
  const [wb, setWb] = useState<WorkbookImpl | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [status, setStatus] = useState<CollabStatus>('connecting');

  useEffect(() => {
    let cancelled = false;
    let provider: WebSocketProvider | null = null;
    let workbook: WorkbookImpl | null = null;
    (async () => {
      try {
        workbook = makeWorkbookFromSample('collab-workbook', 'Shared Workbook');
        provider = new WebSocketProvider({ url: config.url });
        provider.onStatusChange((s) => {
          if (!cancelled) setStatus(s);
        });
        await workbook.attachCollab(provider, config.identity, {
          roomId: config.roomId,
          persistenceKey: `sheets-demo:${config.roomId}`,
        });
        if (cancelled) {
          workbook.detachCollab();
          return;
        }
        setWb(workbook);
      } catch (e) {
        if (!cancelled) setError(e instanceof Error ? e.message : String(e));
      }
    })();
    return () => {
      cancelled = true;
      try { workbook?.detachCollab(); } catch { /* ignore */ }
    };
  }, [config.url, config.roomId, config.identity]);

  if (error) return <div style={{ padding: 16, color: 'crimson' }}>WS error: {error}</div>;
  if (!wb) return <div style={{ padding: 16 }}>Connecting to {config.url}…</div>;
  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100%' }}>
      <div
        style={{
          padding: '4px 12px',
          background: config.identity.color,
          color: 'white',
          fontWeight: 500,
          display: 'flex',
          alignItems: 'center',
          gap: 8,
        }}
      >
        <StatusDot status={status} />
        <span>
          {config.identity.displayName} · room {config.roomId} · {status}
        </span>
      </div>
      <div style={{ flex: 1, minHeight: 0 }}>
        <WorkbookProvider
          workbook={wb}
          currentUser={{ id: config.identity.userId, name: config.identity.displayName }}
          mentionableUsers={DEMO_USERS}
          onCommentEvent={logCommentEvent}
        >
          <WorkbookCanvas width={width} height={height - 28} />
        </WorkbookProvider>
      </div>
    </div>
  );
}

function StatusDot({ status }: { status: CollabStatus }) {
  const color =
    status === 'connected' ? '#2ecc71' : status === 'connecting' ? '#f1c40f' : '#e74c3c';
  return (
    <span
      style={{
        width: 9,
        height: 9,
        borderRadius: '50%',
        background: color,
        display: 'inline-block',
        boxShadow: '0 0 0 2px rgba(255,255,255,0.4)',
      }}
    />
  );
}

function App() {
  const [collabMode, setCollabMode] = useState(false);
  const [readOnly, setReadOnly] = useState(false);
  const [wsConfig] = useState<WsConfig | null>(() => parseWsConfig());
  const [workbook] = useState(() => makeWorkbookFromSample());

  const [dimensions, setDimensions] = useState({
    width: window.innerWidth,
    height: window.innerHeight - 100,
  });

  useEffect(() => {
    const handleResize = () => {
      setDimensions({
        width: window.innerWidth,
        height: window.innerHeight - 100,
      });
    };

    window.addEventListener('resize', handleResize);
    return () => window.removeEventListener('resize', handleResize);
  }, []);

  const handleSave = () => {
    const data = workbook.getData();
    console.log('workbook data', data);
  };

  return (
    <div className="app">
      <header className="app-header">
        <h1>Spindle Sheets Demo</h1>
        <p>Standalone demo of Spindle Sheets library - No backend required</p>
        <button onClick={handleSave} style={{ marginTop: '10px', padding: '8px 16px' }}>
          Save Workbook Data
        </button>
        <button
          onClick={() => setCollabMode((v) => !v)}
          style={{ marginTop: '10px', marginLeft: '8px', padding: '8px 16px' }}
        >
          {collabMode ? 'Single editor' : 'Collab demo'}
        </button>
        {!collabMode && !wsConfig && (
          <button
            onClick={() => setReadOnly((v) => !v)}
            style={{
              marginTop: '10px',
              marginLeft: '8px',
              padding: '8px 16px',
              background: readOnly ? '#fde68a' : undefined,
            }}
          >
            {readOnly ? 'Read-only: ON' : 'Read-only: OFF'}
          </button>
        )}
      </header>
      <main className="app-main">
        {wsConfig ? (
          <WsDemo width={dimensions.width} height={dimensions.height} config={wsConfig} />
        ) : collabMode ? (
          <CollabDemo width={dimensions.width} height={dimensions.height} />
        ) : (
          <WorkbookProvider
            workbook={workbook}
            currentUser={DEMO_USERS[0]}
            mentionableUsers={DEMO_USERS}
            onCommentEvent={logCommentEvent}
          >
            <WorkbookCanvas width={dimensions.width} height={dimensions.height} readOnly={readOnly} />
          </WorkbookProvider>
        )}
      </main>
    </div>
  );
}

export default App;

