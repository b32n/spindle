import { useState, useEffect } from 'react';
import { DocumentProvider, DocumentEditor } from '@b32nio/spindle-docs-react';
import {
  DocumentImpl,
  type DocumentData,
  type HeaderFooterContent,
  type CommentAuthor,
  type DocsCommentEvent,
} from '@b32nio/spindle-docs-core';
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
function logCommentEvent(event: DocsCommentEvent): void {
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
  const roomId = params.get('room') ?? 'docs-demo';
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

// Sample header configuration with document title
const sampleHeader: HeaderFooterContent = {
  blocks: [
    {
      type: 'paragraph',
      alignment: 'left',
      content: [
        { type: 'dynamicField', fieldType: 'title' },
      ],
    },
  ],
  differentFirstPage: true,
  firstPageBlocks: [], // No header on first page
};

// Sample footer configuration with page numbers
const sampleFooter: HeaderFooterContent = {
  blocks: [
    {
      type: 'paragraph',
      alignment: 'center',
      content: [
        { type: 'text', text: 'Page ' },
        { type: 'dynamicField', fieldType: 'pageNumber' },
        { type: 'text', text: ' of ' },
        { type: 'dynamicField', fieldType: 'totalPages' },
      ],
    },
  ],
};

// Sample document data in our JSON format
// This demonstrates how to load a document from a backend/database
const sampleDocumentData: DocumentData = {
  id: 'demo-document',
  title: 'Demo Document',
  defaultPageConfig: {
    size: { w: 816, h: 1056 },
    margins: { top: 96, right: 96, bottom: 96, left: 96, header: 48, footer: 48 },
    orientation: 'portrait',
  },
  // Style pool for efficient storage - same styles are deduplicated
  textStylePool: {
    'style_bold': { bold: true },
    'style_italic': { italic: true },
    'style_code': { fontFamily: 'monospace', backgroundColor: '#f5f5f5' },
    'style_highlight': { backgroundColor: '#fff3cd' },
    'style_link_style': { color: '#0066cc', underline: true },
  },
  paragraphStylePool: {},
  sections: [
    {
      id: 'section_main',
      pageConfig: {
        size: { w: 816, h: 1056 },
        margins: { top: 96, right: 96, bottom: 96, left: 96, header: 48, footer: 48 },
        orientation: 'portrait',
      },
      // Header and footer are part of the section data
      header: sampleHeader,
      footer: sampleFooter,
      blocks: [
        // Heading 1 - centered
        {
          id: 'block_1',
          type: 'heading',
          level: 1,
          alignment: 'center',
          content: [{ type: 'text', text: 'Welcome to Pagent Docs' }],
        },
        // Intro paragraph with mixed styling
        {
          id: 'block_2',
          type: 'paragraph',
          content: [
            { type: 'text', text: 'This is a demonstration of the ' },
            { type: 'text', text: 'Spindle Docs', styleId: 'style_code' },
            { type: 'text', text: ' document editor library. It provides a ' },
            { type: 'text', text: 'rich text editing', styleId: 'style_bold' },
            { type: 'text', text: ' experience similar to Google Docs.' },
          ],
        },
        // Heading 2
        {
          id: 'block_3',
          type: 'heading',
          level: 2,
          content: [{ type: 'text', text: 'Features' }],
        },
        {
          id: 'block_4',
          type: 'paragraph',
          content: [{ type: 'text', text: 'The editor includes support for:' }],
        },
        // Bullet list items
        {
          id: 'block_5',
          type: 'list-item',
          listType: 'bullet',
          level: 0,
          content: [
            { type: 'text', text: 'Rich text formatting (' },
            { type: 'text', text: 'bold', styleId: 'style_bold' },
            { type: 'text', text: ', ' },
            { type: 'text', text: 'italic', styleId: 'style_italic' },
            { type: 'text', text: ', underline)' },
          ],
        },
        {
          id: 'block_6',
          type: 'list-item',
          listType: 'bullet',
          level: 0,
          content: [{ type: 'text', text: 'Multiple heading levels' }],
        },
        {
          id: 'block_7',
          type: 'list-item',
          listType: 'bullet',
          level: 0,
          content: [{ type: 'text', text: 'Bulleted and numbered lists' }],
        },
        {
          id: 'block_8',
          type: 'list-item',
          listType: 'bullet',
          level: 0,
          content: [{ type: 'text', text: 'Tables with customizable columns' }],
        },
        {
          id: 'block_9',
          type: 'list-item',
          listType: 'bullet',
          level: 0,
          content: [{ type: 'text', text: 'Images with captions' }],
        },
        {
          id: 'block_10',
          type: 'list-item',
          listType: 'bullet',
          level: 0,
          content: [{ type: 'text', text: 'Page setup with margins' }],
        },
        // Page Configuration section
        {
          id: 'block_11',
          type: 'heading',
          level: 2,
          content: [{ type: 'text', text: 'Page Configuration' }],
        },
        {
          id: 'block_12',
          type: 'paragraph',
          content: [
            { type: 'text', text: 'Click the page setup button (📄) in the toolbar to configure page size, orientation, and margins. ' },
            { type: 'text', text: 'The ruler at the top', styleId: 'style_highlight' },
            { type: 'text', text: ' allows you to adjust margins by dragging the handles.' },
          ],
        },
        // Sample Table section
        {
          id: 'block_13',
          type: 'heading',
          level: 3,
          content: [{ type: 'text', text: 'Sample Table' }],
        },
        {
          id: 'block_14',
          type: 'table',
          rows: [
            {
              id: 'row_1',
              cells: [
                { id: 'cell_1_1', content: [{ type: 'text', text: 'Feature', styleId: 'style_bold' }] },
                { id: 'cell_1_2', content: [{ type: 'text', text: 'Status', styleId: 'style_bold' }] },
                { id: 'cell_1_3', content: [{ type: 'text', text: 'Notes', styleId: 'style_bold' }] },
              ],
            },
            {
              id: 'row_2',
              cells: [
                { id: 'cell_2_1', content: [{ type: 'text', text: 'Text Formatting' }] },
                { id: 'cell_2_2', content: [{ type: 'text', text: '✅ Complete' }] },
                { id: 'cell_2_3', content: [{ type: 'text', text: 'Bold, italic, underline' }] },
              ],
            },
            {
              id: 'row_3',
              cells: [
                { id: 'cell_3_1', content: [{ type: 'text', text: 'Page Layout' }] },
                { id: 'cell_3_2', content: [{ type: 'text', text: '✅ Complete' }] },
                { id: 'cell_3_3', content: [{ type: 'text', text: 'Multiple page sizes' }] },
              ],
            },
          ],
        },
        {
          id: 'block_15',
          type: 'paragraph',
          content: [
            { type: 'text', text: 'The table above demonstrates basic table support. Tables can have custom column widths and cell content.' },
          ],
        },
        // Getting Started section
        {
          id: 'block_16',
          type: 'heading',
          level: 2,
          content: [{ type: 'text', text: 'Getting Started' }],
        },
        {
          id: 'block_17',
          type: 'paragraph',
          content: [{ type: 'text', text: 'To use this library in your own project:' }],
        },
        // Numbered list
        {
          id: 'block_18',
          type: 'list-item',
          listType: 'numbered',
          level: 0,
          content: [
            { type: 'text', text: 'Install ' },
            { type: 'text', text: '@b32nio/spindle-docs-core', styleId: 'style_code' },
            { type: 'text', text: ' and ' },
            { type: 'text', text: '@b32nio/spindle-docs-react', styleId: 'style_code' },
          ],
        },
        {
          id: 'block_19',
          type: 'list-item',
          listType: 'numbered',
          level: 0,
          content: [
            { type: 'text', text: 'Create a ' },
            { type: 'text', text: 'DocumentImpl', styleId: 'style_code' },
            { type: 'text', text: ' instance' },
          ],
        },
        {
          id: 'block_20',
          type: 'list-item',
          listType: 'numbered',
          level: 0,
          content: [
            { type: 'text', text: 'Wrap your app with ' },
            { type: 'text', text: 'DocumentProvider', styleId: 'style_code' },
          ],
        },
        {
          id: 'block_21',
          type: 'list-item',
          listType: 'numbered',
          level: 0,
          content: [
            { type: 'text', text: 'Use the ' },
            { type: 'text', text: 'DocumentEditor', styleId: 'style_code' },
            { type: 'text', text: ' component' },
          ],
        },
        // Long paragraph for pagination demo
        {
          id: 'block_22',
          type: 'paragraph',
          content: [
            { 
              type: 'text', 
              text: 'Check the documentation for more details on customization and advanced features. '.repeat(20),
            },
          ],
        },
        // Additional content
        {
          id: 'block_23',
          type: 'heading',
          level: 2,
          content: [{ type: 'text', text: 'Additional Information' }],
        },
        {
          id: 'block_24',
          type: 'paragraph',
          content: [
            { 
              type: 'text', 
              text: 'Check the documentation for more details on customization and advanced features. '.repeat(20),
            },
          ],
        },
        {
          id: 'block_25',
          type: 'paragraph',
          content: [
            { 
              type: 'text', 
              text: 'Check the documentation for more details on customization and advanced features. '.repeat(20),
            },
          ],
        },
      ],
    },
  ],
  createdAt: new Date().toISOString(),
  updatedAt: new Date().toISOString(),
};

// Phase 1.6 smoke-test panel. Spins up two DocumentImpl instances on the same
// page, each with its own InMemoryProvider connected to a shared roomId, and
// renders them side-by-side. Edits in one editor should converge into the
// other; remote cursors should render with the peer's color/name.
function CollabDemo({
  initialData,
  paneWidth,
  paneHeight,
}: {
  initialData: DocumentData;
  paneWidth: number;
  paneHeight: number;
}) {
  const [docs, setDocs] = useState<[DocumentImpl, DocumentImpl] | null>(null);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    let cancelled = false;
    let docA: DocumentImpl | null = null;
    let docB: DocumentImpl | null = null;

    (async () => {
      try {
        docA = new DocumentImpl();
        docB = new DocumentImpl();
        // Both docs start from identical content; Yjs will merge the
        // identical state into a no-op when SyncStep1 / SyncStep2 round-trip.
        docA.setData(initialData);
        docB.setData(initialData);

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
          docA.attachCollab(providerA, identityA),
          docB.attachCollab(providerB, identityB),
        ]);

        if (cancelled) {
          docA.detachCollab();
          docB.detachCollab();
          return;
        }
        setDocs([docA, docB]);
      } catch (e) {
        if (!cancelled) setError(e instanceof Error ? e.message : String(e));
      }
    })();

    return () => {
      cancelled = true;
      try { docA?.detachCollab(); } catch { /* ignore */ }
      try { docB?.detachCollab(); } catch { /* ignore */ }
    };
  }, [initialData]);

  if (error) return <div style={{ padding: 16, color: 'crimson' }}>Error: {error}</div>;
  if (!docs) return <div style={{ padding: 16 }}>Connecting collab peers…</div>;

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
        <div style={{ padding: '4px 12px', background: '#ff6b6b', color: 'white', fontWeight: 500 }}>
          Alice
        </div>
        <div style={{ flex: 1, minHeight: 0 }}>
          <DocumentProvider
            document={docs[0]}
            currentUser={DEMO_USERS[0]}
            mentionableUsers={DEMO_USERS}
            onCommentEvent={logCommentEvent}
          >
            <DocumentEditor width={paneWidth} height={paneHeight} showToolbar={true} showRuler={false} />
          </DocumentProvider>
        </div>
      </div>
      <div style={{ ...paneStyle, borderRight: 'none' }}>
        <div style={{ padding: '4px 12px', background: '#4ecdc4', color: 'white', fontWeight: 500 }}>
          Bob
        </div>
        <div style={{ flex: 1, minHeight: 0 }}>
          <DocumentProvider
            document={docs[1]}
            currentUser={DEMO_USERS[1]}
            mentionableUsers={DEMO_USERS}
            onCommentEvent={logCommentEvent}
          >
            <DocumentEditor width={paneWidth} height={paneHeight} showToolbar={true} showRuler={false} />
          </DocumentProvider>
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
  initialData,
  width,
  height,
  config,
}: {
  initialData: DocumentData;
  width: number;
  height: number;
  config: WsConfig;
}) {
  const [doc, setDoc] = useState<DocumentImpl | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [status, setStatus] = useState<CollabStatus>('connecting');

  useEffect(() => {
    let cancelled = false;
    let provider: WebSocketProvider | null = null;
    let document: DocumentImpl | null = null;
    (async () => {
      try {
        document = new DocumentImpl();
        document.setData(initialData);
        provider = new WebSocketProvider({ url: config.url });
        provider.onStatusChange((s) => {
          if (!cancelled) setStatus(s);
        });
        await document.attachCollab(provider, config.identity, {
          roomId: config.roomId,
          persistenceKey: `docs-demo:${config.roomId}`,
        });
        if (cancelled) {
          document.detachCollab();
          return;
        }
        setDoc(document);
      } catch (e) {
        if (!cancelled) setError(e instanceof Error ? e.message : String(e));
      }
    })();
    return () => {
      cancelled = true;
      try { document?.detachCollab(); } catch { /* ignore */ }
    };
  }, [initialData, config.url, config.roomId, config.identity]);

  if (error) return <div style={{ padding: 16, color: 'crimson' }}>WS error: {error}</div>;
  if (!doc) return <div style={{ padding: 16 }}>Connecting to {config.url}…</div>;
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
        <DocumentProvider
          document={doc}
          currentUser={{ id: config.identity.userId, name: config.identity.displayName }}
          mentionableUsers={DEMO_USERS}
          onCommentEvent={logCommentEvent}
        >
          <DocumentEditor width={width} height={height - 28} showToolbar={true} showRuler={true} />
        </DocumentProvider>
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
  const [wsConfig] = useState<WsConfig | null>(() => parseWsConfig());

  const [document] = useState(() => {
    // Create a new document and load the sample data
    const doc = new DocumentImpl();
    doc.setData(sampleDocumentData);
    return doc;
  });

  const [dimensions, setDimensions] = useState({
    width: window.innerWidth,
    height: window.innerHeight - 72,
  });

  useEffect(() => {
    const handleResize = () => {
      setDimensions({
        width: window.innerWidth,
        height: window.innerHeight - 72,
      });
    };

    window.addEventListener('resize', handleResize);
    return () => window.removeEventListener('resize', handleResize);
  }, []);

  const handleSave = () => {
    const data = document.getData();
    console.log('Document data:', JSON.stringify(data, null, 2));
    alert('Document data logged to console (formatted JSON)');
  };

  const handleNew = () => {
    // Create a new blank document
    const blankData: DocumentData = {
      id: 'new-document',
      title: 'Untitled Document',
      defaultPageConfig: sampleDocumentData.defaultPageConfig,
      textStylePool: {},
      paragraphStylePool: {},
      sections: [
        {
          id: 'section_1',
          pageConfig: sampleDocumentData.defaultPageConfig,
          blocks: [
            {
              id: 'block_1',
              type: 'paragraph',
              content: [],
            },
          ],
        },
      ],
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    };
    document.setData(blankData);
  };

  return (
    <div className="app">
      <header className="app-header">
        <div className="header-left">
          <h1>Pagent Docs Demo</h1>
          <p>Standalone demo of Spindle Docs library - No backend required</p>
        </div>
        <div className="header-right">
          <button className="header-button" onClick={handleNew}>
            New Document
          </button>
          <button className="header-button" onClick={handleSave}>
            Save to Console
          </button>
          <button className="header-button" onClick={() => setCollabMode((v) => !v)}>
            {collabMode ? 'Single editor' : 'Collab demo'}
          </button>
        </div>
      </header>
      <main className="app-main">
        {wsConfig ? (
          <WsDemo
            initialData={sampleDocumentData}
            width={dimensions.width}
            height={dimensions.height}
            config={wsConfig}
          />
        ) : collabMode ? (
          <CollabDemo
            initialData={sampleDocumentData}
            paneWidth={Math.floor(dimensions.width / 2) - 1}
            paneHeight={dimensions.height - 28 /* peer header strip */}
          />
        ) : (
          <DocumentProvider
            document={document}
            currentUser={DEMO_USERS[0]}
            mentionableUsers={DEMO_USERS}
            onCommentEvent={logCommentEvent}
          >
            <DocumentEditor
              width={dimensions.width}
              height={dimensions.height}
              showToolbar={true}
              showRuler={true}
            />
          </DocumentProvider>
        )}
      </main>
    </div>
  );
}

export default App;
