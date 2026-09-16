// DeckProvider — supplies the engine + its ElementStore to the component tree.
// The engine is passed in as a prop (the app owns its lifecycle, exactly like
// WorkbookProvider in sheets-react), so the same deck can be shared across
// panes or driven by collab.

import React, { createContext, useContext, useEffect, useMemo } from 'react';
import type { DeckImpl, CommentAuthor, SlidesCommentEvent } from '@b32nio/spindle-slides-core';
import { ElementStore } from './element-store';
import { NodeRegistry } from '../interactions/node-registry';
import { TransientStore } from '../interactions/transient-store';
import { EditingStore } from '../interactions/editing-store';
import { TableSelectionStore } from '../interactions/table-selection-store';
import { TableMetricsStore } from '../interactions/table-metrics-store';
import { UIStore } from '../interactions/ui-store';
import { ConnectorStore } from '../interactions/connector-store';

export interface DeckContextValue {
  deck: DeckImpl;
  store: ElementStore;
  /** element-id → wrapper DOM node, for direct-to-DOM gesture writes. */
  nodes: NodeRegistry;
  /** Per-gesture transient state (guides, marquee) consumed by overlays. */
  transient: TransientStore;
  /** Which element hosts the single live ProseMirror editor. */
  editing: EditingStore;
  /** The current cell-range selection inside a table (local, never in CRDT). */
  tableSel: TableSelectionStore;
  /** Measured row geometry per table, so overlays align with content-driven rows. */
  tableMetrics: TableMetricsStore;
  /** Editor-chrome UI state (e.g. comments sidebar open). */
  ui: UIStore;
  /** Connector hover/draft state for the drawing UI. */
  connectors: ConnectorStore;
  /** The signed-in user, used as the author of new comments. */
  currentUser?: CommentAuthor;
  /** Directory of users that can be @-mentioned in comments. */
  mentionableUsers: CommentAuthor[];
}

const DeckContext = createContext<DeckContextValue | null>(null);

export interface DeckProviderProps {
  deck: DeckImpl;
  children: React.ReactNode;
  currentUser?: CommentAuthor;
  mentionableUsers?: CommentAuthor[];
  /** Host hook for the local user's own comment activity (notifications, etc.). */
  onCommentEvent?: (event: SlidesCommentEvent) => void;
}

export function DeckProvider({ deck, children, currentUser, mentionableUsers, onCommentEvent }: DeckProviderProps): React.ReactElement {
  const store = useMemo(() => new ElementStore(deck), [deck]);
  const nodes = useMemo(() => new NodeRegistry(), [deck]);
  const transient = useMemo(() => new TransientStore(), [deck]);
  const editing = useMemo(() => new EditingStore(), [deck]);
  const tableSel = useMemo(() => new TableSelectionStore(), [deck]);
  const tableMetrics = useMemo(() => new TableMetricsStore(), [deck]);
  const ui = useMemo(() => new UIStore(), [deck]);
  const connectors = useMemo(() => new ConnectorStore(), [deck]);
  // Subscribe in the effect (not the store constructor) so it survives React
  // StrictMode's mount→unmount→mount: connect → dispose → connect.
  useEffect(() => {
    store.connect();
    return () => store.dispose();
  }, [store]);

  useEffect(() => {
    if (!onCommentEvent) return;
    return deck.on('commentEvent', (e) => onCommentEvent(e.payload as SlidesCommentEvent));
  }, [deck, onCommentEvent]);

  const value = useMemo<DeckContextValue>(
    () => ({ deck, store, nodes, transient, editing, tableSel, tableMetrics, ui, connectors, currentUser, mentionableUsers: mentionableUsers ?? [] }),
    [deck, store, nodes, transient, editing, tableSel, tableMetrics, ui, connectors, currentUser, mentionableUsers]
  );
  return <DeckContext.Provider value={value}>{children}</DeckContext.Provider>;
}

export function useDeckContext(): DeckContextValue {
  const ctx = useContext(DeckContext);
  if (!ctx) throw new Error('useDeckContext must be used within a <DeckProvider>');
  return ctx;
}
