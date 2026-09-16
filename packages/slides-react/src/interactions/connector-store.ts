// ConnectorStore — transient state for the connector-drawing UI: which element
// is hovered (so its connection dots show) and the in-progress draft while
// dragging a new connector. Kept out of React commits; only the connection
// overlay subscribes, so hover/drag never re-render the slide's elements.

import type { AnchorId, Point } from '@b32nio/spindle-slides-core';

export interface ConnectorDraft {
  fromElementId: string;
  fromAnchor: AnchorId;
  from: Point;
  to: Point;
  /** The target anchor the end is currently snapped to, if any. */
  snap: { elementId: string; anchor: AnchorId } | null;
  /** The element under the cursor whose dots should be shown as candidates. */
  overElementId: string | null;
}

/** Live drag of an existing line's endpoint (resize+rotate via the tip). */
export interface ConnectorEdit {
  elementId: string;
  end: 'start' | 'end';
  point: Point;
  snap: { elementId: string; anchor: AnchorId } | null;
  /** Element under the cursor whose dots should show as snap candidates. */
  overElementId: string | null;
}

export interface ConnectorState {
  hoverId: string | null;
  draft: ConnectorDraft | null;
  edit: ConnectorEdit | null;
}

const EMPTY: ConnectorState = { hoverId: null, draft: null, edit: null };

export class ConnectorStore {
  private state: ConnectorState = EMPTY;
  private listeners = new Set<() => void>();

  get = (): ConnectorState => this.state;

  setHover(hoverId: string | null): void {
    if (this.state.hoverId === hoverId) return;
    this.state = { ...this.state, hoverId };
    this.emit();
  }

  setDraft(draft: ConnectorDraft | null): void {
    this.state = { ...this.state, draft };
    this.emit();
  }

  setEdit(edit: ConnectorEdit | null): void {
    this.state = { ...this.state, edit };
    this.emit();
  }

  private emit(): void {
    for (const l of this.listeners) l();
  }

  subscribe = (cb: () => void): (() => void) => {
    this.listeners.add(cb);
    return () => this.listeners.delete(cb);
  };
}
