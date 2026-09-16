// Per-gesture transient state (snap guides, marquee rect, and live element
// frames while dragging) that must not trigger a full React commit. Only the
// tiny overlay components subscribe. Element wrapper transforms are written
// straight to the DOM (see gesture.ts); this store carries just what overlays
// need to paint.

import type { GuideLine, Rect, Frame } from '@b32nio/spindle-slides-core';

export interface TransientState {
  guides: GuideLine[];
  marquee: Rect | null;
  /** Live frames of the elements under an active gesture (for the selection box). */
  liveFrames: Map<string, Frame> | null;
}

const EMPTY: TransientState = { guides: [], marquee: null, liveFrames: null };

export class TransientStore {
  private state: TransientState = EMPTY;
  private listeners = new Set<() => void>();

  get = (): TransientState => this.state;

  set(patch: Partial<TransientState>): void {
    this.state = { ...this.state, ...patch };
    for (const l of this.listeners) l();
  }

  clear(): void {
    if (this.state === EMPTY) return;
    this.state = EMPTY;
    for (const l of this.listeners) l();
  }

  subscribe = (cb: () => void): (() => void) => {
    this.listeners.add(cb);
    return () => this.listeners.delete(cb);
  };
}
