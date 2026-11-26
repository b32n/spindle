// Remote-cursor overlay for the TrueLayoutEditor painted pages.
//
// y-prosemirror's yCursorPlugin adds its decorations to ProseMirror's own DOM.
// In TrueLayoutEditor that DOM is the *hidden* input-handling editor — the
// visible pages are painted separately by DomPainter and never see those
// decorations. This overlay closes the gap:
//
//   awareness.on('change')
//     → for each peer, decode Y.RelativePosition → PM absolute position
//     → ask SelectionOverlayManager to translate PM positions to
//       (pageIndex, x, y) on the visible pages
//     → render absolute-positioned <div> caret + name label + selection rects
//
// Positioning math mirrors the local selection overlay (pageY + margins
// + scale) so caret elements live in the same coordinate system as the
// painted pages.

import * as Y from 'yjs';
import { Awareness } from 'y-protocols/awareness';
import { EditorView } from 'prosemirror-view';
import { relativePositionToAbsolutePosition, ySyncPluginKey } from 'y-prosemirror';

import { getPageContentOrigin } from './true-layout-engine';
import type { CaretPosition, SelectionOverlayManager } from './selection-overlay';

interface RemoteUser {
  name?: string;
  color?: string;
}

interface PeerElements {
  caretEl: HTMLElement;
  labelEl: HTMLElement;
  selectionEls: HTMLElement[];
}

export class RemoteCursorOverlay {
  private container: HTMLElement | null = null;
  private editorView: EditorView | null = null;
  private animationFrame: number | null = null;
  private peers = new Map<number, PeerElements>();
  private boundAwarenessHandler: () => void;

  constructor(
    private awareness: Awareness,
    private localOverlay: SelectionOverlayManager,
  ) {
    this.boundAwarenessHandler = () => this.scheduleRender();
    this.awareness.on('change', this.boundAwarenessHandler);
  }

  initialize(container: HTMLElement): void {
    this.container = container;
    this.scheduleRender();
  }

  setEditorView(view: EditorView | null): void {
    this.editorView = view;
    this.scheduleRender();
  }

  /** Re-render when the page layout shifts (block reflow, page split, etc.). */
  refresh(): void {
    this.scheduleRender();
  }

  destroy(): void {
    this.awareness.off('change', this.boundAwarenessHandler);
    if (this.animationFrame !== null) {
      cancelAnimationFrame(this.animationFrame);
      this.animationFrame = null;
    }
    for (const peer of this.peers.values()) {
      peer.caretEl.remove();
      for (const el of peer.selectionEls) el.remove();
    }
    this.peers.clear();
  }

  private scheduleRender(): void {
    if (this.animationFrame !== null) return;
    this.animationFrame = requestAnimationFrame(() => {
      this.animationFrame = null;
      this.render();
    });
  }

  private render(): void {
    if (!this.container) return;
    if (!this.editorView) {
      this.clearAll();
      return;
    }
    const state = this.editorView.state;
    const ystate = ySyncPluginKey.getState(state);
    if (!ystate || !ystate.binding || ystate.binding.mapping.size === 0) {
      // ySyncPlugin hasn't finished mapping the doc yet; positions would
      // resolve to null. Try again next tick.
      return;
    }

    const ydoc: Y.Doc = ystate.doc;
    const ytype: Y.XmlFragment = ystate.type;
    const mapping = ystate.binding.mapping;
    const localClientId = ydoc.clientID;

    const seen = new Set<number>();

    this.awareness.getStates().forEach((aw, clientId) => {
      if (clientId === localClientId) return;
      if (!aw || !aw.cursor) return;
      const user = (aw.user as RemoteUser | undefined) ?? {};

      let anchor: number | null = null;
      let head: number | null = null;
      try {
        anchor = relativePositionToAbsolutePosition(
          ydoc,
          ytype,
          Y.createRelativePositionFromJSON(aw.cursor.anchor),
          mapping,
        );
        head = relativePositionToAbsolutePosition(
          ydoc,
          ytype,
          Y.createRelativePositionFromJSON(aw.cursor.head),
          mapping,
        );
      } catch {
        return;
      }
      if (anchor === null || head === null) return;

      seen.add(clientId);
      this.paintPeer(clientId, user, anchor, head);
    });

    // Drop peers that have left or stopped sharing a cursor.
    for (const [clientId, peer] of this.peers) {
      if (!seen.has(clientId)) {
        peer.caretEl.remove();
        for (const el of peer.selectionEls) el.remove();
        this.peers.delete(clientId);
      }
    }
  }

  private clearAll(): void {
    for (const peer of this.peers.values()) {
      peer.caretEl.remove();
      for (const el of peer.selectionEls) el.remove();
    }
    this.peers.clear();
  }

  private paintPeer(
    clientId: number,
    user: RemoteUser,
    anchor: number,
    head: number,
  ): void {
    if (!this.container) return;
    const layout = this.localOverlay.getLayout();
    if (!layout) return;

    const caretPos = this.localOverlay.getCaretForPos(head);
    if (!caretPos) {
      // Position couldn't be mapped to a visible page (e.g. block not yet
      // measured). Hide any stale caret for this peer.
      const stale = this.peers.get(clientId);
      if (stale) {
        stale.caretEl.remove();
        for (const el of stale.selectionEls) el.remove();
        this.peers.delete(clientId);
      }
      return;
    }

    const color = user.color ?? '#ffa500';
    const name = user.name ?? `User ${clientId}`;

    let peer = this.peers.get(clientId);
    if (!peer) {
      const caretEl = document.createElement('div');
      caretEl.className = 'spindle-remote-caret';
      const labelEl = document.createElement('div');
      labelEl.className = 'spindle-remote-caret-label';
      caretEl.appendChild(labelEl);
      this.container.appendChild(caretEl);
      peer = { caretEl, labelEl, selectionEls: [] };
      this.peers.set(clientId, peer);
    }

    this.positionCaret(peer.caretEl, caretPos, color);
    if (peer.labelEl.textContent !== name) {
      peer.labelEl.textContent = name;
    }
    // Label background must come from inline style — see collab-cursor-styles
    // for why CSS currentColor doesn't work here.
    peer.labelEl.style.backgroundColor = color;

    // Selection rects when anchor ≠ head.
    const from = Math.min(anchor, head);
    const to = Math.max(anchor, head);
    for (const el of peer.selectionEls) el.remove();
    peer.selectionEls = [];

    if (from !== to) {
      const rects = this.localOverlay.getRectsForRange(from, to);
      for (const rect of rects) {
        const el = document.createElement('div');
        el.className = 'spindle-remote-selection-rect';
        this.positionRect(el, rect, color, layout);
        this.container.appendChild(el);
        peer.selectionEls.push(el);
      }
    }
  }

  private positionCaret(
    el: HTMLElement,
    pos: CaretPosition,
    color: string,
  ): void {
    const layout = this.localOverlay.getLayout();
    if (!layout) return;
    const origin = getPageContentOrigin(layout, pos.pageIndex);
    const absX = origin.x + pos.x;
    const absY = origin.y + pos.y;
    el.style.transform = `translate(${absX}px, ${absY}px)`;
    el.style.height = `${pos.height}px`;
    el.style.color = color;
  }

  private positionRect(
    el: HTMLElement,
    rect: { pageIndex: number; x: number; y: number; width: number; height: number },
    color: string,
    layout: NonNullable<ReturnType<SelectionOverlayManager['getLayout']>>,
  ): void {
    const origin = getPageContentOrigin(layout, rect.pageIndex);
    const absX = origin.x + rect.x;
    const absY = origin.y + rect.y;
    el.style.transform = `translate(${absX}px, ${absY}px)`;
    el.style.width = `${rect.width}px`;
    el.style.height = `${rect.height}px`;
    el.style.backgroundColor = color;
  }
}
