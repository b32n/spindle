// Imperative gesture controllers. During a gesture, element wrapper transforms
// are written straight to the DOM (via the NodeRegistry) and guides/marquee go
// into the TransientStore — no React commit per pointer move. On pointerup the
// controller commits once through the engine (one batch / undo entry) and React
// reconciles.

import {
  collectSnapTargets,
  computeSnap,
  boundsCandidates,
  unionAABB,
  rectFromPoints,
  frameIntersectsRect,
  resizeFrame,
  rotateFrame,
  scaleGroup,
  rotateGroup,
  type DeckImpl,
  type SlideElement,
  type Frame,
  type Point,
  type Rect,
  type ResizeHandle,
} from '@b32nio/spindle-slides-core';
import { NodeRegistry } from './node-registry';
import { TransientStore } from './transient-store';

export interface GestureContext {
  deck: DeckImpl;
  nodes: NodeRegistry;
  transient: TransientStore;
  slideId: string;
  slideSize: { w: number; h: number };
  /** Snap threshold in slide units (screen px ÷ zoom). */
  snapThreshold: number;
}

export interface Gesture {
  onMove(p: Point): void;
  onEnd(): void;
}

function frameOf(el: SlideElement): Frame {
  return { x: el.x, y: el.y, w: el.w, h: el.h, rotation: el.rotation };
}

function writeNode(nodes: NodeRegistry, id: string, f: Frame): void {
  const node = nodes.get(id);
  if (!node) return;
  node.style.transform = `translate(${f.x}px, ${f.y}px) rotate(${f.rotation}deg)`;
  node.style.width = `${f.w}px`;
  node.style.height = `${f.h}px`;
}

// ── Move ─────────────────────────────────────────────────────────────────────

export function createMoveGesture(ctx: GestureContext, start: Point, selectedIds: string[]): Gesture {
  const startFrames = new Map<string, Frame>();
  for (const id of selectedIds) {
    const el = ctx.deck.getElement(id);
    if (el) startFrames.set(id, frameOf(el));
  }
  const others = ctx.deck
    .getElementsForSlide(ctx.slideId)
    .filter((e) => !startFrames.has(e.id))
    .map(frameOf);
  const targets = collectSnapTargets(others, ctx.slideSize);
  const startBounds = unionAABB([...startFrames.values()]) ?? { x: 0, y: 0, w: 0, h: 0 };
  let last = new Map<string, Frame>();

  return {
    onMove(p) {
      let dx = p.x - start.x;
      let dy = p.y - start.y;
      const moved = boundsCandidates({ x: startBounds.x + dx, y: startBounds.y + dy, w: startBounds.w, h: startBounds.h });
      const snap = computeSnap(moved.xs, moved.ys, targets, ctx.snapThreshold);
      dx += snap.dx;
      dy += snap.dy;
      const live = new Map<string, Frame>();
      for (const [id, f] of startFrames) {
        const nf = { ...f, x: f.x + dx, y: f.y + dy };
        writeNode(ctx.nodes, id, nf);
        live.set(id, nf);
      }
      last = live;
      ctx.transient.set({ guides: snap.guides, liveFrames: live });
    },
    onEnd() {
      if (last.size > 0) {
        ctx.deck.setFrames([...last].map(([id, frame]) => ({ id, frame })));
      }
      ctx.transient.clear();
    },
  };
}

// ── Resize ─────────────────────────────────────────────────────────────────────

export function createResizeGesture(
  ctx: GestureContext,
  handle: ResizeHandle,
  selectedIds: string[],
  opts: { shiftKey: boolean }
): Gesture {
  const single = selectedIds.length === 1 ? ctx.deck.getElement(selectedIds[0]) : null;

  if (single) {
    const startFrame = frameOf(single);
    const lockAspect = single.type === 'image' ? true : opts.shiftKey;
    const others = ctx.deck.getElementsForSlide(ctx.slideId).filter((e) => e.id !== single.id).map(frameOf);
    const targets = collectSnapTargets(others, ctx.slideSize);
    let last: Frame | null = null;
    return {
      onMove(p) {
        // Snap the dragged pointer to nearby targets (edge snapping) when unrotated.
        let pt = p;
        if (startFrame.rotation === 0) {
          const snap = computeSnap([p.x], [p.y], targets, ctx.snapThreshold);
          pt = { x: p.x + snap.dx, y: p.y + snap.dy };
          ctx.transient.set({ guides: snap.guides });
        }
        const nf = resizeFrame(startFrame, handle, pt, { lockAspect });
        writeNode(ctx.nodes, single.id, nf);
        ctx.transient.set({ liveFrames: new Map([[single.id, nf]]) });
        last = nf;
      },
      onEnd() {
        if (last) ctx.deck.setFrames([{ id: single.id, frame: last }]);
        ctx.transient.clear();
      },
    };
  }

  // Multi-select: uniform (aspect-locked) group scale about the opposite corner.
  const startFrames = new Map<string, Frame>();
  for (const id of selectedIds) {
    const el = ctx.deck.getElement(id);
    if (el) startFrames.set(id, frameOf(el));
  }
  const bounds = unionAABB([...startFrames.values()]) ?? { x: 0, y: 0, w: 1, h: 1 };
  // Opposite corner of the handle is the scale origin.
  const origin: Point = {
    x: handle.includes('w') ? bounds.x + bounds.w : bounds.x,
    y: handle.includes('n') ? bounds.y + bounds.h : bounds.y,
  };
  const startHandlePt: Point = {
    x: handle.includes('w') ? bounds.x : bounds.x + bounds.w,
    y: handle.includes('n') ? bounds.y : bounds.y + bounds.h,
  };
  let last = new Map<string, Frame>();
  return {
    onMove(p) {
      const sx = (p.x - origin.x) / (startHandlePt.x - origin.x || 1);
      const sy = (p.y - origin.y) / (startHandlePt.y - origin.y || 1);
      const s = Math.max(0.05, Math.max(Math.abs(sx), Math.abs(sy)));
      const items = [...startFrames].map(([id, frame]) => ({ id, frame }));
      const scaled = scaleGroup(items, s, s, origin);
      const live = new Map<string, Frame>();
      for (const { id, frame } of scaled) {
        writeNode(ctx.nodes, id, frame);
        live.set(id, frame);
      }
      ctx.transient.set({ liveFrames: live });
      last = live;
    },
    onEnd() {
      if (last.size > 0) ctx.deck.setFrames([...last].map(([id, frame]) => ({ id, frame })));
      ctx.transient.clear();
    },
  };
}

// ── Rotate ─────────────────────────────────────────────────────────────────────

export function createRotateGesture(ctx: GestureContext, selectedIds: string[], opts: { shiftKey: boolean }): Gesture {
  const single = selectedIds.length === 1 ? ctx.deck.getElement(selectedIds[0]) : null;

  if (single) {
    const startFrame = frameOf(single);
    let last: Frame | null = null;
    return {
      onMove(p) {
        const nf = rotateFrame(startFrame, p, { snap: opts.shiftKey });
        writeNode(ctx.nodes, single.id, nf);
        ctx.transient.set({ liveFrames: new Map([[single.id, nf]]) });
        last = nf;
      },
      onEnd() {
        if (last) ctx.deck.setFrames([{ id: single.id, frame: last }]);
        ctx.transient.clear();
      },
    };
  }

  const startFrames = new Map<string, Frame>();
  for (const id of selectedIds) {
    const el = ctx.deck.getElement(id);
    if (el) startFrames.set(id, frameOf(el));
  }
  const bounds = unionAABB([...startFrames.values()]) ?? { x: 0, y: 0, w: 1, h: 1 };
  const pivot: Point = { x: bounds.x + bounds.w / 2, y: bounds.y + bounds.h / 2 };
  const startAngle = (p: Point) => (Math.atan2(p.y - pivot.y, p.x - pivot.x) * 180) / Math.PI;
  let a0: number | null = null;
  let last = new Map<string, Frame>();
  return {
    onMove(p) {
      if (a0 === null) a0 = startAngle(p);
      let delta = startAngle(p) - a0;
      if (opts.shiftKey) delta = Math.round(delta / 15) * 15;
      const items = [...startFrames].map(([id, frame]) => ({ id, frame }));
      const rotated = rotateGroup(items, delta, pivot);
      const live = new Map<string, Frame>();
      for (const { id, frame } of rotated) {
        writeNode(ctx.nodes, id, frame);
        live.set(id, frame);
      }
      ctx.transient.set({ liveFrames: live });
      last = live;
    },
    onEnd() {
      if (last.size > 0) ctx.deck.setFrames([...last].map(([id, frame]) => ({ id, frame })));
      ctx.transient.clear();
    },
  };
}

// ── Marquee ──────────────────────────────────────────────────────────────────

export function createMarqueeGesture(ctx: GestureContext, start: Point, additive: boolean, baseIds: string[]): Gesture {
  const elements = ctx.deck.getElementsForSlide(ctx.slideId);
  return {
    onMove(p) {
      const rect: Rect = rectFromPoints(start, p);
      ctx.transient.set({ marquee: rect });
      const hit = elements.filter((e) => frameIntersectsRect(frameOf(e), rect)).map((e) => e.id);
      const ids = additive ? Array.from(new Set([...baseIds, ...hit])) : hit;
      ctx.deck.setSelection({ slideId: ctx.slideId, elementIds: ids });
    },
    onEnd() {
      ctx.transient.clear();
    },
  };
}

/** Expand a set of ids to include all members of any group they belong to. */
export function expandGroups(deck: DeckImpl, ids: string[]): string[] {
  const out = new Set<string>();
  for (const id of ids) {
    const el = deck.getElement(id);
    if (!el) continue;
    if (el.groupId) for (const m of deck.getGroupMembers(el.groupId)) out.add(m);
    else out.add(id);
  }
  return [...out];
}
