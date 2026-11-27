// Shrink-to-fit and resize-to-fit measurement. Pure geometry lives in slides-core (`fitShrinkScale`); this module reads
// the live DOM: it sets the autofit variables on a text body, asks the browser whether the text overflows, and keeps a
// registry of the interactive text bodies so a toolbar action can refit an element that is not being edited.

import {
  autofitMode,
  fitShrinkScale,
  resolveInsets,
  textVert,
  type BodyStyle,
  type DeckImpl,
  type ShrinkResult,
} from '@b32nio/spindle-slides-core';

const bodies = new Map<string, HTMLElement>();

/** Register the interactive (stage) text body of an element; returns the unregister function. */
export function registerTextBody(id: string, node: HTMLElement): () => void {
  bodies.set(id, node);
  return () => {
    if (bodies.get(id) === node) bodies.delete(id);
  };
}

function setScale(box: HTMLElement, s: ShrinkResult): void {
  box.style.setProperty('--spindle-fs', String(s.fontScale));
  box.style.setProperty('--spindle-ls', String(1 - s.lnSpcReduction));
}

/** True when the content of `box` runs past its padding box along the text's flow axis (1px of slack for rounding). */
export function overflows(box: HTMLElement, vertical: boolean): boolean {
  return vertical ? box.scrollWidth > box.clientWidth + 1 : box.scrollHeight > box.clientHeight + 1;
}

/**
 * Shrink-to-fit a text body in place: find the largest scale at which the text fits, leave it applied through the
 * variables and return it. The box is measured top-anchored (a centred overflow would spill above the scroll origin,
 * where it cannot be measured) and restored afterwards.
 */
export function measureShrink(box: HTMLElement, vertical = false): ShrinkResult {
  const prev = box.style.justifyContent;
  box.style.justifyContent = 'flex-start';
  const result = fitShrinkScale((candidate) => {
    setScale(box, candidate);
    return !overflows(box, vertical);
  });
  box.style.justifyContent = prev;
  setScale(box, result);
  return result;
}

/** The height a text body needs to show all its text: its padding box with the height released. */
export function measureContentHeight(box: HTMLElement): number {
  const prev = box.style.height;
  box.style.height = 'auto';
  const h = Math.ceil(box.offsetHeight);
  box.style.height = prev;
  return h;
}

/** The body style after a shrink fit, keeping the other autofit fields. */
export function withShrink(body: BodyStyle | undefined, r: ShrinkResult): BodyStyle {
  return { ...body, autofit: { mode: 'shrink', fontScale: r.fontScale, lnSpcReduction: r.lnSpcReduction } };
}

/** Persist a fit result for an element when it differs from what is stored. */
export function persistFit(deck: Pick<DeckImpl, 'getElement' | 'updateElement'>, id: string, box: HTMLElement): void {
  const el = deck.getElement(id);
  if (!el || (el.type !== 'text' && el.type !== 'shape')) return;
  const body = el.bodyStyle;
  const mode = autofitMode(body);
  if (mode === 'shrink') {
    const r = measureShrink(box, textVert(body) !== 'horz');
    const cur = body?.autofit;
    if (cur?.fontScale !== r.fontScale || cur?.lnSpcReduction !== r.lnSpcReduction) {
      deck.updateElement(id, { bodyStyle: withShrink(body, r) });
    }
  } else if (mode === 'resize' && textVert(body) === 'horz' && el.type === 'text') {
    const h = measureContentHeight(box);
    if (h !== Math.ceil(el.h) && h > resolveInsets(body).t + resolveInsets(body).b) deck.updateElement(id, { h });
  }
}

/** Refit an element's text to its box from the registered stage body (no-op when the element is not on the stage). */
export function refitElement(deck: Pick<DeckImpl, 'getElement' | 'updateElement'>, id: string): void {
  const box = bodies.get(id);
  // A detached node has no layout (it would measure as fitting at scale 1), so never fit from one.
  if (box?.isConnected) persistFit(deck, id, box);
}

/**
 * Fit bookkeeping for one live editor session. Each document change calls `live()` (rescale through the CSS variables
 * only, so typing stays cheap); the commit calls `settle()` while the editor is attached (it needs layout) to store the
 * scale or resize the box; teardown calls `flush()`, which stores the last live scale without measuring because the
 * node may already be detached.
 */
export function createFitController(
  deck: Pick<DeckImpl, 'getElement' | 'updateElement'>,
  id: string,
  box: HTMLElement,
  disabled: boolean
): { live: () => void; settle: () => void; flush: () => void } {
  let pending: ShrinkResult | null = null;
  const bodyOf = (): BodyStyle | undefined => (deck.getElement(id) as { bodyStyle?: BodyStyle } | undefined)?.bodyStyle;
  return {
    live: () => {
      const body = bodyOf();
      if (disabled || autofitMode(body) !== 'shrink') return;
      pending = measureShrink(box, textVert(body) !== 'horz');
    },
    settle: () => {
      if (disabled) return;
      persistFit(deck, id, box);
      pending = null;
    },
    flush: () => {
      const body = bodyOf();
      if (pending && autofitMode(body) === 'shrink') deck.updateElement(id, { bodyStyle: withShrink(body, pending) });
      pending = null;
    },
  };
}

/** A ref callback that registers the stage's text body for an element (React 19 runs the returned cleanup on detach). */
export function textBodyRef(id: string, interactive: boolean): (node: HTMLDivElement | null) => (() => void) | undefined {
  return (node) => (node && interactive ? registerTextBody(id, node) : undefined);
}
