/**
 * Internal links: a link whose href is `#name` points at the bookmark of that name (Word's cross-references
 * and tables of contents). Clicking one scrolls to the bookmark.
 */

import type { Node as PmNode } from 'prosemirror-model';

/** The bookmark name an href points at (`#name`), or null when it is not an internal link. */
export function internalLinkTarget(href: string | null | undefined): string | null {
  if (!href || href.length < 2 || !href.startsWith('#')) return null;
  const raw = href.slice(1);
  try {
    return decodeURIComponent(raw);
  } catch {
    return raw;
  }
}

/** The document position of the bookmark called `name`, or null when there is none (the first one wins). */
export function findBookmarkPos(doc: PmNode, name: string): number | null {
  let found: number | null = null;
  doc.descendants((node, pos) => {
    if (found !== null) return false;
    if (node.type.name === 'bookmark' && node.attrs.name === name) {
      found = pos;
      return false;
    }
    return true;
  });
  return found;
}

/** The smallest painted element that covers a document position, or null when its page is not painted. */
export function findElementAtPos(container: ParentNode, pos: number): HTMLElement | null {
  let best: HTMLElement | null = null;
  let bestSize = Infinity;
  container.querySelectorAll<HTMLElement>('[data-pm-start]').forEach((el) => {
    const start = Number(el.dataset.pmStart);
    const end = Number(el.dataset.pmEnd);
    if (Number.isNaN(start) || Number.isNaN(end) || pos < start || pos >= end) return;
    if (end - start < bestSize) {
      best = el;
      bestSize = end - start;
    }
  });
  return best;
}
