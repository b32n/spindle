import type { ImportLossItem } from '../export/loss';
import type { OpaqueDropped } from './types';

/** A "kept, not editable" loss item: the part is stored and re-exported unchanged. */
export function keptLossItem(item: {
  id: string;
  label: string;
  count: number;
  where?: string[];
  effect?: string;
}): ImportLossItem {
  return {
    ...item,
    effect: item.effect ?? 'kept: stored unchanged and saved back, not editable',
    state: 'kept',
  };
}

/**
 * One loss item for parts the store refused (never a silent drop). `lost`, because they are not saved;
 * `where` lists the part names (capped at 5).
 */
export function droppedLossItem(
  id: string,
  label: string,
  dropped: readonly OpaqueDropped[]
): ImportLossItem {
  const where = dropped.flatMap((d) => (d.partName ? [d.partName] : [])).slice(0, 5);
  const full = dropped.some((d) => d.reason === 'store-full');
  return {
    id,
    label,
    count: dropped.length,
    ...(where.length > 0 && { where }),
    effect: full
      ? 'lost: dropped, the document’s kept parts are over the size limit'
      : 'lost: dropped, too large to keep',
    state: 'lost',
  };
}
