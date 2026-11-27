// Spill index — the dynamic-array (spill) overlay for a sheet.
//
// A spilling formula's anchor cell holds the formula and the array's top-left
// value; the rest of the array fills neighbouring cells. Those spilled cells
// are NOT stored in `sheet.cells` — they are a derived overlay, recomputed by
// re-evaluating anchor formulas, so they never sync over the collaboration
// CRDT (only the anchor's formula does).

export interface SpillRegion {
  anchorRow: number;
  anchorCol: number;
  rows: number;
  cols: number;
  values: unknown[][];
}

function posKey(row: number, col: number): string {
  return `${row}:${col}`;
}

export class SpillIndex {
  /** anchor "row:col" -> region */
  private byAnchor = new Map<string, SpillRegion>();
  /** covered, non-anchor "row:col" -> anchor "row:col" */
  private coverage = new Map<string, string>();

  /** Register (or replace) an anchor's spill region. */
  register(region: SpillRegion): void {
    const anchorKey = posKey(region.anchorRow, region.anchorCol);
    this.byAnchor.set(anchorKey, region);
    for (let r = 0; r < region.rows; r++) {
      for (let c = 0; c < region.cols; c++) {
        if (r === 0 && c === 0) continue; // the anchor is not "covered"
        this.coverage.set(posKey(region.anchorRow + r, region.anchorCol + c), anchorKey);
      }
    }
  }

  /** Remove an anchor's region; returns it so the caller can revisit its footprint. */
  unregister(anchorRow: number, anchorCol: number): SpillRegion | undefined {
    const anchorKey = posKey(anchorRow, anchorCol);
    const region = this.byAnchor.get(anchorKey);
    if (!region) return undefined;
    this.byAnchor.delete(anchorKey);
    for (let r = 0; r < region.rows; r++) {
      for (let c = 0; c < region.cols; c++) {
        const key = posKey(region.anchorRow + r, region.anchorCol + c);
        if (this.coverage.get(key) === anchorKey) this.coverage.delete(key);
      }
    }
    return region;
  }

  /** Every registered region. */
  regions(): IterableIterator<SpillRegion> {
    return this.byAnchor.values();
  }

  /** The region anchored at this cell, if any. */
  regionAt(anchorRow: number, anchorCol: number): SpillRegion | undefined {
    return this.byAnchor.get(posKey(anchorRow, anchorCol));
  }

  /** Whether a cell is covered by a spill it does not itself anchor. */
  isCovered(row: number, col: number): boolean {
    return this.coverage.has(posKey(row, col));
  }

  /** The spilled value shown at a covered cell, or undefined if not covered. */
  spilledValueAt(row: number, col: number): unknown | undefined {
    const anchorKey = this.coverage.get(posKey(row, col));
    if (!anchorKey) return undefined;
    const region = this.byAnchor.get(anchorKey);
    if (!region) return undefined;
    return region.values[row - region.anchorRow]?.[col - region.anchorCol];
  }

  /** The anchor position covering a cell (covered cell or the anchor itself). */
  anchorOf(row: number, col: number): { row: number; col: number } | undefined {
    if (this.byAnchor.has(posKey(row, col))) return { row, col };
    const anchorKey = this.coverage.get(posKey(row, col));
    if (!anchorKey) return undefined;
    const [r, c] = anchorKey.split(':');
    return { row: Number(r), col: Number(c) };
  }

  clear(): void {
    this.byAnchor.clear();
    this.coverage.clear();
  }
}

/**
 * What a legacy (CSE) array formula shows over its extent (`span`, the rows and columns of the range it was entered over): the
 * result cut to it. As Excel does, a result of one row (or one column) is repeated down (or across) a bigger extent, so a single value fills all
 * of it; a dimension where a result of more than one row (or column) is smaller than the extent is padded with `#N/A`. For a one-cell extent
 * it is the first value.
 */
export function fitToExtent(result: unknown, span: { rows: number; cols: number }): unknown {
  const matrix = Array.isArray(result) ? (result as unknown[][]) : [[result]];
  // (a blank cell of the result is null: it stays blank, only what the result does not reach is `#N/A`)
  const first = matrix[0]?.[0] === undefined ? '#N/A' : matrix[0][0];
  if (span.rows * span.cols === 1) return first;
  const repeatDown = matrix.length === 1;
  const repeatAcross = matrix[0]?.length === 1;
  return Array.from({ length: span.rows }, (_, r) =>
    Array.from({ length: span.cols }, (_, c) => {
      const value = matrix[repeatDown ? 0 : r]?.[repeatAcross ? 0 : c];
      return value === undefined ? '#N/A' : value;
    })
  );
}
