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
