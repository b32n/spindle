// Formula dependency graph for incremental recalculation.

import type {
  FormulaGraph,
  FormulaNode,
  FormulaDependencies,
  RangeDependency,
  CellResolver,
  CellValue,
} from './types';

export class FormulaGraphImpl implements FormulaGraph {
  nodes: Map<string, FormulaNode> = new Map();

  /**
   * Reverse index for single-cell dependencies: for any referenced cell key —
   * a formula cell *or* a plain-value cell — the formulas that read it.
   */
  private cellDependents: Map<string, Set<string>> = new Map();

  /** Formula keys that have at least one range dependency (scanned for containment). */
  private rangeFormulas: Set<string> = new Set();

  /** Cell keys of volatile formulas — kept in sync so recalc need not scan. */
  private volatileNodes: Set<string> = new Set();

  addFormula(
    cellKey: string,
    formula: string,
    dependencies: FormulaDependencies,
    volatile = false
  ): void {
    // If this cell was already a formula, unlink its previous dependencies
    // first so re-adding it leaves no stale index entries.
    const existing = this.nodes.get(cellKey);
    if (existing) this.unlink(cellKey, existing);

    const node: FormulaNode = {
      cellKey,
      formula,
      dependencies: new Set(dependencies.cells),
      rangeRects: dependencies.ranges.slice(),
      isDirty: true,
      volatile,
    };
    this.nodes.set(cellKey, node);

    for (const depKey of node.dependencies) {
      let set = this.cellDependents.get(depKey);
      if (!set) {
        set = new Set();
        this.cellDependents.set(depKey, set);
      }
      set.add(cellKey);
    }

    if (node.rangeRects.length > 0) this.rangeFormulas.add(cellKey);
    else this.rangeFormulas.delete(cellKey);

    if (volatile) this.volatileNodes.add(cellKey);
    else this.volatileNodes.delete(cellKey);
  }

  removeFormula(cellKey: string): void {
    const node = this.nodes.get(cellKey);
    if (!node) return;
    this.unlink(cellKey, node);
    this.nodes.delete(cellKey);
    this.rangeFormulas.delete(cellKey);
    this.volatileNodes.delete(cellKey);
  }

  /** Drop a formula from every index that points at it. */
  private unlink(cellKey: string, node: FormulaNode): void {
    for (const depKey of node.dependencies) {
      const set = this.cellDependents.get(depKey);
      if (!set) continue;
      set.delete(cellKey);
      if (set.size === 0) this.cellDependents.delete(depKey);
    }
  }

  setSpillFootprint(cellKey: string, rects: RangeDependency[]): void {
    const node = this.nodes.get(cellKey);
    if (!node) return;
    node.spillRects = rects.length > 0 ? rects.slice() : undefined;
    if (rects.length > 0 || node.rangeRects.length > 0) this.rangeFormulas.add(cellKey);
    else this.rangeFormulas.delete(cellKey);
  }

  markClean(cellKey: string, value: CellValue): void {
    const node = this.nodes.get(cellKey);
    if (node) {
      node.isDirty = false;
      node.cachedValue = value;
    }
  }

  /**
   * Invalidate every formula's cached value. Used after a structural change
   * (sheet delete/rename) that can affect cross-sheet references anywhere in the
   * workbook, so a subsequent full re-evaluation can't return a stale cache for
   * a formula that happens to be evaluated before a dependency it references.
   */
  markAllDirty(): void {
    for (const node of this.nodes.values()) {
      node.isDirty = true;
      node.cachedValue = undefined;
    }
  }

  /** Is `idx` inside the rectangle, resolving its corner keys to indices now? */
  private rectContains(
    rect: RangeDependency,
    idx: { sheetId: string; row: number; col: number },
    resolveCell: CellResolver
  ): boolean {
    const a = resolveCell(rect.startKey);
    const b = resolveCell(rect.endKey);
    if (!a || !b) return false;
    // The range and the candidate cell must live on the same sheet — otherwise
    // a same-index cell on another sheet would falsely test inside the range.
    if (a.sheetId !== b.sheetId || a.sheetId !== idx.sheetId) return false;
    return (
      idx.row >= Math.min(a.row, b.row) &&
      idx.row <= Math.max(a.row, b.row) &&
      idx.col >= Math.min(a.col, b.col) &&
      idx.col <= Math.max(a.col, b.col)
    );
  }

  /** Formulas that directly depend on `key` — via a single ref or a range. */
  private directDependents(key: string, resolveCell: CellResolver): Set<string> {
    const result = new Set(this.cellDependents.get(key) ?? []);
    if (this.rangeFormulas.size > 0) {
      const idx = resolveCell(key);
      if (idx) {
        for (const formulaKey of this.rangeFormulas) {
          if (result.has(formulaKey)) continue;
          const node = this.nodes.get(formulaKey);
          if (!node) continue;
          const rects = node.spillRects ? [...node.rangeRects, ...node.spillRects] : node.rangeRects;
          for (const rect of rects) {
            if (this.rectContains(rect, idx, resolveCell)) {
              result.add(formulaKey);
              break;
            }
          }
        }
      }
    }
    return result;
  }

  collectDirty(
    seeds: string[],
    resolveCell: CellResolver
  ): { dirty: Set<string>; edges: Map<string, Set<string>> } {
    const dirty = new Set<string>();
    const edges = new Map<string, Set<string>>();
    const frontier: string[] = [];

    const enqueue = (key: string): void => {
      if (dirty.has(key)) return;
      dirty.add(key);
      frontier.push(key);
      const node = this.nodes.get(key);
      if (node) {
        node.isDirty = true;
        node.cachedValue = undefined;
      }
    };

    // Seeds: direct dependents of every changed cell, and every volatile formula.
    for (const seed of seeds) {
      for (const formulaKey of this.directDependents(seed, resolveCell)) {
        enqueue(formulaKey);
      }
    }
    for (const volatileKey of this.volatileNodes) {
      enqueue(volatileKey);
    }

    while (frontier.length > 0) {
      const key = frontier.pop();
      if (key === undefined) continue;
      for (const dependent of this.directDependents(key, resolveCell)) {
        // `dependent` depends on `key`; both are dirty, so this is a real edge.
        let deps = edges.get(dependent);
        if (!deps) {
          deps = new Set();
          edges.set(dependent, deps);
        }
        deps.add(key);
        enqueue(dependent);
      }
    }

    return { dirty, edges };
  }

  topologicalOrder(
    dirty: Set<string>,
    edges: Map<string, Set<string>>
  ): { ordered: string[]; cyclic: string[] } {
    // Kahn's algorithm over the explicit dependency edges of the dirty set.
    const inDegree = new Map<string, number>();
    const dependentsOf = new Map<string, string[]>();
    for (const key of dirty) inDegree.set(key, 0);

    for (const [formula, deps] of edges) {
      if (!dirty.has(formula)) continue;
      let degree = 0;
      for (const dep of deps) {
        if (!dirty.has(dep)) continue;
        degree++;
        let list = dependentsOf.get(dep);
        if (!list) {
          list = [];
          dependentsOf.set(dep, list);
        }
        list.push(formula);
      }
      inDegree.set(formula, degree);
    }

    const queue: string[] = [];
    for (const [key, degree] of inDegree) {
      if (degree === 0) queue.push(key);
    }

    const ordered: string[] = [];
    let head = 0;
    while (head < queue.length) {
      const key = queue[head];
      head++;
      ordered.push(key);
      for (const dependent of dependentsOf.get(key) ?? []) {
        const next = (inDegree.get(dependent) ?? 0) - 1;
        inDegree.set(dependent, next);
        if (next === 0) queue.push(dependent);
      }
    }

    // Anything still owing dependencies is in, or downstream of, a cycle.
    const cyclic: string[] = [];
    for (const [key, degree] of inDegree) {
      if (degree > 0) cyclic.push(key);
    }
    return { ordered, cyclic };
  }
}
