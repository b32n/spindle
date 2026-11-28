// Style pool for shared style objects

import type { CellStyle } from './types';

export class StylePool {
  private styles: Map<string, CellStyle> = new Map();
  private styleToId: Map<string, string> = new Map();
  private nextId = 1;

  getOrCreate(style: CellStyle): string {
    const styleKey = this.getStyleKey(style);
    const existingId = this.styleToId.get(styleKey);
    if (existingId) {
      return existingId;
    }

    const id = `style_${this.nextId++}`;
    this.styles.set(id, style);
    this.styleToId.set(styleKey, id);
    return id;
  }

  /** Whether a style with these properties is in the pool already (so `getOrCreate` would not add one). */
  has(style: CellStyle): boolean {
    return this.styleToId.has(this.getStyleKey(style));
  }

  get(styleId: string): CellStyle | undefined {
    return this.styles.get(styleId);
  }

  getStyleKey(style: CellStyle): string {
    // Create a deterministic key from style properties
    const keys = Object.keys(style).sort();
    return keys.map((key) => `${key}:${style[key as keyof CellStyle]}`).join('|');
  }

  clear(): void {
    this.styles.clear();
    this.styleToId.clear();
    this.nextId = 1;
  }

  size(): number {
    return this.styles.size;
  }

  getAllStyles(): Map<string, CellStyle> {
    return new Map(this.styles);
  }

  setStyleToId(styleToId: Map<string, string>): void {
    this.styleToId = styleToId;
  }
  setStyles(styles: Map<string, CellStyle>): void {
    this.styles = styles;
  }

  /**
   * Load an entire serialized style pool at once. Rebuilds both maps and the id
   * counter in a single pass. (Do NOT set entries one at a time via setStyles —
   * that replaces the whole map each call, collapsing the pool to one entry.)
   */
  load(entries: Record<string, CellStyle>): void {
    this.styles = new Map();
    this.styleToId = new Map();
    let maxId = 0;
    for (const [id, style] of Object.entries(entries)) {
      // A peer/imported id could be __proto__/constructor; Map.set is safe.
      if (id === '__proto__' || id === 'constructor' || id === 'prototype') continue;
      this.styles.set(id, style);
      this.styleToId.set(this.getStyleKey(style), id);
      const n = parseInt(id.split('_')[1] || '0', 10);
      if (Number.isFinite(n) && n > maxId) maxId = n;
    }
    this.nextId = maxId + 1;
  }

  getNextId(): number {
    return this.nextId;
  }

  setNextId(nextId: number): void {
    this.nextId = nextId;
  }
}

