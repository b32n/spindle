// Format pool for shared format objects

import type { CellFormat } from './types';

export class FormatPool {
  private formats: Map<string, CellFormat> = new Map();
  private formatToId: Map<string, string> = new Map();
  private nextId = 1;

  getOrCreate(format: CellFormat): string {
    const formatKey = this.getFormatKey(format);
    const existingId = this.formatToId.get(formatKey);
    if (existingId) {
      return existingId;
    }

    const id = `format_${this.nextId++}`;
    this.formats.set(id, format);
    this.formatToId.set(formatKey, id);
    return id;
  }

  /** Whether a format with these properties is in the pool already (so `getOrCreate` would not add one). */
  has(format: CellFormat): boolean {
    return this.formatToId.has(this.getFormatKey(format));
  }

  get(formatId: string): CellFormat | undefined {
    return this.formats.get(formatId);
  }

  getFormatKey(format: CellFormat): string {
    // Create a deterministic key from format properties
    const keys = Object.keys(format).sort();
    return keys.map((key) => `${key}:${format[key as keyof CellFormat]}`).join('|');
  }

  clear(): void {
    this.formats.clear();
    this.formatToId.clear();
    this.nextId = 1;
  }

  size(): number {
    return this.formats.size;
  }

  getAllFormats(): Map<string, CellFormat> {
    return new Map(this.formats);
  }

  setFormatToId(formatToId: Map<string, string>): void {
    this.formatToId = formatToId;
  }
  setFormats(formats: Map<string, CellFormat>): void {
    this.formats = formats;
  }

  /**
   * Load an entire serialized format pool at once (see StylePool.load — setting
   * entries one at a time collapses the pool to a single entry).
   */
  load(entries: Record<string, CellFormat>): void {
    this.formats = new Map();
    this.formatToId = new Map();
    let maxId = 0;
    for (const [id, format] of Object.entries(entries)) {
      if (id === '__proto__' || id === 'constructor' || id === 'prototype') continue;
      this.formats.set(id, format);
      this.formatToId.set(this.getFormatKey(format), id);
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
