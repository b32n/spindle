/**
 * What an importer cannot bring into Spindle. Importers are lossy by design
 * (Spindle models a subset of OOXML); instead of dropping features silently, a
 * pre-scan of the file returns an ImportLossReport so the host (or the
 * ImportConfirmDialog) can ask before replacing the open document.
 */
export type ImportLossFormat = 'xlsx' | 'docx' | 'csv';

export interface ImportLossItem {
  /** Stable identifier, e.g. `xlsx.conditionalFormatting`. */
  id: string;
  /** Human-readable name of the feature that is lost. */
  label: string;
  /** How many occurrences were found (rules, comments, notes…). */
  count: number;
  /** Where they were found (sheet names…), capped at a few entries. */
  where?: string[];
  /** What happens to it on import, in a short phrase. */
  effect: string;
}

export interface ImportLossReport {
  format: ImportLossFormat;
  /** Empty when the file imports without losing anything Spindle can detect. */
  items: ImportLossItem[];
}

export function createLossReport(format: ImportLossFormat, items: ImportLossItem[]): ImportLossReport {
  return { format, items: items.filter((i) => i.count > 0) };
}

/** True when importing would lose something. */
export function hasImportLoss(report: ImportLossReport): boolean {
  return report.items.length > 0;
}

/** One line per lost feature, e.g. `3 conditional formatting rules — not imported`. */
export function describeImportLoss(report: ImportLossReport): string[] {
  return report.items.map((i) => {
    const where = i.where && i.where.length > 0 ? ` (${i.where.join(', ')})` : '';
    return `${i.count} × ${i.label}${where} — ${i.effect}`;
  });
}

/** Count regex matches (global) in text. */
export function countMatches(text: string, re: RegExp): number {
  return text.match(re)?.length ?? 0;
}
