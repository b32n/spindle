/** Generic fidelity-harness types: format-neutral, shared by every format adapter (xlsx today, pptx later). */

export interface ManifestSource {
  /** Tool that produced the file (e.g. "LibreOffice headless" or "programmatic: hand-authored OOXML"). */
  tool: string;
  /** Version of the tool where one exists and matters (producer-variety files). */
  version?: string;
  os: string;
}

export interface ManifestEntry {
  id: string;
  filename: string;
  title: string;
  source: ManifestSource;
  license: string;
  /** Construct ids from the format's registry. Each one is verified against the file's package. */
  constructs: string[];
  notes: string;
  /** Script that regenerates the file, relative to the fixtures directory. */
  generator?: string;
  /** The file is deliberately broken: the importer is expected to throw. */
  expectImportError?: boolean;
  /** Too large to commit: a generator writes it on demand and the file-exists check skips it. */
  generatedAtRuntime?: boolean;
}

export interface Manifest {
  version: number;
  fixtures: ManifestEntry[];
}

/** A construct probe over a parsed package of type `P`. */
export interface Construct<P> {
  id: string;
  label: string;
  /** How many occurrences of the construct the package contains (0 = absent). */
  count: (p: P) => number;
}

export interface ManifestOptions {
  /** Lower-case file extensions (with the dot) a fixture may use. */
  extensions: string[];
  /** The registry vocabulary. */
  constructIds: string[];
  /** Used in error messages, e.g. "XLSX". */
  label: string;
}

/** Plain JSON value; every snapshot is built from these. */
export type Json = null | boolean | number | string | Json[] | { [key: string]: Json };
export type JsonObject = { [key: string]: Json };

export interface LossItem {
  id: string;
  label: string;
  count: number;
  where?: string[];
  effect: string;
}

export interface LossSnapshot {
  format: string;
  items: LossItem[];
}

export type SnapshotStatus = 'match' | 'mismatch' | 'missing' | 'created' | 'updated' | 'unchecked';
/** Statuses that do not fail a run (`unchecked` = a run-time file that has no golden by design). */
export const GOOD_STATUSES: readonly SnapshotStatus[] = [
  'match',
  'created',
  'updated',
  'unchecked',
];

export interface DiffEntry {
  path: string;
  kind: 'added' | 'removed' | 'changed';
  before?: string;
  after?: string;
}

/** Size counters of a model (sheets, cells, ...), supplied by the format adapter. */
export type ModelCounts = { [name: string]: number };

export interface DiffReport {
  /** Number of differences between the imported model and the re-imported exported model. */
  total: number;
  added: number;
  removed: number;
  changed: number;
  /** Difference counts per field, with indexes (and anything the adapter folds) stripped. */
  byField: { [field: string]: number };
  before: ModelCounts;
  after: ModelCounts;
  /** The first differences in document order (`truncated` when there are more than listed). */
  entries: DiffEntry[];
  truncated: boolean;
}

export type RoundtripStatus = 'ok' | 'skipped-import-error' | 'export-error' | 'reimport-error';

export interface RoundtripSnapshot {
  status: RoundtripStatus;
  error?: string;
  exportWarnings: string[];
  reimportWarnings: string[];
  diff: DiffReport | null;
}

/** Totals of a package ledger (every part of the input package, classified). */
export interface LedgerTotals {
  mapped: number;
  opaque: number;
  dropped: number;
  unaccounted: number;
  total: number;
}

/** Outcome of comparing the engine's recalculated values with the file's cached values. */
export interface ValuesTotals {
  compared: number;
  matched: number;
  mismatched: number;
}

/** How an import result relates to the manifest's `expectImportError`. */
export type ImportExpectation = 'as-expected' | 'unexpected-error' | 'unexpected-success';
