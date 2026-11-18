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

export interface LossReport {
  format: string;
  items: LossItem[];
}

/** The slice of a docs-core `DocumentImpl` the harness needs. */
export interface DocumentLike {
  getData(): unknown;
}

/**
 * The editor-level DOCX API, injected so the same harness runs against docs-core source (Jest)
 * and docs-core's built output (CLI) without importing either itself.
 */
export interface EditorApi {
  createDocument(): DocumentLike;
  importFromDocx(bytes: Uint8Array, document: DocumentLike): Promise<{ warnings: string[]; loss: LossReport }>;
  exportToDocx(document: DocumentLike): Promise<{ bytes: Uint8Array; warnings: string[] }>;
  inspectDocx(bytes: Uint8Array): Promise<LossReport>;
}

export interface ManifestSource {
  /** Tool that produced the file, with version (e.g. "LibreOffice 25.2.3.2" or "python-docx 1.2.0 + lxml"). */
  tool: string;
  os: string;
}

export interface ManifestEntry {
  id: string;
  filename: string;
  title: string;
  source: ManifestSource;
  license: string;
  /** Construct ids (see constructs.ts). Each one is verified against the file's OOXML. */
  constructs: string[];
  notes: string;
  /** Script that regenerates the file, relative to the fixtures directory. */
  generator?: string;
  /** The file is deliberately broken: `importFromDocx` is expected to throw. */
  expectImportError?: boolean;
}

export interface Manifest {
  version: number;
  fixtures: ManifestEntry[];
}

export type SnapshotKind = 'model' | 'loss' | 'roundtrip';
export const SNAPSHOT_KINDS: SnapshotKind[] = ['model', 'loss', 'roundtrip'];

export type SnapshotStatus = 'match' | 'mismatch' | 'missing' | 'created' | 'updated';

export interface DiffEntry {
  path: string;
  kind: 'added' | 'removed' | 'changed';
  before?: string;
  after?: string;
}

export interface ModelCounts {
  sections: number;
  blocks: { [type: string]: number };
  textChars: number;
}

export interface DiffReport {
  /** Number of differences between the imported model and the re-imported exported model. */
  total: number;
  added: number;
  removed: number;
  changed: number;
  /** Difference counts per field, indexes stripped (`sections[].blocks[].content[].bold`). */
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

export interface ModelSnapshot {
  importError: string | null;
  warnings: string[];
  document: Json | null;
}

export interface LossSnapshot {
  format: string;
  items: LossItem[];
}

export interface FixtureResult {
  loss: LossSnapshot;
  model: ModelSnapshot;
  roundtrip: RoundtripSnapshot;
}
