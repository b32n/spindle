import type { Json, LossSnapshot, RoundtripSnapshot } from '../fidelity-core/types';

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
  importFromDocx(bytes: Uint8Array, document: DocumentLike): Promise<{ warnings: string[]; loss: LossSnapshot }>;
  exportToDocx(document: DocumentLike): Promise<{ bytes: Uint8Array; warnings: string[] }>;
  inspectDocx(bytes: Uint8Array): Promise<LossSnapshot>;
}

export const DOCX_KINDS = ['model', 'loss', 'roundtrip'] as const;
export type DocxKind = (typeof DOCX_KINDS)[number];

export interface ModelSnapshot {
  importError: string | null;
  warnings: string[];
  document: Json | null;
}

export interface FixtureResult {
  loss: LossSnapshot;
  model: ModelSnapshot;
  roundtrip: RoundtripSnapshot;
}
