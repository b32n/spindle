/** Generic fidelity-harness types (format-neutral; X3 lifts these into scripts/fidelity-core). */

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
