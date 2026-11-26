// PPTX interchange shapes shared by `@b32nio/spindle-io/pptx` and the slides packages. Only the loss
// detector's options exist so far; the reader and writer types arrive with the importer (plan K9).

/** Options for the loss detector (`detectPptxLoss`). All optional. */
export interface DetectPptxLossOptions {
  /** `'summary'` (default) drives the import-confirmation dialog; `'full'` adds the finer ids the fidelity harness records. */
  detail?: 'summary' | 'full';
}
