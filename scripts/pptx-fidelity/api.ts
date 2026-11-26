import type { Json, LossSnapshot } from '../fidelity-core/types';

/**
 * The editor-level PPTX API, injected so the same harness runs against the sources (Jest) and against the built
 * packages (the CLI) without importing either itself.
 *
 * Only the loss detector exists today. `importPptx` and `exportPptx` are optional on purpose: while they are absent the
 * harness records `importError: "no importer"` for the `model` kind and skips the round trip. When the importer lands
 * (plan K9 to K15) `createApi` is given the functions and the same goldens start to hold models.
 */
export interface PptxEditorApi {
  /** The loss report at the `full` detail level. */
  inspectPptx(bytes: Uint8Array): Promise<LossSnapshot>;
  /** Import into a JSON-friendly presentation model (normalised by `normalize.ts`). */
  importPptx?(bytes: Uint8Array): Promise<{ warnings: string[]; presentation: unknown }>;
  /** Export a model produced by `importPptx`. */
  exportPptx?(presentation: unknown): Promise<Uint8Array>;
}

/** The parts of `@b32nio/spindle-io/pptx` (and, later, the slides core) the API is built from. */
export interface PptxModule {
  detectPptxLoss(bytes: Uint8Array, options?: { detail: 'full' }): Promise<LossSnapshot>;
  importPptx?: PptxEditorApi['importPptx'];
  exportPptx?: PptxEditorApi['exportPptx'];
}

/** Bind the harness to a module. The loss detector is asked for the `full` detail level (every loss id, not only the dialog's). */
export function createApi(mod: PptxModule): PptxEditorApi {
  return {
    inspectPptx: (bytes) => mod.detectPptxLoss(bytes, { detail: 'full' }),
    ...(mod.importPptx ? { importPptx: mod.importPptx } : {}),
    ...(mod.exportPptx ? { exportPptx: mod.exportPptx } : {}),
  };
}

/** The built `@b32nio/spindle-io/pptx` (run `npm run build` first). */
export async function loadBuiltApi(): Promise<PptxEditorApi> {
  return createApi((await import('@b32nio/spindle-io/pptx')) as unknown as PptxModule);
}

/** What a model snapshot holds when there is no importer. */
export const NO_IMPORTER = 'no importer';

export type { Json };
