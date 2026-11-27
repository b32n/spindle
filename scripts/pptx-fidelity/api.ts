import type { Json, LossSnapshot } from '../fidelity-core/types';

/**
 * The editor-level PPTX API, injected so the same harness runs against the sources (Jest) and against the built
 * packages (the CLI) without importing either itself.
 *
 * The loss detector, the reader (`importPptx`, plan K9) and the writer (`exportPptx`, K12) exist. `importPptx` and `exportPptx`
 * stay optional so the harness still runs against a build without them: while the importer is absent it records
 * `importError: "no importer"` for the `model` kind, and while the exporter is absent the round trip records `no exporter`.
 */
export interface PptxEditorApi {
  /** The loss report at the `full` detail level. */
  inspectPptx(bytes: Uint8Array): Promise<LossSnapshot>;
  /** Import into a JSON-friendly presentation model (normalised by `normalize.ts`). */
  importPptx?(bytes: Uint8Array): Promise<{ warnings: string[]; presentation: unknown }>;
  /** Export a model produced by `importPptx`; `warnings` say what the writer left out. */
  exportPptx?(presentation: unknown): Promise<{ bytes: Uint8Array; warnings: string[] }>;
}

/** The parts of `@b32nio/spindle-io/pptx` (and, later, the slides core) the API is built from. */
export interface PptxModule {
  detectPptxLoss(bytes: Uint8Array, options?: { detail: 'full' }): Promise<LossSnapshot>;
  /** The reader: returns the interchange result, whose `warnings` the harness records beside the model. */
  importPptx?: (bytes: Uint8Array) => Promise<{ warnings: string[] }>;
  exportPptx?: PptxEditorApi['exportPptx'];
}

const bindImport =
  (read: NonNullable<PptxModule['importPptx']>): NonNullable<PptxEditorApi['importPptx']> =>
  async (bytes) => {
    const { warnings, ...presentation } = await read(bytes);
    return { warnings, presentation };
  };

/** Bind the harness to a module. The loss detector is asked for the `full` detail level (every loss id, not only the dialog's). */
export function createApi(mod: PptxModule): PptxEditorApi {
  return {
    inspectPptx: (bytes) => mod.detectPptxLoss(bytes, { detail: 'full' }),
    ...(mod.importPptx ? { importPptx: bindImport(mod.importPptx) } : {}),
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
