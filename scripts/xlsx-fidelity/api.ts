import type { LossSnapshot } from '../fidelity-core/types';

/** The slice of a sheets-core `WorkbookImpl` the harness needs. */
export interface WorkbookLike {
  getData(): unknown;
}

/**
 * The editor-level XLSX API, injected so the same harness runs against sheets-core source (Jest) and
 * against its built output (the CLI) without importing either itself.
 */
export interface XlsxEditorApi {
  createWorkbook(): WorkbookLike;
  importFromXlsx(
    bytes: Uint8Array,
    workbook: WorkbookLike
  ): Promise<{ warnings: string[]; loss: LossSnapshot }>;
  exportToXlsx(workbook: WorkbookLike): Promise<Uint8Array>;
  inspectXlsx(bytes: Uint8Array): Promise<LossSnapshot>;
  /** Whether the formula engine can parse this formula text. */
  parsesFormula(formula: string): boolean;
}

/** The parts of `@b32nio/spindle-sheets-core` the API is built from. */
export interface SheetsCoreModule {
  WorkbookImpl: new (id: string, name: string) => WorkbookLike;
  importFromXlsx(
    bytes: Uint8Array,
    workbook: never
  ): Promise<{ warnings: string[]; loss: LossSnapshot }>;
  exportToXlsx(workbook: never): Promise<Uint8Array>;
  inspectXlsx(bytes: Uint8Array, options?: { detail: 'full' }): Promise<LossSnapshot>;
  FormulaParser: new () => { parse(formula: string): { error?: string } };
}

/** Bind the harness to a sheets-core module. `inspectXlsx` is asked for the `full` detail level (every loss id, not only the dialog's). */
export function createApi(core: SheetsCoreModule): XlsxEditorApi {
  const parser = new core.FormulaParser();
  return {
    createWorkbook: () => new core.WorkbookImpl('fixture', 'Untitled'),
    importFromXlsx: (bytes, workbook) => core.importFromXlsx(bytes, workbook as never),
    exportToXlsx: (workbook) => core.exportToXlsx(workbook as never),
    inspectXlsx: (bytes) => core.inspectXlsx(bytes, { detail: 'full' }),
    parsesFormula: (formula) => parser.parse(formula).error === undefined,
  };
}

/** The built `@b32nio/spindle-sheets-core` (run `npm run build` first). */
export async function loadBuiltApi(): Promise<XlsxEditorApi> {
  return createApi((await import('@b32nio/spindle-sheets-core')) as unknown as SheetsCoreModule);
}
