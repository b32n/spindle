import { detectPptxLoss, exportPptx, importPptx } from '@b32nio/spindle-io/pptx';
import {
  createApi,
  type PptxEditorApi,
  type PptxModule,
} from '../../../../scripts/pptx-fidelity/api';

/**
 * The harness's editor API bound to the io source (the CLI binds the built package instead). The io reader
 * and writer (`importPptx`, plan K9; `exportPptx`, K12) are wired; slides-core's own `importFromPptx` replaces it when it lands (K15).
 */
export const sourceApi: PptxEditorApi = createApi({ detectPptxLoss, importPptx, exportPptx } as unknown as PptxModule);

/** The same binding without the reader and the writer: what the harness sees against a build that has no importer (`importError: "no importer"`). */
export const detectOnlyApi: PptxEditorApi = createApi({ detectPptxLoss } as unknown as PptxModule);
