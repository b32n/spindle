import { detectPptxLoss } from '@b32nio/spindle-io/pptx';
import {
  createApi,
  type PptxEditorApi,
  type PptxModule,
} from '../../../../scripts/pptx-fidelity/api';

/**
 * The harness's editor API bound to the io source (the CLI binds the built package instead). slides-core has no PPTX
 * importer yet, so only the loss detector is wired; `importPptx` appears here when the importer lands.
 */
export const sourceApi: PptxEditorApi = createApi({ detectPptxLoss } as unknown as PptxModule);
