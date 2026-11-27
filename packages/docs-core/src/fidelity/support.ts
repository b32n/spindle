import { DocumentImpl } from '../document';
import { exportToDocx, importFromDocx, inspectDocx } from '../export/docx';
import { createApi, type DocsCoreModule } from '../../../../scripts/docx-fidelity/api';

/** The harness's editor API bound to docs-core source (the CLI binds the built package instead). */
export const sourceApi = createApi({ DocumentImpl, importFromDocx, exportToDocx, inspectDocx } as unknown as DocsCoreModule);

/** Read a nested value by dotted path (`sections.0.blocks.1.id`) so tests need no `any`. */
export function get(value: unknown, dotted: string): unknown {
  return dotted.split('.').reduce<unknown>((acc, key) => (acc == null ? undefined : (acc as Record<string, unknown>)[key]), value);
}
