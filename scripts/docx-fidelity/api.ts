import type { DocumentLike, EditorApi } from './types';

/** The parts of `@b32nio/spindle-docs-core` the API is built from. */
export interface DocsCoreModule {
  DocumentImpl: new (id: string, name: string) => DocumentLike;
  importFromDocx(bytes: Uint8Array, document: never): ReturnType<EditorApi['importFromDocx']>;
  exportToDocx(document: never): ReturnType<EditorApi['exportToDocx']>;
  inspectDocx(bytes: Uint8Array, options?: { detail: 'full' }): ReturnType<EditorApi['inspectDocx']>;
}

/** Bind the harness to a docs-core module. `inspectDocx` is asked for the `full` detail level (every loss id, not only the dialog's). */
export function createApi(core: DocsCoreModule): EditorApi {
  return {
    createDocument: () => new core.DocumentImpl('fixture', 'Untitled'),
    importFromDocx: (bytes, document) => core.importFromDocx(bytes, document as never),
    exportToDocx: (document) => core.exportToDocx(document as never),
    inspectDocx: (bytes) => core.inspectDocx(bytes, { detail: 'full' }),
  };
}

/** The built `@b32nio/spindle-docs-core` (run `npm run build` first). */
export async function loadBuiltApi(): Promise<EditorApi> {
  return createApi((await import('@b32nio/spindle-docs-core')) as unknown as DocsCoreModule);
}
