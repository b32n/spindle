import typescript from '@rollup/plugin-typescript';
import { defineConfig } from 'rollup';

// Two entry points:
//   .          → core slides engine (no Yjs dep for non-collab consumers)
//   ./collab   → Yjs-backed collaboration binding
//
// Only the main entry emits .d.ts files (tsc walks every reachable source
// file when generating declarations, so dist/collab/*.d.ts come out of the
// main pass). The collab pass uses declaration: false to avoid clobbering
// dist/collab/index.d.ts with the main entry's types.

const tsPluginMain = typescript({
  tsconfig: './tsconfig.json',
  declaration: true,
  declarationDir: './dist',
  rootDir: './src',
  outputToFilesystem: true,
});
const tsPluginNoDecl = typescript({
  tsconfig: './tsconfig.json',
  declaration: false,
  declarationMap: false,
  // tsconfig has composite: true, which requires declaration emit; disable it
  // here too so this pass doesn't conflict with declaration: false above (TS6304).
  composite: false,
  outputToFilesystem: true,
});

const external = [
  '@b32nio/spindle-shared',
  'yjs',
  'y-protocols',
  'y-protocols/awareness',
  'y-protocols/sync',
  'y-indexeddb',
  'y-prosemirror',
  'prosemirror-model',
  'prosemirror-state',
  'prosemirror-transform',
  'lib0/encoding',
  'lib0/decoding',
];

export default defineConfig([
  {
    // The base engine dynamically imports ./collab/binding (the only Yjs
    // importer) from attachCollab(), so it code-splits into its own chunk that
    // consumers only load when they opt into collaboration. That keeps Yjs out
    // of the base bundle — hence output.dir (code-splitting) instead of a
    // single output.file.
    input: 'src/index.ts',
    output: [
      { dir: 'dist', format: 'cjs', entryFileNames: 'index.js', chunkFileNames: 'chunks/[name]-[hash].js', sourcemap: true },
      { dir: 'dist', format: 'esm', entryFileNames: 'index.esm.js', chunkFileNames: 'chunks/[name]-[hash].esm.js', sourcemap: true },
    ],
    external,
    plugins: [tsPluginMain],
  },
  {
    input: 'src/collab/index.ts',
    output: [
      { file: 'dist/collab/index.js', format: 'cjs', sourcemap: true },
      { file: 'dist/collab/index.esm.js', format: 'esm', sourcemap: true },
    ],
    external,
    plugins: [tsPluginNoDecl],
  },
]);
