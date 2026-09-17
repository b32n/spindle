import typescript from '@rollup/plugin-typescript';
import { defineConfig } from 'rollup';

// Two entry points:
//   .          → core document engine (no Yjs dep pulled in for consumers
//                who don't use collab).
//   ./collab   → Yjs-backed collaboration binding. Pulls in yjs / y-protocols.

// Only the main entry emits .d.ts files — tsc walks every reachable source
// file and writes declarations under ./dist/, including dist/collab/*.d.ts.
// If the collab entry also emitted declarations, it would overwrite
// dist/collab/index.d.ts with the wrong content (the main entry's types).
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
  'y-indexeddb',
  'y-prosemirror',
  'prosemirror-model',
  'prosemirror-state',
  'prosemirror-transform',
  'prosemirror-schema-basic',
  'prosemirror-schema-list',
  'prosemirror-keymap',
  'prosemirror-history',
  'prosemirror-commands',
  'prosemirror-inputrules',
  'y-protocols/sync',
  'lib0/encoding',
  'lib0/decoding',
];

export default defineConfig([
  {
    input: 'src/index.ts',
    output: [
      { file: 'dist/index.js', format: 'cjs', sourcemap: true },
      { file: 'dist/index.esm.js', format: 'esm', sourcemap: true },
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
