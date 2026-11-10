import typescript from '@rollup/plugin-typescript';
import { defineConfig } from 'rollup';

// Entry points:
//   .        → everything (xlsx + docx + csv re-exports)
//   ./xlsx   → XLSX only (needs exceljs)
//   ./docx   → DOCX only (needs docx, jszip, fast-xml-parser)
// Every bare import (exceljs, docx, @b32nio/spindle-shared, …) stays external;
// only our own source is bundled. Hosts and core packages import the subpaths
// so a consumer that needs one format never loads the other format's libraries.

const ts = () =>
  typescript({
    tsconfig: './tsconfig.json',
    declaration: true,
    declarationDir: './dist',
    rootDir: './src',
    outputToFilesystem: true,
  });

const external = (id) => !/^[./]/.test(id);

const entry = (input, dir) => ({
  input,
  output: [
    { file: `dist/${dir}index.js`, format: 'cjs', sourcemap: true },
    { file: `dist/${dir}index.esm.js`, format: 'esm', sourcemap: true },
  ],
  external,
  plugins: [ts()],
});

export default defineConfig([
  entry('src/index.ts', ''),
  entry('src/xlsx/index.ts', 'xlsx/'),
  entry('src/docx/index.ts', 'docx/'),
]);
