import typescript from '@rollup/plugin-typescript';
import { defineConfig } from 'rollup';

// Entry points:
//   .            → framework-agnostic utilities (EventEmitter, collaboration, …)
//   ./react      → React UI helpers (ResponsiveToolbar). React stays external, and
//                  the subpath keeps React out of the pure `.` entry so non-React
//                  consumers (the *-core packages) never pull it in.
//   ./export/csv  → CSV read/write helpers. Each export format gets its own
//   ./export/xlsx   subpath (same reasoning as ./react) so a consumer that
//                    only needs CSV never bundles xlsx/docx/pptx/pdf
//                    machinery, and vice versa.

const tsMain = typescript({
  tsconfig: './tsconfig.json',
  declaration: true,
  declarationDir: './dist',
  rootDir: './src',
  outputToFilesystem: true,
});
const tsReact = typescript({
  tsconfig: './tsconfig.json',
  declaration: true,
  declarationDir: './dist',
  rootDir: './src',
  outputToFilesystem: true,
});
const tsExportCsv = typescript({
  tsconfig: './tsconfig.json',
  declaration: true,
  declarationDir: './dist',
  rootDir: './src',
  outputToFilesystem: true,
});
const tsExportXlsx = typescript({
  tsconfig: './tsconfig.json',
  declaration: true,
  declarationDir: './dist',
  rootDir: './src',
  outputToFilesystem: true,
});

export default defineConfig([
  {
    input: 'src/index.ts',
    output: [
      { file: 'dist/index.js', format: 'cjs', sourcemap: true },
      { file: 'dist/index.esm.js', format: 'esm', sourcemap: true },
    ],
    external: [],
    plugins: [tsMain],
  },
  {
    input: 'src/react/index.ts',
    output: [
      { file: 'dist/react/index.js', format: 'cjs', sourcemap: true },
      { file: 'dist/react/index.esm.js', format: 'esm', sourcemap: true },
    ],
    // Externalize every bare import (react / react-dom); bundle only our source.
    external: (id) => !/^[./]/.test(id),
    plugins: [tsReact],
  },
  {
    input: 'src/export/csv/index.ts',
    output: [
      { file: 'dist/export/csv/index.js', format: 'cjs', sourcemap: true },
      { file: 'dist/export/csv/index.esm.js', format: 'esm', sourcemap: true },
    ],
    external: [],
    plugins: [tsExportCsv],
  },
  {
    input: 'src/export/xlsx/index.ts',
    output: [
      { file: 'dist/export/xlsx/index.js', format: 'cjs', sourcemap: true },
      { file: 'dist/export/xlsx/index.esm.js', format: 'esm', sourcemap: true },
    ],
    external: ['exceljs'],
    plugins: [tsExportXlsx],
  },
]);
