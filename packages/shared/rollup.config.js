import typescript from '@rollup/plugin-typescript';
import { defineConfig } from 'rollup';

// Two entry points:
//   .        → framework-agnostic utilities (EventEmitter, collaboration, …)
//   ./react  → React UI helpers (ResponsiveToolbar). React stays external, and
//              the subpath keeps React out of the pure `.` entry so non-React
//              consumers (the *-core packages) never pull it in.

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
]);
