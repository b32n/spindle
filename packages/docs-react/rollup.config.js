import typescript from '@rollup/plugin-typescript';
import { defineConfig } from 'rollup';

export default defineConfig({
  input: 'src/index.ts',
  output: [
    {
      file: 'dist/index.js',
      format: 'cjs',
      sourcemap: true,
    },
    {
      file: 'dist/index.esm.js',
      format: 'esm',
      sourcemap: true,
    },
  ],
  // Externalize every bare import (all deps/peerDeps); bundle only our own
  // relative/absolute source. Keeps the list from drifting as deps change.
  external: (id) => !/^[./]/.test(id),
  plugins: [
    typescript({
      tsconfig: './tsconfig.json',
      declaration: true,
      declarationDir: './dist',
      rootDir: './src',
      outputToFilesystem: true,
    }),
  ],
});

