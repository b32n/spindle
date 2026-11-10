import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import path from 'node:path';

export default defineConfig({
  plugins: [react()],
  resolve: {
    alias: {
      '@b32nio/spindle-sheets-core': path.resolve(import.meta.dirname, '../../packages/sheets-core/src'),
      '@b32nio/spindle-sheets-react': path.resolve(import.meta.dirname, '../../packages/sheets-react/src'),
      '@b32nio/spindle-shared': path.resolve(import.meta.dirname, '../../packages/shared/src'),
      '@b32nio/spindle-io': path.resolve(import.meta.dirname, '../../packages/io/src'),
      '@b32nio/spindle-transport-websocket': path.resolve(import.meta.dirname, '../../packages/transport-websocket/src'),
    },
  },
  server: {
    port: 5175,
  },
});
