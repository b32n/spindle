import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import path from 'path';

export default defineConfig({
  plugins: [react()],
  resolve: {
    alias: {
      '@b32nio/spindle-sheets-core': path.resolve(__dirname, '../../packages/sheets-core/src'),
      '@b32nio/spindle-sheets-react': path.resolve(__dirname, '../../packages/sheets-react/src'),
      '@b32nio/spindle-shared': path.resolve(__dirname, '../../packages/shared/src'),
      '@b32nio/spindle-transport-websocket': path.resolve(__dirname, '../../packages/transport-websocket/src'),
    },
  },
  server: {
    port: 5175,
  },
});
