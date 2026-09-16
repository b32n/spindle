import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

// Small React SPA for spindle.b32n.io, deployed to Cloudflare Pages.
export default defineConfig({
  plugins: [react()],
  base: '/',
  server: {
    // Docs content lives in the repo-root `documentation/` folder, one level
    // above this Vite root — allow the dev server to read it.
    fs: { allow: ['..'] },
  },
  build: {
    outDir: 'dist',
    target: 'es2020',
  },
});
