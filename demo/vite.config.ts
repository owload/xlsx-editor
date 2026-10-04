import { fileURLToPath } from 'node:url';
import react from '@vitejs/plugin-react';
import { defineConfig } from 'vite';

// Demo / manual-testing app: mounts the editor straight from the sources.
export default defineConfig({
  root: fileURLToPath(new URL('.', import.meta.url)),
  plugins: [react()],
  resolve: {
    alias: {
      '@owload/xlsx-editor': fileURLToPath(new URL('../src/index.ts', import.meta.url)),
    },
  },
});
