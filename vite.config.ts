import { defineConfig } from 'vitest/config';

// Library build (`npm run build`) and unit tests (`npm test`).
// The demo app has its own config in demo/vite.config.ts.
export default defineConfig({
  build: {
    sourcemap: true,
    cssCodeSplit: false,
    lib: {
      entry: { 'xlsx-editor': 'src/index.ts', extension: 'src/extension.ts' },
      formats: ['es'],
      fileName: (_format, name) => `${name}.js`,
      cssFileName: 'style',
    },
    rollupOptions: {
      external: ['react', 'react/jsx-runtime'],
    },
  },
  test: {
    // happy-dom provides DOMParser for the xlsx reader tests; everything else runs in plain Node.
    environment: 'happy-dom',
    include: ['src/**/*.test.{ts,tsx}'],
  },
});
