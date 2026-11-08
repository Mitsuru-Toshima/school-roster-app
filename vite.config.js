import { defineConfig } from 'vite';

export default defineConfig({
  root: './',
  server: {
    headers: {
      'Cross-Origin-Opener-Policy': 'same-origin-allow-popups'
    }
  },
  build: {
    outDir: 'dist'
  }
});
