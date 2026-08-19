/**
 * vite.config.ts: build + test configuration.
 *
 * `base: './'` because the add-in is served from a GitHub Pages *project* path
 * (https://nuiaz.github.io/outlook-crm-addin/), so asset URLs must be relative.
 * That same relative base is what lets the Office manifest point at the Pages
 * URL and have the taskpane load inside Outlook with no server of our own.
 */
import { defineConfig } from 'vitest/config';
import react from '@vitejs/plugin-react';

export default defineConfig({
  plugins: [react()],
  base: './',
  build: { sourcemap: true, outDir: 'dist' },
  test: {
    // jsdom gives window, localStorage and DOM events. It does NOT give
    // `Office`; tests install their own stub (tests/office-stub.ts) so the
    // real host-detection and test-mode fallback paths are exercised.
    environment: 'jsdom',
    globals: true,
    setupFiles: ['./vitest.setup.ts'],
    include: ['tests/**/*.test.ts', 'tests/**/*.test.tsx'],
  },
});
