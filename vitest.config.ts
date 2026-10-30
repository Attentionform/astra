/// <reference types="vitest" />
import { defineConfig } from 'vite';
import path from 'path';
import os from 'os';

// One worker per 2GB RAM, floor 1, cap 4. A 7.6GB host lands on 3.
const totalMemGb = os.totalmem() / (1024 * 1024 * 1024);
const safeWorkers = Math.max(1, Math.min(4, Math.floor(totalMemGb / 2)));

// Clamp workers to RAM, not core count. Override for large CI hosts.
// Non-numeric, zero, or negative VITEST_MAX_WORKERS values fall back to
// the computed safeWorkers (never 0/NaN); a warning goes to stderr.
function resolveMaxWorkers(): number {
  const raw = process.env.VITEST_MAX_WORKERS;
  if (raw === undefined || raw.trim() === '') return safeWorkers;
  const parsed = Number.parseInt(raw, 10);
  if (!Number.isFinite(parsed) || parsed < 1) {
    process.stderr.write(
      `[vitest] invalid VITEST_MAX_WORKERS=${JSON.stringify(raw)}; falling back to ${safeWorkers}\n`,
    );
    return safeWorkers;
  }
  return parsed;
}

export default defineConfig({
  resolve: {
    alias: {
      '@attentionform/astra-engine': path.resolve(__dirname, 'packages/engine/src'),
      '@attentionform/astra-mcp-server': path.resolve(__dirname, 'packages/mcp-server/src/index.ts'),
      '@attentionform/ooxml-validator': path.resolve(__dirname, 'packages/ooxml-validator/src/index.ts'),
    },
  },
  test: {
    globals: true,
    environment: 'node', // Use node for backend testing, we can override in specific tests
    globalSetup: [path.resolve(__dirname, 'tests/global-setup.ts')],
    setupFiles: [path.resolve(__dirname, 'tests/setup.ts')],
    testTimeout: 30000,
    hookTimeout: 30000,
    include: [
      'tests/unit/**/*.{test,spec}.{js,mjs,cjs,ts,mts,cts,jsx,tsx}',
      'tests/integration/**/*.{test,spec}.{js,mjs,cjs,ts,mts,cts,jsx,tsx}',
      'tests/harness/**/*.{test,spec}.{js,mjs,cjs,ts,mts,cts,jsx,tsx}',
      'packages/*/tests/**/*.{test,spec}.{js,mjs,cjs,ts,mts,cts,jsx,tsx}'
    ],
    // Clamp workers to RAM, not core count. Override for large CI hosts.
    maxWorkers: resolveMaxWorkers(),
    minWorkers: 1,
    exclude: [
      '**/node_modules/**',
      '**/dist/**',
    ],
  },
});
