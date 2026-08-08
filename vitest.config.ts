import { defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    // NOTE: `.spec.ts`, matching outlook-access (the source of the extracted
    // SharePoint code). teams-access uses `.test.ts`. The two existing repos
    // are already inconsistent; this one picks a suffix and enforces it here.
    include: ['test_scripts/**/*.spec.ts'],
    environment: 'node',
    globals: false,
    reporters: ['default'],
    coverage: {
      provider: 'v8',
      reporter: ['text', 'lcov'],
      include: ['src/**/*.ts'],
      exclude: ['src/cli.ts', 'src/auth/capture.ts'],
    },
  },
});
