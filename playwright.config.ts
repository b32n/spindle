import { defineConfig, devices } from '@playwright/test';

const CI = !!process.env.CI;

// Fixed ports so tests can derive URLs; chosen to avoid the demos' dev ports.
export const PORTS = { sheets: 5175, docs: 5176, slides: 5177, collab: 1234 } as const;

export default defineConfig({
  testDir: './e2e',
  outputDir: './e2e/.results',
  fullyParallel: true,
  forbidOnly: CI,
  retries: CI ? 2 : 0,
  workers: CI ? 2 : undefined,
  reporter: CI
    ? [['blob'], ['github']]
    : [['list'], ['html', { open: 'never', outputFolder: 'e2e/.report' }]],
  timeout: 30_000,
  expect: { timeout: 7_000 },
  use: {
    trace: 'retain-on-failure',
    screenshot: 'only-on-failure',
    video: 'retain-on-failure',
    permissions: ['clipboard-read', 'clipboard-write'],
    viewport: { width: 1280, height: 800 },
  },
  projects: [
    { name: 'chromium', use: { ...devices['Desktop Chrome'] } },
    {
      name: 'firefox',
      // Firefox/WebKit do not support the clipboard-read/write permissions.
      use: { ...devices['Desktop Firefox'], permissions: [] },
    },
    { name: 'webkit', use: { ...devices['Desktop Safari'], permissions: [] } },
  ],
  webServer: [
    {
      command: `npx vite examples/sheets-demo --port ${PORTS.sheets} --strictPort --host 127.0.0.1`,
      url: `http://127.0.0.1:${PORTS.sheets}`,
      reuseExistingServer: !CI,
    },
    {
      command: `npx vite examples/docs-demo --port ${PORTS.docs} --strictPort --host 127.0.0.1`,
      url: `http://127.0.0.1:${PORTS.docs}`,
      reuseExistingServer: !CI,
    },
    {
      command: `npx vite examples/slides-demo --port ${PORTS.slides} --strictPort --host 127.0.0.1`,
      url: `http://127.0.0.1:${PORTS.slides}`,
      reuseExistingServer: !CI,
    },
    {
      command: `npx tsx examples/collab-server/src/server.ts`,
      env: { PORT: String(PORTS.collab) },
      // WebSocket server: HTTP upgrade-only, plain GET returns 426 or similar.
      port: PORTS.collab,
      reuseExistingServer: !CI,
    },
  ],
});
