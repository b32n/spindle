import { defineConfig, devices } from '@playwright/test';

const CI = !!process.env.CI;
// PW_TARGET=preview runs the suite against production builds served by `vite preview`
// (build first: `npm run build:demos`) instead of the dev servers.
const PREVIEW = process.env.PW_TARGET === 'preview';
const serve = (demo: string, port: number) =>
  `npx vite ${PREVIEW ? 'preview' : ''} examples/${demo} --port ${port} --strictPort --host 127.0.0.1`;

// Fixed ports so tests can derive URLs; chosen to avoid the demos' dev ports.
export const PORTS = { sheets: 5175, docs: 5176, slides: 5177, collab: 1234 } as const;

// Visual regression (toHaveScreenshot, Chromium only) is opt-in: its baselines come from CI's Linux
// image, so it only runs when PW_VISUAL=1 (the `Visual` CI job, the nightly, `npm run test:visual`).
const VISUAL = process.env.PW_VISUAL === '1';
const VISUAL_PROJECT = {
  name: 'visual',
  testMatch: '**/visual/**/*.spec.ts',
  use: {
    ...devices['Desktop Chrome'],
    viewport: { width: 1280, height: 720 },
    deviceScaleFactor: 1,
  },
};

export default defineConfig({
  testDir: './e2e',
  outputDir: './e2e/.results',
  fullyParallel: true,
  forbidOnly: CI,
  // Retries stay on in CI so a flaky test is still diagnosed (traces/videos of the failed
  // attempt), but `failOnFlakyTests` turns "passed on retry" into a failing run so flakiness
  // is fixed instead of silently absorbed. Override locally/nightly with --retries=0.
  retries: CI ? 2 : 0,
  failOnFlakyTests: CI,
  workers: CI ? 2 : undefined,
  reporter: CI
    ? [['blob'], ['github']]
    : [['list'], ['html', { open: 'never', outputFolder: 'e2e/.report' }]],
  timeout: 30_000,
  expect: {
    timeout: 7_000,
    toHaveScreenshot: { maxDiffPixelRatio: 0.01, animations: 'disabled', caret: 'hide' },
  },
  // One platform-independent baseline directory: only CI's Linux image produces/uses them.
  snapshotPathTemplate: '{testDir}/visual/__screenshots__/{arg}{ext}',
  use: {
    trace: 'retain-on-failure',
    screenshot: 'only-on-failure',
    video: 'retain-on-failure',
    permissions: ['clipboard-read', 'clipboard-write'],
    viewport: { width: 1280, height: 800 },
  },
  projects: [
    ...(VISUAL ? [VISUAL_PROJECT] : []),
    {
      name: 'chromium',
      use: { ...devices['Desktop Chrome'] },
      testIgnore: ['**/mobile/**', '**/visual/**'],
    },
    {
      name: 'firefox',
      testIgnore: ['**/mobile/**', '**/visual/**'],
      // Firefox/WebKit do not support the clipboard-read/write permissions.
      use: { ...devices['Desktop Firefox'], permissions: [] },
    },
    {
      name: 'webkit',
      use: { ...devices['Desktop Safari'], permissions: [] },
      // The docs editor drives a hidden ProseMirror whose selection/focus timing
      // differs in WebKit (most typing scenarios fail there); tracked as a known gap.
      testIgnore: ['**/docs/**', '**/mobile/**', '**/visual/**'],
    },
    // Touch/mobile smoke: a small @mobile set per app (e2e/mobile) on emulated phones. These
    // run in the same shards as the desktop projects, so CI time grows by seconds, not a job.
    {
      name: 'mobile-chrome',
      testMatch: '**/mobile/**/*.spec.ts',
      use: { ...devices['Pixel 7'], permissions: [] },
    },
    {
      name: 'mobile-safari',
      testMatch: '**/mobile/**/*.spec.ts',
      use: { ...devices['iPhone 14'], permissions: [] },
    },
  ],
  webServer: [
    {
      command: serve('sheets-demo', PORTS.sheets),
      url: `http://127.0.0.1:${PORTS.sheets}`,
      reuseExistingServer: !CI,
    },
    {
      command: serve('docs-demo', PORTS.docs),
      url: `http://127.0.0.1:${PORTS.docs}`,
      reuseExistingServer: !CI,
    },
    {
      command: serve('slides-demo', PORTS.slides),
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
