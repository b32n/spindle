// Optional pre-push gate: runs the @smoke Playwright subset when
// SPINDLE_PREPUSH_E2E=1 is set. Off by default so pushes stay fast.
//   SPINDLE_PREPUSH_E2E=1 git push
import { spawnSync } from 'node:child_process';

if (process.env.SPINDLE_PREPUSH_E2E !== '1') process.exit(0);
console.log('[hooks] pre-push: running E2E smoke (SPINDLE_PREPUSH_E2E=1)');
const r = spawnSync('npm', ['run', 'test:e2e:smoke', '--', '--project=chromium'], {
  stdio: 'inherit',
  shell: process.platform === 'win32',
});
process.exit(r.status ?? 1);
