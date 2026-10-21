// `prepare` script: install the git hooks for local development only.
// Never fails npm install: skipped in CI, when SKIP_GIT_HOOKS is set, outside a
// git checkout (e.g. an npm tarball) or if simple-git-hooks isn't installed
// (production install).
import { existsSync } from 'node:fs';
import { spawnSync } from 'node:child_process';

if (process.env.CI || process.env.SKIP_GIT_HOOKS || !existsSync('.git')) process.exit(0);

const r = spawnSync('npx', ['--no-install', 'simple-git-hooks'], {
  stdio: 'inherit',
  shell: process.platform === 'win32',
});
if (r.status !== 0) console.warn('[hooks] simple-git-hooks not installed; skipping git hook setup');
process.exit(0);
