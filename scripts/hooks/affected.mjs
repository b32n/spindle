// lint-staged helper: for the staged source files passed as arguments, run the
// type-check and the *related* unit tests of each affected workspace package.
// Usage (from lint-staged): node scripts/hooks/affected.mjs <files...>
import { spawnSync } from 'node:child_process';
import path from 'node:path';

const root = process.cwd();
const byPackage = new Map();
for (const file of process.argv.slice(2)) {
  const rel = path.relative(root, path.resolve(file)).split(path.sep);
  if (rel[0] !== 'packages' || !rel[1]) continue;
  const list = byPackage.get(rel[1]) ?? [];
  list.push(path.resolve(file));
  byPackage.set(rel[1], list);
}

let failed = false;
const run = (cmd, args, cwd) => {
  const r = spawnSync(cmd, args, { cwd, stdio: 'inherit', shell: process.platform === 'win32' });
  if (r.status !== 0) failed = true;
};

for (const [pkg, files] of byPackage) {
  const cwd = path.join(root, 'packages', pkg);
  console.log(`\n[hooks] ${pkg}: type-check`);
  run('npx', ['tsc', '--noEmit'], cwd);
  console.log(`[hooks] ${pkg}: related unit tests`);
  run('npx', ['jest', '--passWithNoTests', '--findRelatedTests', ...files], cwd);
}
process.exit(failed ? 1 : 0);
