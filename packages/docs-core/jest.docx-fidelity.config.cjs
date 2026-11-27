// Coverage for the DOCX fidelity script in /scripts (docx-fidelity; the shared fidelity-core is covered by the sheets-core suite), which lives outside this
// package's rootDir (so the default config cannot collect them):  npm run test:fidelity:coverage
const path = require('node:path');
const base = require('./jest.config.cjs');

// The base config's TypeScript `paths` are relative to this package; with rootDir at the repository root they must
// be absolute.
const [pattern, options] = base.transform['^.+\\.ts$'];
const here = (rel) => path.resolve(__dirname, rel);
const paths = Object.fromEntries(
  Object.entries(options.tsconfig.paths).map(([name, targets]) => [name, targets.map(here)])
);

module.exports = {
  ...base,
  rootDir: '../..',
  roots: ['<rootDir>/packages/docs-core/src/fidelity'],
  moduleNameMapper: { '^@b32nio/spindle-docs-core$': here('./src/index.ts') },
  transform: { [pattern]: ['ts-jest', { ...options, tsconfig: { ...options.tsconfig, paths } }] },
  collectCoverageFrom: ['scripts/docx-fidelity/**/*.ts'],
  coverageDirectory: '<rootDir>/packages/docs-core/coverage',
};
