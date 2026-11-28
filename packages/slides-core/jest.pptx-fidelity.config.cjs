// Coverage for the PPTX fidelity scripts in /scripts/pptx-fidelity, which live outside this package's rootDir (so the
// default config cannot collect them):  npm run test:fidelity:coverage
// The generic modules in scripts/fidelity-core are covered by the sheets-core suite (npm run test:fidelity:coverage there).
const path = require('node:path');
const base = require('./jest.config.cjs');

// The base config's TypeScript `paths` and module mapping are relative to this package; with rootDir at the repository
// root they must be absolute.
const [pattern, options] = base.transform['^.+\\.ts$'];
const here = (rel) => path.resolve(__dirname, rel);
const paths = Object.fromEntries(
  Object.entries(options.tsconfig.paths).map(([name, targets]) => [name, targets.map(here)])
);
const moduleNameMapper = Object.fromEntries(
  Object.entries(base.moduleNameMapper).map(([re, target]) => [re, target.replace('<rootDir>', __dirname)])
);

module.exports = {
  ...base,
  rootDir: '../..',
  roots: ['<rootDir>/packages/slides-core/src/fidelity'],
  moduleNameMapper,
  transform: { [pattern]: ['ts-jest', { ...options, tsconfig: { ...options.tsconfig, paths } }] },
  collectCoverageFrom: ['scripts/pptx-fidelity/**/*.ts', 'scripts/review-pack/helpers.ts'],
  coverageDirectory: path.resolve(__dirname, 'coverage'),
};
