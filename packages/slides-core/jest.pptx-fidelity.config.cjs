// Coverage for the PPTX fidelity scripts in /scripts/pptx-fidelity, which live outside this package's rootDir (so the
// default config cannot collect them):  npm run test:fidelity:coverage
// The generic modules in scripts/fidelity-core are covered by the sheets-core suite (npm run test:fidelity:coverage there).
const path = require('node:path');
const base = require('./jest.config.cjs');

module.exports = {
  ...base,
  rootDir: '../..',
  roots: ['<rootDir>/packages/slides-core/src/fidelity'],
  collectCoverageFrom: ['scripts/pptx-fidelity/**/*.ts'],
  coverageDirectory: path.resolve(__dirname, 'coverage'),
};
