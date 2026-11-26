// Coverage for the XLSX fidelity scripts in /scripts, which live outside this package's rootDir (so the
// default config cannot collect them):  npm run test:fidelity:coverage
const base = require('./jest.config.cjs');

module.exports = {
  ...base,
  rootDir: '../..',
  roots: ['<rootDir>/packages/sheets-core/src/fidelity'],
  collectCoverageFrom: ['scripts/xlsx-fidelity/**/*.ts'],
  coverageDirectory: '<rootDir>/packages/sheets-core/coverage',
};
