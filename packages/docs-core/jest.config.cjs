/** @type {import('jest').Config} */
module.exports = {
  preset: 'ts-jest',
  testEnvironment: 'node',
  roots: ['<rootDir>/src'],
  // The fidelity runner (scripts/docx-fidelity) imports the built package by name; tests bind the source instead.
  moduleNameMapper: { '^@b32nio/spindle-docs-core$': '<rootDir>/src/index.ts' },
  transform: {
    '^.+\\.ts$': [
      'ts-jest',
      {
        tsconfig: {
          module: 'commonjs',
          moduleResolution: 'node',
          esModuleInterop: true,
          verbatimModuleSyntax: false,
          composite: false,
          declaration: false,
          noUnusedLocals: false,
          noUnusedParameters: false,
          // Classic 'node' resolution above doesn't understand package.json
          // "exports" maps (see sheets-core/jest.config.cjs for the same fix).
          paths: {
            '@b32nio/spindle-shared/export/docx': ['../shared/src/export/docx/index.ts'],
            '@b32nio/spindle-io/docx': ['../io/src/docx/index.ts'],
            '@b32nio/spindle-docs-core': ['./src/index.ts'],
          },
        },
      },
    ],
  },
};
