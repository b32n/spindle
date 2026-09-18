/** @type {import('jest').Config} */
module.exports = {
  preset: 'ts-jest',
  testEnvironment: 'node',
  roots: ['<rootDir>/src'],
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
          },
        },
      },
    ],
  },
};
