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
          // Classic 'node' resolution doesn't understand package.json "exports"
          // maps, so resolve the shared subpaths to source for type-checking.
          paths: {
            '@b32nio/spindle-shared': ['../shared/src/index.ts'],
            '@b32nio/spindle-shared/export/csv': ['../shared/src/export/csv/index.ts'],
            '@b32nio/spindle-shared/export/xlsx': ['../shared/src/export/xlsx/index.ts'],
            '@b32nio/spindle-shared/export/docx': ['../shared/src/export/docx/index.ts'],
          },
        },
      },
    ],
  },
  // Jest's own resolver reads the exports map (built dist), but tests should
  // run against source, like the other packages' path mappings.
  moduleNameMapper: {
    '^@b32nio/spindle-shared$': '<rootDir>/../shared/src/index.ts',
    '^@b32nio/spindle-shared/export/(csv|xlsx|docx)$': '<rootDir>/../shared/src/export/$1/index.ts',
  },
};
