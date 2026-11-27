/** @type {import('jest').Config} */
module.exports = {
  preset: 'ts-jest',
  testEnvironment: 'node',
  roots: ['<rootDir>/src'],
  // The retired `exceljs` reader (S2b): unchanged apart from its name, it only runs as the oracle of reader-differential.test.ts and is deleted
  // with `exceljs` in S14, so measuring it would only report on code nobody maintains. Listed here, not hidden by a comment in the file.
  coveragePathIgnorePatterns: ['/node_modules/', '<rootDir>/src/xlsx/reader-exceljs\\.ts$'],
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
            '@b32nio/spindle-shared/export/pptx': ['../shared/src/export/pptx/index.ts'],
          },
        },
      },
    ],
  },
  // Jest's own resolver reads the exports map (built dist), but tests should
  // run against source, like the other packages' path mappings.
  moduleNameMapper: {
    '^@b32nio/spindle-shared$': '<rootDir>/../shared/src/index.ts',
    '^@b32nio/spindle-shared/export/(csv|xlsx|docx|pptx)$': '<rootDir>/../shared/src/export/$1/index.ts',
  },
};
