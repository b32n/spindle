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
          // "exports" maps, so it can't find a subpath import like
          // @b32nio/spindle-shared/export/csv (only real for type-checking —
          // ts-jest emits a plain `require(...)` either way, which Jest's own
          // resolver finds fine via that same exports map at runtime).
          paths: {
            '@b32nio/spindle-shared/export/csv': ['../shared/src/export/csv/index.ts'],
          },
        },
      },
    ],
  },
};
