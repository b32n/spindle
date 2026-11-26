/** @type {import('jest').Config} */
module.exports = {
  preset: 'ts-jest',
  testEnvironment: 'node',
  roots: ['<rootDir>/src'],
  // The PPTX fidelity runner (scripts/pptx-fidelity) imports the built io package by name; tests bind the source instead.
  moduleNameMapper: {
    '^@b32nio/spindle-io/pptx$': '<rootDir>/../io/src/pptx/index.ts',
    '^@b32nio/spindle-shared$': '<rootDir>/../shared/src/index.ts',
    '^@b32nio/spindle-shared/export/pptx$': '<rootDir>/../shared/src/export/pptx/index.ts',
  },
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
          // Classic 'node' resolution doesn't understand package.json "exports" maps, so the workspace subpaths the
          // fidelity runner and the io source import are resolved to source for type-checking.
          paths: {
            '@b32nio/spindle-io/pptx': ['../io/src/pptx/index.ts'],
            '@b32nio/spindle-shared': ['../shared/src/index.ts'],
            '@b32nio/spindle-shared/export/pptx': ['../shared/src/export/pptx/index.ts'],
          },
        },
      },
    ],
  },
};
