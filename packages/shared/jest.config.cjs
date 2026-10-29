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
          // Type-check the optional spindle-io peer against the local stubs (as tsconfig.json does).
          paths: {
            '@b32nio/spindle-io/*': ['./src/export/io-stubs/*'],
          },
        },
      },
    ],
  },
};
