/** @type {import('jest').Config} */
module.exports = {
  preset: 'ts-jest',
  testEnvironment: 'node',
  roots: ['<rootDir>/src'],
  transform: {
    '^.+\\.tsx?$': [
      'ts-jest',
      {
        tsconfig: {
          module: 'commonjs',
          moduleResolution: 'node',
          jsx: 'react-jsx',
          esModuleInterop: true,
          verbatimModuleSyntax: false,
          composite: false,
          declaration: false,
          noUnusedLocals: false,
          noUnusedParameters: false,
          // Classic 'node' resolution doesn't understand package.json "exports"
          // maps, so type-check subpath imports against the source; Jest's own
          // resolver handles the runtime require via the same map.
          paths: {
            '@b32nio/spindle-shared/react': ['../shared/src/react/index.ts'],
          },
        },
      },
    ],
  },
};
