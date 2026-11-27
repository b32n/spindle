/** @type {import('jest').Config} */
module.exports = {
  preset: 'ts-jest',
  testEnvironment: 'node',
  roots: ['<rootDir>/src'],
  // The shared package's subpath export is not visible to node-style resolution.
  moduleNameMapper: { '^@b32nio/spindle-shared/react$': '<rootDir>/../shared/dist/react/index.js' },
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
          baseUrl: '.',
          paths: { '@b32nio/spindle-shared/react': ['../shared/dist/react/index.d.ts'] },
          composite: false,
          declaration: false,
          noUnusedLocals: false,
          noUnusedParameters: false,
        },
      },
    ],
  },
};
