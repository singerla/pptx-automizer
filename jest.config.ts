module.exports = {
  preset: 'ts-jest',
  testEnvironment: 'node',
  // helpers/ must not be collected as suites
  testMatch: ['**/__tests__/**/*.test.ts'],
  // Tier-1 package invariants run on every archive written by any test
  setupFilesAfterEnv: ['<rootDir>/__tests__/helpers/setup-pptx-invariants.ts'],
  collectCoverageFrom: ['src/**/*.ts', '!src/dev.ts'],
};
