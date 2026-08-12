/**
 * Tier 3 (ROADMAP Phase 5): visual regression of ~12 curated golden decks,
 * rendered by the pinned pptx-thumbnailer container in tools/render-pptx/.
 *
 * Deliberately a separate config: `yarn test` stays fast and Docker-free;
 * `yarn test:visual` needs Docker only. Visual suites end in `.deck.ts` so
 * the main run's `*.test.ts` glob never collects them.
 *
 * The renderer is a change detector, not a correctness oracle — LibreOffice
 * fidelity is not PowerPoint fidelity. Never conclude PowerPoint-correctness
 * from green pixels.
 */
module.exports = {
  preset: 'ts-jest',
  testEnvironment: 'node',
  testMatch: ['**/__tests__/visual/**/*.deck.ts'],
  globalSetup: '<rootDir>/__tests__/visual/helpers/global-setup.ts',
  globalTeardown: '<rootDir>/__tests__/visual/helpers/global-teardown.ts',
  // Golden decks are written through Automizer.write, so Tier-1 package
  // invariants guard them like any other suite
  setupFilesAfterEnv: ['<rootDir>/__tests__/helpers/setup-pptx-invariants.ts'],
  // LibreOffice in the container converts one deck at a time
  maxWorkers: 1,
  testTimeout: 180000,
};
