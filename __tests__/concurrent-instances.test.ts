import Automizer, { AutomizerSummary } from '../src/index';
import { ILogger, log, runWithLogger } from '../src/helper/logger';
import JSZip from 'jszip';
import * as fs from 'fs';

const outputDir = `${__dirname}/pptx-output`;

const makeAutomizer = () =>
  new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir,
    // cleanup runs removedUnusedImages, which reads the content tracker;
    // with a shared tracker, two concurrent writes corrupt each other here.
    cleanup: true,
  });

const buildCharts = (filename: string): Promise<AutomizerSummary> =>
  makeAutomizer()
    .loadRoot(`RootTemplate.pptx`)
    .load(`SlideWithCharts.pptx`, 'charts')
    .addSlide('charts', 1)
    .addSlide('charts', 2)
    .write(filename);

const buildImages = (filename: string): Promise<AutomizerSummary> =>
  makeAutomizer()
    .loadRoot(`RootTemplate.pptx`)
    .load(`SlideWithImages.pptx`, 'images')
    .addSlide('images', 1)
    .addSlide('images', 2)
    .write(filename);

const loadOutput = async (filename: string): Promise<JSZip> =>
  JSZip.loadAsync(fs.readFileSync(`${outputDir}/${filename}`));

/**
 * Assert that a concurrently written archive contains exactly the same
 * parts with exactly the same content as its sequentially written twin.
 * Any shared state between instances (content tracker, logger) surfaces
 * here as diverging parts - extra/missing media, foreign relations.
 */
const expectSameArchive = async (concurrent: JSZip, sequential: JSZip) => {
  const partNames = (zip: JSZip) =>
    Object.keys(zip.files)
      .filter((name) => !zip.files[name].dir)
      .sort();

  expect(partNames(concurrent)).toEqual(partNames(sequential));

  for (const name of partNames(sequential)) {
    const contentConcurrent = await concurrent.file(name).async('uint8array');
    const contentSequential = await sequential.file(name).async('uint8array');
    expect(
      `${name}: ${
        Buffer.compare(
          Buffer.from(contentConcurrent),
          Buffer.from(contentSequential),
        ) === 0
          ? 'equal'
          : 'differs'
      }`,
    ).toBe(`${name}: equal`);
  }
};

test('two automizer instances write valid output in parallel', async () => {
  // Baseline: the same two decks built one after another.
  await buildCharts(`concurrent-instances-charts-sequential.test.pptx`);
  await buildImages(`concurrent-instances-images-sequential.test.pptx`);

  // Regression under test: both instances written in parallel.
  const [resultCharts, resultImages] = await Promise.all([
    buildCharts(`concurrent-instances-charts.test.pptx`),
    buildImages(`concurrent-instances-images.test.pptx`),
  ]);

  expect(resultCharts.slides).toBe(3);
  expect(resultImages.slides).toBe(3);

  const zipCharts = await loadOutput('concurrent-instances-charts.test.pptx');
  const zipImages = await loadOutput('concurrent-instances-images.test.pptx');

  await expectSameArchive(
    zipCharts,
    await loadOutput('concurrent-instances-charts-sequential.test.pptx'),
  );
  await expectSameArchive(
    zipImages,
    await loadOutput('concurrent-instances-images-sequential.test.pptx'),
  );

  // No cross-contamination between the two outputs.
  expect(zipCharts.file(/ppt\/charts\/chart\d+\.xml/).length).toBeGreaterThan(0);
  expect(zipImages.file(/ppt\/charts\/chart\d+\.xml/).length).toBe(0);
  expect(zipImages.file(/ppt\/media\/.+/).length).toBeGreaterThan(0);
});

test('concurrently running instances log through their own logger', async () => {
  const capture = (messages: string[]): ILogger => ({
    error: (message) => messages.push(String(message)),
    warn: (message) => messages.push(String(message)),
    info: (message) => messages.push(String(message)),
    debug: (message) => messages.push(String(message)),
  });
  const messagesA: string[] = [];
  const messagesB: string[] = [];

  await Promise.all([
    runWithLogger(capture(messagesA), async () => {
      await new Promise((resolve) => setTimeout(resolve, 20));
      log.info('from A');
    }),
    runWithLogger(capture(messagesB), async () => {
      log.warn('from B');
      await new Promise((resolve) => setTimeout(resolve, 40));
      log.info('from B again');
    }),
  ]);

  expect(messagesA).toEqual(['from A']);
  expect(messagesB).toEqual(['from B', 'from B again']);
});
