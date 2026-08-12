/**
 * Tier-3 comparison core: render a written deck through the pinned
 * pptx-thumbnailer container and diff each slide against its committed
 * baseline PNG.
 *
 * The baseline means "what the pinned renderer showed when a human last
 * blessed it", not "what PowerPoint shows". A red deck answers "which feature
 * area changed visually" — whether the change is a regression or an intended
 * improvement is the reviewer's call, made by looking at the diff PNGs.
 *
 * Update flow: UPDATE_BASELINES=1 yarn test:visual rewrites
 * __tests__/visual-baselines/<deck>/ and the PR shows before/after images.
 */
import * as fs from 'fs';
import * as path from 'path';
import { PNG } from 'pngjs';
import pixelmatch from 'pixelmatch';
import { ThumbnailerClient } from 'pptx-thumbnailer';

const OUTPUT_DIR = path.resolve(__dirname, '../../pptx-output');
const BASELINE_ROOT = path.resolve(__dirname, '../../visual-baselines');
// Failure evidence (actual + diff PNGs); gitignored
const FAILURE_ROOT = path.resolve(__dirname, '../../visual-output');

// Antialiasing guarantees noise — never compare byte-equality. Per-pixel
// color threshold plus a small allowance of differing pixels.
const PIXEL_THRESHOLD = 0.1;
const MAX_DIFF_PIXEL_RATIO = 0.001;
const RENDER_WIDTH = 400;

const updateMode = process.env.UPDATE_BASELINES === '1';

const slideBaseline = (deck: string, index: number) =>
  path.join(BASELINE_ROOT, deck, `slide-${index}.png`);

/**
 * Renders `outputFile` (a deck previously written to __tests__/pptx-output)
 * and asserts every slide matches its baseline. `expectedSlides` pins the
 * deck length so a silently dropped or duplicated slide fails loudly instead
 * of shifting all following comparisons.
 */
export const expectDeckMatchesBaselines = async (
  deck: string,
  outputFile: string,
  expectedSlides: number,
): Promise<void> => {
  const thumbnailerUrl = process.env.THUMBNAILER_URL;
  if (!thumbnailerUrl) {
    throw new Error(
      'THUMBNAILER_URL is not set — run via `yarn test:visual` so the renderer container is started',
    );
  }

  const client = new ThumbnailerClient(thumbnailerUrl);
  const slides = await client.thumbnails(
    fs.readFileSync(path.join(OUTPUT_DIR, outputFile)),
    { width: RENDER_WIDTH },
  );

  expect(slides.length).toBe(expectedSlides);

  if (updateMode) {
    const deckDir = path.join(BASELINE_ROOT, deck);
    fs.rmSync(deckDir, { recursive: true, force: true });
    fs.mkdirSync(deckDir, { recursive: true });
    for (const slide of slides) {
      fs.writeFileSync(slideBaseline(deck, slide.index), slide.data);
    }
    return;
  }

  const failures: string[] = [];

  for (const slide of slides) {
    const baselinePath = slideBaseline(deck, slide.index);
    if (!fs.existsSync(baselinePath)) {
      failures.push(
        `slide ${slide.index}: no baseline at ${path.relative(process.cwd(), baselinePath)} — run UPDATE_BASELINES=1 yarn test:visual`,
      );
      continue;
    }

    const baseline = PNG.sync.read(fs.readFileSync(baselinePath));
    const actual = PNG.sync.read(slide.data);

    const failureDir = path.join(FAILURE_ROOT, deck);
    if (
      baseline.width !== actual.width ||
      baseline.height !== actual.height
    ) {
      fs.mkdirSync(failureDir, { recursive: true });
      const actualPath = path.join(failureDir, `slide-${slide.index}.actual.png`);
      fs.writeFileSync(actualPath, slide.data);
      failures.push(
        `slide ${slide.index}: dimensions changed ${baseline.width}x${baseline.height} -> ${actual.width}x${actual.height} (actual: ${path.relative(process.cwd(), actualPath)})`,
      );
      continue;
    }

    const diff = new PNG({ width: baseline.width, height: baseline.height });
    const diffPixels = pixelmatch(
      baseline.data,
      actual.data,
      diff.data,
      baseline.width,
      baseline.height,
      { threshold: PIXEL_THRESHOLD },
    );
    const allowed = Math.ceil(
      baseline.width * baseline.height * MAX_DIFF_PIXEL_RATIO,
    );

    if (diffPixels > allowed) {
      fs.mkdirSync(failureDir, { recursive: true });
      const actualPath = path.join(failureDir, `slide-${slide.index}.actual.png`);
      const diffPath = path.join(failureDir, `slide-${slide.index}.diff.png`);
      fs.writeFileSync(actualPath, slide.data);
      fs.writeFileSync(diffPath, PNG.sync.write(diff));
      failures.push(
        `slide ${slide.index}: ${diffPixels} pixels differ (allowed ${allowed}) — see ${path.relative(process.cwd(), diffPath)}`,
      );
    }
  }

  if (failures.length) {
    throw new Error(
      `golden deck '${deck}' diverged from its baselines:\n` +
        failures.map((failure) => `  - ${failure}`).join('\n') +
        '\n  If the change is intended: UPDATE_BASELINES=1 yarn test:visual and review the PNG diff in the PR.',
    );
  }
};
