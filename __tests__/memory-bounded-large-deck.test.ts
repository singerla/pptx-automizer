import { execFileSync } from 'child_process';
import * as path from 'path';

/**
 * Regression gate for the ROADMAP performance track: appending large slides
 * must not keep every appended slide's parsed DOM in the archive buffer
 * (an xmldom document costs ~25x its XML source size — before the fix a
 * 60-slide synthetic deck held ~124 buffered DOMs and ~566 MB of heap).
 *
 * The deck is built in a child process with --expose-gc so the heap number
 * is measured after a full GC and is independent of jest's own allocations.
 */
test('memory stays bounded while appending large slides', () => {
  const slideCount = 60;

  const stdout = execFileSync(
    process.execPath,
    [
      '-r',
      require.resolve('ts-node/register/transpile-only'),
      '--expose-gc',
      path.join(__dirname, 'helpers', 'large-deck-memory-child.ts'),
      String(slideCount),
    ],
    {
      cwd: path.join(__dirname, '..'),
      encoding: 'utf-8',
      timeout: 240_000,
    },
  );

  const lastLine = stdout.trim().split('\n').pop();
  const result = JSON.parse(lastLine) as {
    slides: number;
    status: string;
    bufferedParts: number;
    heapUsedMb: number;
  };

  expect(result.status).toBe('finished');
  expect(result.slides).toBe(slideCount);

  // Eviction gate: only shared parts (presentation.xml + rels,
  // [Content_Types].xml, ...) may stay buffered — never O(slides).
  // Pre-fix value: 2 * slideCount + 4.
  expect(result.bufferedParts).toBeLessThanOrEqual(10);

  // Heap gate with wide margins: ~54 MB after the fix, ~566 MB before.
  expect(result.heapUsedMb).toBeLessThan(200);
}, 240_000);
