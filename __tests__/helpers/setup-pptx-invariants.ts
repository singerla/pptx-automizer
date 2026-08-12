import fs from 'fs';
import Automizer from '../../src/automizer';
import { AutomizerSummary } from '../../src/types/types';
import { checkPptxInvariants } from './pptx-invariants';

/**
 * ROADMAP Phase 5, Tier 1 — registered via jest `setupFilesAfterEnv`.
 *
 * Wraps `Automizer.write()` so every archive written by any test suite is
 * checked against the package invariants without the suite opting in. A
 * violation fails the test that wrote the file, with the part and
 * relationship named in the message.
 *
 * Escape hatch for tests that intentionally write a broken archive:
 * `withoutPptxInvariants(async () => { ... })`.
 */

let invariantsEnabled = true;

export async function withoutPptxInvariants<T>(
  scope: () => Promise<T>,
): Promise<T> {
  invariantsEnabled = false;
  try {
    return await scope();
  } finally {
    invariantsEnabled = true;
  }
}

type HasGetLocation = { getLocation(location: string, type?: string): string };

const originalWrite = Automizer.prototype.write;

Automizer.prototype.write = async function (
  location: string,
): Promise<AutomizerSummary> {
  const summary = await originalWrite.call(this, location);
  if (invariantsEnabled) {
    const outputPath = (this as unknown as HasGetLocation).getLocation(
      location,
      'output',
    );
    const { errors } = await checkPptxInvariants(fs.readFileSync(outputPath));
    if (errors.length) {
      throw new Error(
        `pptx package invariants violated in ${outputPath}:\n` +
          errors.map((error) => `  - ${error}`).join('\n'),
      );
    }
  }
  return summary;
};
