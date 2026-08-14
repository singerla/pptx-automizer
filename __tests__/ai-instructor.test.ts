/**
 * ROADMAP 6.4: AI-INSTRUCTOR.md is generated from tools/ai-instructor/
 * template.md + the docs corpus. This test pins the committed file to the
 * build output, so editing a docs page (or the template) without regenerating
 * fails `yarn test`.
 */
import * as fs from 'fs';
import * as path from 'path';
import { buildAiInstructor } from '../tools/ai-instructor/build';

test('AI-INSTRUCTOR.md matches its build (regenerate: yarn docs:ai)', () => {
  const committed = fs.readFileSync(
    path.resolve(__dirname, '../AI-INSTRUCTOR.md'),
    'utf-8',
  );
  expect(committed).toBe(buildAiInstructor());
});
