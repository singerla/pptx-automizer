/**
 * ROADMAP 6.4 — AI-INSTRUCTOR.md is generated, not written:
 * `template.md` holds the hand-written parts (mental model, rules, minimal
 * example) and pulls everything else from the docs corpus via directives
 *
 *   <!-- include: docs/<page>.md -->            whole page
 *   <!-- include: docs/<page>.md § <Heading> -->  one section
 *
 * Regenerate with `yarn docs:ai`; `__tests__/ai-instructor.test.ts` fails
 * whenever the committed AI-INSTRUCTOR.md drifts from this build.
 */
import * as fs from 'fs';
import * as path from 'path';

const repoRoot = path.resolve(__dirname, '../..');

// The docs are served from GitHub Pages only (ROADMAP 6.5); every page has a
// Markdown twin at its URL + `.md`, so relative docs links can be absolutized.
const SITE_ROOT = 'https://singerla.github.io/pptx-automizer/';

const readDoc = (relPath: string): { title: string; body: string } => {
  const raw = fs.readFileSync(path.join(repoRoot, relPath), 'utf-8');
  const fm = raw.match(/^---\r?\n([\s\S]*?)\r?\n---\r?\n?/);
  if (!fm) throw new Error(`ai-instructor: ${relPath} has no frontmatter`);
  const title = fm[1].match(/^title:\s*(.+)$/m)?.[1]?.trim();
  if (!title) throw new Error(`ai-instructor: ${relPath} has no title`);
  return { title, body: raw.slice(fm[0].length).trim() };
};

// Relative docs links point at the published Markdown twins in the generated
// file ('./api/index.md' would be served as 'api.md', see the llms-txt plugin).
const rewriteLinks = (md: string): string =>
  md
    .replace(/\]\((\.\/[^)#]+)\/index\.md(#[^)]*)?\)/g, ']($1.md$2)')
    .replace(/\]\(\.\//g, `](${SITE_ROOT}`);

// Shift heading levels outside fenced code blocks (a '#' at line start inside
// a fence is e.g. a bash comment, not a heading).
const shiftHeadings = (md: string, shift: number): string => {
  let inFence = false;
  return md
    .split('\n')
    .map((line) => {
      if (/^\s*(```|~~~)/.test(line)) {
        inFence = !inFence;
        return line;
      }
      if (inFence) return line;
      const heading = line.match(/^(#{1,6})(\s.*)$/);
      if (!heading) return line;
      const level = Math.min(6, Math.max(1, heading[1].length + shift));
      return '#'.repeat(level) + heading[2];
    })
    .join('\n');
};

const renderPage = (relPath: string): string => {
  const { title, body } = readDoc(relPath);
  return `## ${title}\n\n${rewriteLinks(shiftHeadings(body, 1))}`;
};

const renderSection = (relPath: string, heading: string): string => {
  const { body } = readDoc(relPath);
  const lines = body.split('\n');
  let inFence = false;
  let start = -1;
  let level = 0;
  let end = lines.length;
  for (let i = 0; i < lines.length; i++) {
    if (/^\s*(```|~~~)/.test(lines[i])) inFence = !inFence;
    if (inFence) continue;
    const match = lines[i].match(/^(#{1,6})\s+(.*?)\s*$/);
    if (!match) continue;
    if (start === -1) {
      if (match[2] === heading) {
        start = i;
        level = match[1].length;
      }
    } else if (match[1].length <= level) {
      end = i;
      break;
    }
  }
  if (start === -1) {
    throw new Error(`ai-instructor: section "${heading}" not in ${relPath}`);
  }
  const section = lines.slice(start, end).join('\n').trim();
  return rewriteLinks(shiftHeadings(section, 2 - level));
};

export const buildAiInstructor = (): string => {
  const template = fs.readFileSync(
    path.join(__dirname, 'template.md'),
    'utf-8',
  );
  const out = template.replace(
    /^<!-- include: (\S+)(?: § (.+?))? -->$/gm,
    (_, relPath: string, heading?: string) =>
      heading ? renderSection(relPath, heading) : renderPage(relPath),
  );
  return out.trimEnd() + '\n';
};

if (require.main === module) {
  const target = path.join(repoRoot, 'AI-INSTRUCTOR.md');
  fs.writeFileSync(target, buildAiInstructor());
  console.log(`generated ${target}`);
}
