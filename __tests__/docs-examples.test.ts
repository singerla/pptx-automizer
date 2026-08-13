/**
 * Phase 6.1 (ROADMAP): every fenced ```ts block in the documentation corpus
 * must typecheck against the current source. Blocks that are deliberately
 * partial opt out with ```ts ignore (GitHub still highlights them as ts).
 *
 * Each block is wrapped in a synthetic module (`export {};` prefix) and the
 * whole batch is checked by one TS program, so the src type graph is only
 * built once. `import ... from 'pptx-automizer'` resolves to src/index.ts
 * via a paths mapping.
 */
import * as fs from 'fs';
import * as path from 'path';
import * as ts from 'typescript';

const repoRoot = path.resolve(__dirname, '..');

type DocBlock = {
  mdFile: string;
  // 1-based line of the opening ``` fence in the md file
  fenceLine: number;
  code: string;
  ignored: boolean;
  virtualName: string;
};

const collectMdFiles = (): string[] => {
  const files = ['README.md', 'AI-INSTRUCTOR.md'];
  const docsDir = path.join(repoRoot, 'docs');
  const walk = (dir: string) => {
    for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
      const full = path.join(dir, entry.name);
      if (entry.isDirectory()) walk(full);
      else if (entry.name.endsWith('.md'))
        files.push(path.relative(repoRoot, full));
    }
  };
  if (fs.existsSync(docsDir)) walk(docsDir);
  return files.filter((f) => fs.existsSync(path.join(repoRoot, f)));
};

const extractBlocks = (mdFile: string): DocBlock[] => {
  const lines = fs
    .readFileSync(path.join(repoRoot, mdFile), 'utf-8')
    .split('\n');
  const blocks: DocBlock[] = [];
  let open: { fenceLine: number; ignored: boolean; code: string[] } | null =
    null;
  lines.forEach((line, i) => {
    if (!open) {
      const fence = line.match(/^```ts\b(.*)$/);
      if (fence) {
        open = {
          fenceLine: i + 1,
          ignored: /\bignore\b/.test(fence[1]),
          code: [],
        };
      }
    } else if (/^```\s*$/.test(line)) {
      blocks.push({
        mdFile,
        fenceLine: open.fenceLine,
        code: open.code.join('\n'),
        ignored: open.ignored,
        virtualName: path.join(
          repoRoot,
          `__docs-examples__/${mdFile.replace(/[\\/]/g, '__')}.L${
            open.fenceLine
          }.ts`,
        ),
      });
      open = null;
    } else {
      open.code.push(line);
    }
  });
  return blocks;
};

const allBlocks = collectMdFiles().flatMap(extractBlocks);
const checkedBlocks = allBlocks.filter((b) => !b.ignored);

// `export {};` turns every snippet into a module: block-local scope and
// top-level await both depend on it. One prefix line to subtract when
// mapping diagnostics back to md line numbers.
const PRELUDE = 'export {};\n';

const compilerOptions: ts.CompilerOptions = {
  noEmit: true,
  strict: false,
  noImplicitAny: true,
  strictBindCallApply: true,
  esModuleInterop: true,
  skipLibCheck: true,
  target: ts.ScriptTarget.ES2020,
  lib: ['lib.es2020.d.ts'],
  // top-level await is legal in the docs' async snippets
  module: ts.ModuleKind.ES2022,
  moduleResolution: ts.ModuleResolutionKind.Node10,
  types: ['node'],
  baseUrl: repoRoot,
  paths: {
    'pptx-automizer': ['src/index.ts'],
    // deep imports must use the published layout (only dist/ ships to npm)
    'pptx-automizer/dist/*': ['src/*'],
  },
};

// conventional docs context (`pres`, `slide`, `modify`, …) — see the file
const contextFile = path.join(__dirname, 'helpers/docs-example-context.d.ts');

const typecheckBatch = (): Map<string, string[]> => {
  const virtual = new Map<string, string>(
    checkedBlocks.map((b) => [b.virtualName, PRELUDE + b.code]),
  );
  const host = ts.createCompilerHost(compilerOptions);
  const defaultGetSourceFile = host.getSourceFile.bind(host);
  const defaultFileExists = host.fileExists.bind(host);
  const defaultReadFile = host.readFile.bind(host);
  host.getSourceFile = (fileName, languageVersion, ...rest) => {
    const content = virtual.get(path.resolve(fileName));
    if (content !== undefined) {
      return ts.createSourceFile(fileName, content, languageVersion, true);
    }
    return defaultGetSourceFile(fileName, languageVersion, ...rest);
  };
  host.fileExists = (fileName) =>
    virtual.has(path.resolve(fileName)) || defaultFileExists(fileName);
  host.readFile = (fileName) =>
    virtual.get(path.resolve(fileName)) ?? defaultReadFile(fileName);

  const program = ts.createProgram(
    [contextFile, ...checkedBlocks.map((b) => b.virtualName)],
    compilerOptions,
    host,
  );

  const failures = new Map<string, string[]>();
  for (const block of checkedBlocks) {
    const sourceFile = program.getSourceFile(block.virtualName);
    if (!sourceFile) {
      failures.set(block.virtualName, ['block source file missing']);
      continue;
    }
    const diagnostics = [
      ...program.getSyntacticDiagnostics(sourceFile),
      ...program.getSemanticDiagnostics(sourceFile),
    ];
    if (diagnostics.length) {
      failures.set(
        block.virtualName,
        diagnostics.map((d) => {
          const message = ts.flattenDiagnosticMessageText(
            d.messageText,
            '\n  ',
          );
          if (d.file && d.start !== undefined) {
            const pos = d.file.getLineAndCharacterOfPosition(d.start);
            // line 0 is the prelude; md line = fence + snippet line
            const mdLine = block.fenceLine + pos.line;
            return `${block.mdFile}:${mdLine} — TS${d.code}: ${message}`;
          }
          return `TS${d.code}: ${message}`;
        }),
      );
    }
  }
  return failures;
};

describe('documented ts examples compile (README, AI-INSTRUCTOR, docs/)', () => {
  it('finds the corpus', () => {
    // If extraction silently broke, every other test would pass vacuously.
    expect(allBlocks.length).toBeGreaterThan(50);
  });

  const failures = typecheckBatch();

  checkedBlocks.forEach((block) => {
    it(`${block.mdFile}:${block.fenceLine}`, () => {
      const errors = failures.get(block.virtualName);
      if (errors) {
        throw new Error(errors.join('\n'));
      }
    });
  });

  allBlocks
    .filter((b) => b.ignored)
    .forEach((block) => {
      it.skip(`${block.mdFile}:${block.fenceLine} (ts ignore)`, (): void =>
        undefined);
    });
});
