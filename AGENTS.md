# AGENTS.md — Working on pptx-automizer

Guidance for AI agents (and new contributors) working on this codebase.
For the improvement plan derived from the 2026-08 architecture audit, see [ROADMAP.md](./ROADMAP.md).
For the end-user AI guide (how to *use* the library with an AI assistant), see [AI-INSTRUCTOR.md](./AI-INSTRUCTOR.md).

## What this library does

`pptx-automizer` is a **template-based** .pptx generator for Node.js. It does not
build presentations from scratch — it opens existing .pptx files (which are ZIP
archives of OOXML XML parts), copies slides/masters/shapes between them, and
mutates the XML through callback-based modifiers. New-from-scratch shapes are
delegated to the bundled PptxGenJS bridge.

Everything ultimately manipulates XML DOM nodes (via `@xmldom/xmldom`) inside a
ZIP archive (via `jszip`, or extracted to disk in `fs` mode).

## Commands

| Task | Command | Notes |
|---|---|---|
| Build | `yarn build` (tsc → `dist/`) | CJS only, no bundler |
| Test | `yarn test` (jest, ts-jest) | ~94 integration suites; writes real .pptx files to `__tests__/pptx-output/` |
| Single test | `npx jest __tests__/<name>.test.ts` | |
| Dev sandbox | `yarn dev` | Runs `src/dev.ts` with nodemon; scratchpad for manual testing |
| Lint | `yarn lint` | **Currently broken**: ESLint 9 is installed but config is legacy `.eslintrc.json` (flat config required). See ROADMAP. |
| Coverage | `yarn test-coverage` | **Broken glob**: `collectCoverageFrom` in `jest.config.ts` points at `.js` files |

There is **no CI** — run the test suite locally before considering a change done.

## Repository map

```
src/
  index.ts                 Public API surface: exports Automizer, `modify.*`, `read.*`, types
  automizer.ts             Facade/orchestrator class (load templates, addSlide, write/stream)
  dev.ts                   Manual dev playground (not part of the API, but currently compiled to dist/)
  classes/
    template.ts            Template = archive wrapper. Dual role: root (output) OR source template
    has-shapes.ts          ⚠ 1300-line base class of Slide/Master/Layout — element import pipeline
    slide.ts               Slide append logic, layout selection, placeholder merging
    master.ts, layout.ts   SlideMaster / SlideLayout import
    shape.ts               Base class for copyable shapes
  shapes/                  Chart, Image, Diagram, OLEObject, Hyperlink, GenericShape
  helper/
    xml-helper.ts          Generic XML/archive manipulation (append, removeIf, rel-id handling)
    modify-*-helper.ts     The public `modify.*` callback factories (text, table, chart, image, …)
    html-to-multitext-helper.ts / multitext-helper.ts
                           Rich-text pipeline: HTML → MultiTextParagraph[] → DrawingML
                           (known bugs + rework plan: see ROADMAP "HTML → PPTX text" track)
    xml-slide-helper.ts    Read-side slide introspection (getAllElements, dimensions, …)
    content-tracker.ts     ⚠ Global singleton tracking copied files/relations (used by cleanup)
    archive/               IArchive impls: archive-jszip.ts (default), archive-fs.ts (debugging)
    generate/              PptxGenJS bridge for `slide.generate(...)`
  types/                   Public + internal type defs (chart-types, table-types, modify-types, …)
  interfaces/              Interfaces (IArchive, ISlide, PresTemplate vs RootPresTemplate, …)
__tests__/                 Integration tests + template .pptx files (pptx-templates/)
```

## Core execution model (important!)

Nearly everything the user calls is **deferred**. `addSlide()`, `modifyElement()`,
`addElement()` only *queue* work. The actual XML manipulation happens inside
`automizer.write()` / `.stream()` / `.getJSZip()`, which run
`finalizePresentation()`:

1. `writeMasterSlides()` — append queued masters + their layouts
2. `writeSlides()` — for each queued slide: copy slide XML + `_rels`, copy related
   content (charts/images/…), run the PptxGenJS generator, find + import/modify/remove
   queued elements, apply modification callbacks, clean unsupported tags
3. `writeMediaFiles()`, `normalizePresentation()` (slide-ID normalization, optional cleanup),
   then user-level `modify()` callbacks on `ppt/presentation.xml`

Consequences for agents:

- A bug reported "when writing" usually originates from a callback queued much
  earlier. Trace the queue (`importElements`, `modifications`, `relModifications`).
- Errors thrown inside user shape callbacks are **swallowed** (`console.warn` in
  `shape.ts:applyCallbacks`). Silent failure is a known design weakness.
- Element type detection is tag-sniffing in `has-shapes.ts:analyzeElement()`
  (`c:chart`, `p:nvPicPr`, `dgm:relIds`, `p:oleObj`, hyperlink detection, else generic shape).

## Conventions & gotchas

- **OOXML paths are built inline everywhere** (`ppt/slides/slide${n}.xml`,
  `ppt/slides/_rels/slide${n}.xml.rels`). There is no central path builder yet.
  Slide/master/layout numbering is 1-based and file-name-driven.
- **Relationship IDs**: newly created rels get an `rId<max+1>-created` suffix
  (`xml-helper.ts:getNextRelId`). Don't "fix" this casually; cleanup logic depends on it.
- **Global state**: `contentTracker` (content-tracker.ts) and `Logger`
  (general-helper.ts) are module-level singletons. Two concurrent `Automizer`
  instances interfere with each other. Known issue, see ROADMAP.
- **Error style is inconsistent** (string throws vs `Error` objects vs
  `console.error` + `return undefined`). When touching code, prefer `throw new Error(...)`,
  but don't mass-convert in an unrelated PR.
- **XML caching**: `ArchiveJszip.readXml` parses once and caches the DOM per file
  path; `writeXml` only updates the cache. Serialization happens at output time.
  Reading a file "fresh" mid-run returns the mutated cached DOM — that's expected.
- **`creationId` vs name selectors**: shapes can be found by PowerPoint creationId
  (stable across renames) or by shape name (Selection pane). Both code paths must
  keep working; tests cover both (`modify-by-creation-id.test.ts`).
- **OOXML child-element order matters**: children of `a:pPr`/`a:rPr` (and most
  DrawingML property containers) must follow the schema sequence — e.g. in
  `a:pPr`: `lnSpc` → `spcBef` → `spcAft` → `buClr`/`buFont`/`buChar`/`buNone` →
  `defRPr`. Appending in call order instead of schema order is what triggers
  PowerPoint's "repair" prompt. `multitext-helper.ts` currently violates this
  (see ROADMAP). When creating property elements, insert schema-aware, don't
  just `appendChild`.
- **PPTX text is flat, never nested**: `txBody` → list of `a:p` → list of `a:r`.
  Anything hierarchical (HTML, nested lists) must be projected: inline nesting →
  accumulated run properties, list nesting → 0-based `lvl` attribute (0–8) per
  paragraph, block-in-block → innermost block becomes the paragraph.
- Uses **DOM lib types** (`lib: ["es2020", "dom"]` in tsconfig) mixed with xmldom.
  `XMLDocument` in some signatures is the browser type — misleading but compiles.
  `strict` mode is **off**.
- Language niceties: some template files/tests use German shape/layout names
  (e.g. `'Titel und Inhalt'`) — that's intentional, they mirror real templates.

## Testing rules

- Tests are integration-style: build a real presentation from `__tests__/pptx-templates/`,
  write it to `__tests__/pptx-output/`, and assert on the summary (slide/chart counts).
  They verify "does not crash / counts match", **not** XML correctness. When fixing an
  XML-level bug, add an assertion that reads the output archive and checks the XML
  (see ROADMAP "testing strategy") — don't just bump a count.
- The output .pptx files are real; when in doubt about visual/PowerPoint-level
  correctness, tell the maintainer which output file to open in PowerPoint.
- Never commit anything under `__tests__/pptx-output/`, `__customer__/`, or `dist/`
  (all gitignored). `src/dev-customer.ts` is gitignored too.
- New features need a test with a template .pptx. Prefer reusing existing templates
  in `__tests__/pptx-templates/` over adding new binaries.

## Public API stability

The published surface is what `src/index.ts` exports — most importantly the
`modify.*` namespace, `Automizer` itself, and the callback/data types
(`ChartData`, `TableData`, `TextStyle`, …). These are used by downstream projects
(automizer-data, Ensemblio). Treat them as semver-stable: additive changes are
fine, renames/behavioral changes need a deprecation path and a ROADMAP entry.

Internals (`XmlHelper`, `HasShapes`, shape classes) are exported or reachable but
undocumented — refactoring them is allowed per the ROADMAP, but keep `modify.*`
signatures intact.
