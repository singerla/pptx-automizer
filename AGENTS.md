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
    has-shapes.ts          Base class of Slide/Master/Layout: source/target context + deferred
                           queues; delegates the actual work to collaborators (Phase 2):
    element-importer.ts      element queue, getElementInfo/findElementOnSlide, typed dispatch
    related-content-copier.ts  copyRelatedContent (charts/images/diagrams/OLE/hyperlinks)
    slide-notes-copier.ts    notesSlide copy + number remapping
    placeholder-normalizer.ts  cleanSlide, duplicate/orphan placeholder cleanup
    content-type-registry.ts   presentation.xml slide lists + [Content_Types].xml entries
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
    content-tracker.ts     ContentTracker: per-instance tracking of copied files/relations (used by cleanup)
    media-deduplicator.ts  Checksum index of ppt/media on the root template: an
                           identical image is copied once and shared by all relations
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
- Errors thrown inside user shape callbacks **reject `write()`** with a typed
  `CallbackError` (Phase 1 policy: fail loudly). The lenient legacy behavior —
  log a warning, skip the modification — is opt-in via `continueOnError: true`.
- Element type detection is a tag-sniffing detector registry in
  `helper/shape-type-detector.ts` (`c:chart`, `p:nvPicPr`, `dgm:relIds`,
  `p:oleObj`, hyperlink detection, else generic shape). New shape types are
  one registry entry. Dispatch to the shape classes goes through the
  `IShapeAction` interface (`append`/`modify`/`remove`, called explicitly).

## Conventions & gotchas

- **OOXML part paths come from `helper/ppt-paths.ts`** (`PptPaths.slide(n)`,
  `PptPaths.slideRels(n)`, `PptPaths.chartPart('chart', n)`, …). Don't build
  `ppt/...` strings inline — add a helper to `PptPaths` if none fits.
  Slide/master/layout numbering is 1-based and file-name-driven.
- **Relationship IDs**: newly created rels get an `rId<max+1>-created` suffix
  (`xml-helper.ts:getNextRelId`). Don't "fix" this casually; cleanup logic depends on it.
- **No module-level state** (ROADMAP Phase 3): the `ContentTracker` instance is
  owned by `Automizer`/root template and reached via `archive.contentTracker`
  on the **output** archive (optional — absent on source archives, tracking
  calls no-op). The `log` facade resolves the active logger per async call
  tree (`AsyncLocalStorage`; `Automizer` entry points wrap themselves in
  `runWithLogger`). Don't reintroduce singletons — concurrent `Automizer`
  instances are supported and guarded by `__tests__/concurrent-instances.test.ts`.
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
  PowerPoint's "repair" prompt. Don't just `appendChild` — use
  `XmlHelper.insertInSchemaOrder(parent, child, order)` when you add a single
  element, or `XmlHelper.sortChildrenBySchema(parent, order)` when several
  independent code paths contribute to the same container (how `a:pPr` in
  `multitext-helper.ts` and `a:rPr` in `ModifyTextHelper.style` do it; the
  sequences live next to those call sites as `PPR_CHILD_ORDER` /
  `RPR_CHILD_ORDER`).
- **PPTX text is flat, never nested**: `txBody` → list of `a:p` → list of `a:r`.
  Anything hierarchical (HTML, nested lists) must be projected: inline nesting →
  accumulated run properties, list nesting → 0-based `lvl` attribute (0–8) per
  paragraph, block-in-block → innermost block becomes the paragraph.
- Uses **DOM lib types** (`lib: ["es2020", "dom"]` in tsconfig) mixed with xmldom.
  `XMLDocument` in some signatures is the browser type — misleading but compiles.
  `strict` mode is **off**.
- Language niceties: some template files/tests use German shape/layout names
  (e.g. `'Titel und Inhalt'`) — that's intentional, they mirror real templates.

## Adding a new `modify.*` helper

The most common contribution (and the most common AI-agent task) is "make
property X of shape type Y modifiable". It is additive and low-risk — but only
if all five steps are done. A helper that exists but isn't exported, tested and
documented is invisible.

1. **Implement it next to its peers.** Shape geometry/appearance →
   `helper/modify-shape-helper.ts`; text → `modify-text-helper.ts`; colors →
   `modify-color-helper.ts`; charts/tables/images → the respective helper. The
   signature is curried: `static setFoo = (params) => (element: XmlElement): void => {}`.
   Create new XML nodes through `helper/xml-elements.ts` (`XmlElements`) rather
   than ad-hoc `createElement` chains, and add a factory there if none fits.
2. **Reuse the existing vocabulary.** Check `types/modify-types.ts` /
   `shape-types.ts` first — e.g. `Color`, `Border { tag, type, weight, color }`,
   `TextStyle`, `ShapeCoordinates`. Two conventions for the same concept (say,
   line weight on table cells vs. on shapes) is a worse outcome than a slightly
   awkward reuse. New public types must be exported from `src/index.ts`.
3. **Wire it into the public surface**: `src/index.ts` — add the `const setFoo =
   XHelper.setFoo;` line *and* the entry in the `modify` namespace object. Both,
   or the helper is unreachable for users.
4. **Test it** in `__tests__/`, following the existing suites (e.g.
   `modify-shapes.test.ts`). Reuse a template from `__tests__/pptx-templates/`
   instead of adding a binary. Per the testing rules below, assert on the
   resulting **XML** read back from the output archive — not just
   `expect(result.slides).toBe(n)`.
5. **Document it** on the matching `docs/` feature page, in the same PR as the
   API change, then run `yarn docs:ai`. `AI-INSTRUCTOR.md` is **generated**
   from the docs corpus (ROADMAP 6.4) — never edit it directly; its
   hand-written parts live in `tools/ai-instructor/template.md`, and
   `__tests__/ai-instructor.test.ts` fails `yarn test` when the committed file
   drifts. Every fenced ```ts example in the corpus is typechecked by
   `__tests__/docs-examples.test.ts` on every `yarn test`, so a stale or
   invented signature fails the suite.

Then run `yarn test` and `npx tsc --noEmit` locally (CI runs the same, plus the
tier-2/3 gates) — and name the output .pptx that should be opened in PowerPoint
to verify.

Recurring OOXML pitfalls when implementing such a helper:

- **The element usually isn't there.** Unmodified properties are inherited from
  theme/master/shape style and absent from the slide XML. Handle "modify
  existing" and "create missing" both.
- **Scope lookups to the right container.** `element.getElementsByTagName('a:ln')`
  on a `<p:sp>` also finds line properties inside `a:rPr`. Get `p:spPr` first,
  then scan its direct children.
- **Respect schema child order** (see "Conventions & gotchas" above) — wrong
  order produces a file PowerPoint offers to repair; `yarn validate:pptx`
  (tier 2) catches it. Insert via `XmlHelper.insertInSchemaOrder`.
- **Units differ per attribute**: EMU for geometry and line width (1pt = 12700,
  1cm = 360000), 1/60000° for rotation, 1/100pt for font size, 1/1000% for
  percentages. Document the unit your parameter expects.

## Testing rules

> A four-tier testing model (XML assertions → package invariants → OOXML schema
> validation → visual regression via pptx-thumbnailer) is specified in
> ROADMAP Phase 5. All four tiers have landed; the rules below reflect them.

- **Tier 1 runs automatically**: `__tests__/helpers/setup-pptx-invariants.ts`
  (registered via jest `setupFilesAfterEnv`) validates **every archive any test
  writes** — referenced relationships resolve, parts are covered by
  `[Content_Types].xml`, the slide list is intact, all XML is well-formed. If it
  fails your test, you produced a broken deck: fix the change (or the test's
  deck setup, e.g. an internal hyperlink to a slide that doesn't exist). A test
  that *intentionally* writes a broken archive wraps its `write()` in
  `withoutPptxInvariants(...)` from the same module. Pre-existing tolerated
  behaviors (stale unreferenced rels, the never-copied notesMaster, orphaned
  parts after `removeExistingSlides` without cleanup) are classified as
  `knownIssues`, not errors — escalate one to an error only together with the
  library fix.
- **Tier 0 — assert on the XML, not just counts**: use
  `expectXml(outputFile, partPath)` from `__tests__/helpers/expect-xml.ts`
  (`.toContainElement`, `.toContainElementTimes`, `.toHaveAttribute`, plus
  `raw()`/`doc()` escape hatches; see `modify-existing-chart.test.ts` for the
  pattern). **Every bug fix adds the tier-0 assertion that would have caught
  it** — don't just bump a count. Note: colors are normalized to uppercase hex
  in output (`CCAA4F`), and row `label`s are not cell text.
- Tests are integration-style: build a real presentation from `__tests__/pptx-templates/`,
  write it to `__tests__/pptx-output/`, and assert on the summary (slide/chart counts)
  plus tier-0 XML assertions for whatever the test claims to verify.
- **Tier 2 — OOXML schema validation**: `yarn validate:pptx` (needs Docker
  only, no .NET) validates all templates and outputs with the Open XML SDK;
  CI runs it as the `validate-pptx` job. New schema errors fail the gate.
  `tools/validate-pptx/allowlist.json` holds baseline template noise plus
  documented library bugs (see ROADMAP "Bug track — schema violations found by
  the Tier-2 validator") — **never add an allowlist entry to silence an error
  your change introduced**; removing an entry belongs to the fix for it.
- Snapshot rule (when tier-0 snapshots land): snapshot the *modified subtree
  only*, canonicalized — never whole parts, never whole-file hashes.
- **Tier 3 — visual regression**: `yarn test:visual` (needs Docker only)
  renders curated golden decks (`__tests__/visual/*.deck.ts`) through the
  pinned pptx-thumbnailer container (`tools/render-pptx/`) and perceptually
  diffs each slide against `__tests__/visual-baselines/<deck>/`; CI runs it as
  the `visual-regression` job, uploading actual+diff PNGs on failure. It is a
  **change detector, not a correctness oracle** — LibreOffice fidelity is not
  PowerPoint fidelity; never conclude PowerPoint-correctness from green pixels.
  On an *intended* visual change, regenerate with
  `UPDATE_BASELINES=1 yarn test:visual` in the same PR so the reviewer sees the
  before/after PNGs. **Never render all suites** — ~12 curated golden decks
  only (see the deck table in ROADMAP Phase 5); new decks must be small
  (1–5 slides), stick to fonts shipped in the renderer image
  (Liberation/Carlito families), and reuse existing templates. Changing
  anything in `tools/render-pptx/Dockerfile` (base digest, thumbnailer
  version, fonts) invalidates all baselines: regenerate them in the same PR.
- The output .pptx files are real; when in doubt about visual/PowerPoint-level
  correctness, tell the maintainer which output file to open in PowerPoint.
- Never commit anything under `__tests__/pptx-output/`, `__customer__/`, or `dist/`
  (all gitignored). `src/dev-customer.ts` is gitignored too.
- New features need a test with a template .pptx. Prefer reusing existing templates
  in `__tests__/pptx-templates/` over adding new binaries.
- **Documented examples compile** (ROADMAP Phase 6.1): every fenced ```ts block
  in `README.md`, `AI-INSTRUCTOR.md` and `docs/**/*.md` is typechecked against
  src by `__tests__/docs-examples.test.ts` in every `yarn test`. Change a
  documented API → update its examples in the same commit, or the suite fails.
  A deliberately partial snippet opts out with a ```ts ignore fence. Snippets
  may use the conventional context variables (`pres`, `slide`, `modify`, the
  `Modify*Helper` classes, …) declared in
  `__tests__/helpers/docs-example-context.d.ts` — add a name there only for a
  docs-wide convention, never to silence one failing example.
- **`AI-INSTRUCTOR.md` is generated** (ROADMAP 6.4): edit the `docs/` pages or
  `tools/ai-instructor/template.md` and run `yarn docs:ai` in the same commit
  — `__tests__/ai-instructor.test.ts` pins the committed file to the build.

## Public API stability

The published surface is what `src/index.ts` exports — most importantly the
`modify.*` namespace, `Automizer` itself, and the callback/data types
(`ChartData`, `TableData`, `TextStyle`, …). These are used by downstream projects
(e.g. automizer-data). Treat them as semver-stable: additive changes are
fine, renames/behavioral changes need a deprecation path and a ROADMAP entry.

Internals (`XmlHelper`, `HasShapes`, shape classes) are exported or reachable but
undocumented — refactoring them is allowed per the ROADMAP, but keep `modify.*`
signatures intact.
