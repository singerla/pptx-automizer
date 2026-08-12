# ROADMAP — Architecture audit & improvement plan

Audit date: 2026-08-11 (v0.8.2). Goal: consolidate architecture **before**
adding further functionality. Findings are ordered into phases; each phase is
shippable on its own and phases 0–2 unblock the rest.

Legend: 🐛 bug · 🔧 tooling · 🏗 architecture · 🧪 testing · 📖 docs

---

## Phase 0 — Correctness & tooling quick wins (hours, not days) — ✅ done 2026-08-11

These are small, independent fixes. Do them first; they cost nothing and remove
noise from every later change.

1. 🐛 **`package.json` `types` field is wrong**: `"types": "dist/index.d.js"`
   → must be `dist/index.d.ts`. It currently works only because TS falls back to
   looking for a `.d.ts` next to `main`.
2. 🐛 **Wrong duration divisor** in `automizer.ts:write()`:
   `(Date.now() - this.timer) / 600` — should be `/ 1000` (seconds). The summary
   `duration` reported to users is wrong by ~40%.
3. 🔧 **`yarn lint` is broken**: ESLint 9 is installed but the config is legacy
   `.eslintrc.json`. Migrate to flat `eslint.config.mjs` (typescript-eslint v8
   provides `tseslint.config()` helper). Until then, nobody is linting.
4. 🔧 **Coverage config broken**: `jest.config.ts` has
   `collectCoverageFrom: ['src/helper/{!(pretty),}.js']` — a `.js` glob in a TS
   repo. Replace with `['src/**/*.ts', '!src/dev.ts']`.
5. 🔧 **`src/dev.ts` ships to npm**: it's compiled into `dist/dev.js` and
   published. Exclude it (`tsconfig` exclude or move to `dev/` outside `src`).
   Same for the gitignored-but-compiled `dev-customer.ts` pattern.
6. 🔧 **`files: ["dist", "README"]`** — the `README` entry refers to a
   non-existent file (`README.md` is auto-included anyway). Add `AI-INSTRUCTOR.md`
   here once it exists so users get it from `node_modules`.
7. 🔧 **No CI.** Add a GitHub Actions workflow: `yarn install`, `tsc --noEmit`,
   `yarn test` on Node 18/20/22. The repo takes external PRs (e.g. #193) with no
   automated safety net — this is the single highest-leverage tooling fix.
8. 🐛 **`Automizer.getLocation()` silently returns `undefined`** for a missing
   template (it logs and falls through) — the user then gets
   `File not found: undefined` from deep inside `FileHelper`. Throw a
   `TemplateNotFoundError` with the searched dirs right there.

## Phase 1 — Error handling & logging (foundation for everything else) — ✅ done 2026-08-11

Current state: 14 string `throw`s (`throw 'no chart found'` style), `throw new Error`
elsewhere, `console.error` + `return undefined` in the element pipeline
(`has-shapes.ts:getElementInfo`), swallowed callback exceptions
(`shape.ts:applyCallbacks` → `console.warn`), and ~40 raw `console.*` calls next
to a half-finished `Logger` singleton.

- 🏗 Introduce a small error hierarchy in `src/errors.ts`:
  `AutomizerError` → `TemplateNotFoundError`, `SlideNotFoundError`,
  `ElementNotFoundError`, `ArchiveError`, `OutputError`. Replace all string throws
  (string throws have no stack trace and break `instanceof` handling downstream).
- 🏗 Decide and document a policy for user-callback errors: default should be
  **fail loudly** (rethrow with slide/element context), with an opt-in
  `params.continueOnError` for the current lenient behavior. Today a typo'd
  selector or a crashing callback produces a silently wrong deck — the worst
  failure mode for a reporting tool.
- 🏗 Make `Logger` an instance owned by `Automizer` (injectable, default console),
  route all `console.*` through it. Keeps library output silent by default when
  embedded in servers.

## Phase 2 — Decompose `HasShapes` and centralize OOXML paths — ✅ done 2026-08-11

`src/classes/has-shapes.ts` is a 1300-line base class mixing at least six
concerns. It's where most future features will land, so pay this debt first.

- 🏗 ✅ Extract collaborators (composition over inheritance) — `HasShapes` keeps
  its (undocumented) public surface and delegates; collaborators live in
  `src/classes/` and read their context from the owning instance:
  - `ElementImporter` — queueing + `getElementInfo`/`findElementOnSlide`/dispatch
    (the `importedSelectedElements` switch)
  - `RelatedContentCopier` — `copyRelatedContent` (charts/images/diagrams/OLE/hyperlinks)
  - `SlideNotesCopier` — the notesSlide number remapping trio
  - `PlaceholderNormalizer` — `removeDuplicatePlaceholders`/`normalizePlaceholderShapes`/`cleanSlide`
  - `ContentTypeRegistry` — the `appendToContentType`/`appendToSlideList` family
    (now also owns the notes + theme content-type entries)
- 🏗 ✅ **Central `PptPaths` helper** (`src/helper/ppt-paths.ts`): `slide(n)`,
  `slideRels(n)`, `slideMaster(n)`, `slideLayout(n)`, `notesSlide(n)`,
  `chartPart(name, n)`, `media(file)`, `embedding(file)`, generic
  `part(type, n)`/`partRels(type, n)`, relative rel-target variants and
  `partName()` for `[Content_Types].xml`. Replaced ~40 inline template strings
  across classes, shapes and helpers.
- 🏗 ✅ Shape-type detector registry (`src/helper/shape-type-detector.ts`)
  replaces the `analyzeElement` if-chain. New shape types (video/audio are on
  the wish list) become one registry entry; hyperlinks keep their custom
  `analyze` handler.
- 🏗 ✅ `IShapeAction { append; modify; remove }`
  (`src/interfaces/ishape-action.ts`), implemented by all shape classes;
  dispatch calls methods explicitly, `ImportElement.mode` is typed
  `ShapeActionMode`. Behavior note: `OLEObject.append/modify` now throw a
  descriptive `AutomizerError` instead of crashing with an opaque `TypeError`
  (OLE still supports `remove` only).

## Phase 3 — Kill global state (enables concurrent instances) — ✅ done 2026-08-11

- 🏗 ✅ `contentTracker` module-level singleton removed. The instance is owned
  by `Automizer` (`automizer.content`), shared with the root `Template`
  (`rootTemplate.content`) and attached to the **output archive**
  (`IArchive.contentTracker`, optional) — which is exactly the object every
  former direct importer (`xml-helper.ts`, `file-helper.ts`, `chart.ts`,
  `modify-presentation-helper.ts`) already had in scope. Source-template
  archives carry no tracker; tracking calls no-op there. Also fixed
  `ContentTracker.getRelationTag`/`pushRelationTagTargets` referencing the
  singleton instead of `this`.
- 🏗 ✅ Module-level `setActiveLogger` removed. The `log` facade resolves the
  active logger per async call tree via `AsyncLocalStorage`: `Automizer` wraps
  its entry points (`write`, `stream`, `getJSZip`, `finalizePresentation`,
  `setCreationIds`) in `runWithLogger(this.logger, …)`, so concurrent
  instances log through their own logger. Explicit threading was ruled out
  because public `modify.*` signatures cannot carry a logger (see the
  no-signature-changes rule). Outside any instance context, `log` falls back
  to a default `ConsoleLogger`.
- 🧪 ✅ Regression test `__tests__/concurrent-instances.test.ts`: two
  instances (charts deck / images deck, `cleanup: true`) written via
  `Promise.all` must produce archives part-for-part identical to their
  sequentially built twins, plus a logger-isolation test. On the pre-Phase-3
  code this test crashes with `Could not find file ppt/charts/chart1.xml@RootTemplate.pptx`
  — instance B tripping over instance A's tracked relations.

## Phase 4 — Template & archive layer clarity — ✅ done 2026-08-11

- 🏗 ✅ `Template` split into `SourceTemplate` and `OutputTemplate` extending a
  shared abstract `Template` base. `Template.import` stays as the factory and
  returns the concrete union; all `as PresTemplate`/`as RootPresTemplate` casts
  and the `'name' in template` check (`isPresTemplate`) are gone.
  `ITemplate.file` is typed `AutomizerFile` (was the wrong `ArchiveInput`).
- 🏗 ✅ Archive initialization is a private idempotent `ensureOpen()` awaited by
  every method that needs the loaded archive (also covers `write`/`folder`/
  `output` being called first — a latent gap in `ArchiveFs`). All
  `await x.archive` pseudo-awaits removed; `extract()`-created instances with a
  preloaded inner archive short-circuit the guard.
- 🏗 ✅ `zipCopyWithRelations`/`zipCopyByIndex` take a typed
  `ArchiveCopyContext` instead of an untyped "parentClass"; `Archive` buffer
  methods and the remaining loose helpers annotated.
- 🔧 ✅ `noImplicitAny` + `strictBindCallApply` are on (tsconfig). Genuine
  catches: `groupElements` indexed `GroupedByType` with `ElementInfo['type']`
  instead of `['visualType']`; a dead `'breaks'` key in
  `groupSimilarParagraphs`; `SlidePlaceholder.id` declared `number` but always
  a string. `read.readWorkbookData`/`readChartInfo` gained typed accumulators
  (exported `WorkbookData`/`ChartInfo`); `setBulletList` content is the
  exported `BulletListContent`.
  ⏳ `strictNullChecks` deferred: ~424 errors as of 2026-08-11 — its own
  phase-sized chore, best tackled file-by-file with `// @ts-expect-error`-free
  boundaries.
- 🔧 ✅ `dom` removed from `lib`; `XmlDocument`/`XmlElement` alias
  `@xmldom/xmldom`'s `Document`/`Element`, plus new `XmlElementCollection`/
  `XmlNodeCollection` aliases replacing `HTMLCollectionOf`/`NodeListOf`.
  Published `.d.ts` reference `@xmldom/xmldom` (a runtime dependency).
  ⏳ Dual CJS+ESM output deferred — packaging change, do it as a release of
  its own.
- 🏗 ✅ The pptxgenjs generator bridge is dynamically imported and only when at
  least one slide has `generate()` elements; a modify-only run never loads
  pptxgenjs (verified against the compiled output). ⏳ Making pptxgenjs an
  **optional peer dependency** deferred — breaking packaging change; the lazy
  import already removes the runtime cost.

## Phase 5 — Testing strategy (four tiers)

The 94-suite integration harness is a real asset (it exercises actual OOXML end
to end). Its weakness: assertions are almost all `expect(result.slides).toBe(n)` —
the content of the produced XML is unverified, so regressions that corrupt output
without crashing pass green.

A test on generated pptx can answer one of three different questions, and no
single fixture type answers all of them:

1. **Did my edit arrive in the XML?** (precision) → only XML assertions.
2. **Is the file valid — will PowerPoint open it without the repair prompt?**
   (validity) → only a schema validator. XML assertions can't catch this (the
   wrongly-ordered XML is exactly what you asserted) and LibreOffice can't
   either (it renders invalid files happily).
3. **Does it look right?** (rendering) → only pixels.

Hence four tiers. Tiers 0–1 run in every `yarn test`; tiers 2–3 are separate CI
jobs. Note the complementarity: an intentionally *invisible* XML change keeps
tier 3 green while tier 0 asserts it; an intended *visible* change fails tier 3
once, and the reviewer approves a PNG diff in the PR instead of opening the file.

### Tier 0 — per-assertion XML checks (fast, always on)

- 🧪 ✅ **Output-assertion helper** (2026-08-12): `__tests__/helpers/expect-xml.ts` —
  `expectXml(output, 'ppt/slides/slide1.xml').toContainElement('a:t', 'my text')`,
  plus `toContainElementTimes`/`toHaveAttribute` and `raw()`/`doc()` escape
  hatches. Retro-fitted onto `modify-existing-chart`, `modify-existing-table`
  and `replace-tagged-text` (the latter two previously asserted nothing).
- 🧪 **Targeted subtree snapshots**: jest snapshots of the *modified subtree
  only* (the shape's `<p:sp>`, a chart's `<c:ser>`), canonicalized
  (pretty-printed, attributes sorted) so diffs are readable. Never snapshot
  whole parts — serialization refactors would churn every snapshot and train
  everyone to `jest -u` blindly. **No whole-file hashes/fingerprints**: the ZIP
  layer is nondeterministic (timestamps, ordering) and `rId<n>-created`
  counters shift with unrelated changes; and a hash mismatch tells you *that*
  something changed, not *what*.
- 🧪 **Round-trip assertions**: re-open the *written* file with the library's
  own read side (`getInfo()`, `xml-slide-helper`) and assert on what it reports.
  An independent code path as oracle, zero infrastructure.
- 🧪 Unit tests for pure helpers (`modify-text-helper`, `modify-color-helper`,
  `cell-id-helper`, `general-helper`) — pure functions on DOM nodes, cheap to
  test, currently only covered incidentally.

### Tier 1 — package invariants on every written archive (fast, automatic) — ✅ done 2026-08-12

Implemented in `__tests__/helpers/pptx-invariants.ts`; there was no shared
write helper to hook, so `setup-pptx-invariants.ts` (jest `setupFilesAfterEnv`)
wraps `Automizer.prototype.write` — **every** output file is checked without
individual suites opting in, and a violation fails the test that wrote the
file. A self-test (`pptx-invariants.test.ts`) proves each corruption class is
still detected. Checks:

- every `r:id`/`r:embed`/`r:link` referenced in a part resolves to an entry in
  that part's `_rels` file;
- every rel target resolves to an existing part in the archive (and vice versa:
  no orphaned media/chart parts);
- every part is covered by `[Content_Types].xml` (override or default);
- `ppt/presentation.xml` slide list ↔ actual `slide<n>.xml` parts are consistent.

This catches the most common corruption class (dangling relationships) at
near-zero cost and turns the existing 94 suites into 94 validity probes.

Not hypothetical: during the HTML→text track a test wrote an internal slide
link to `slide3.xml` in a two-slide deck. Everything passed, the XML looked
right, and the only symptom was a hyperlink that silently did nothing when the
deck was opened in PowerPoint. A rel-target-resolves check would have failed
that test instantly — this invariant is the one that pays for itself first.
(Confirmed on landing day: the hook immediately caught the same class again in
`htmlToMultiText-hyperlinks.test.ts`, plus one test that does it deliberately —
that one now wraps its write in `withoutPptxInvariants(...)`.)

**Baseline findings (2026-08-12).** Strict checking initially failed 33 suites;
triage showed the dangling rels were almost all *unreferenced* (no `r:*`
attribute points at them), i.e. tolerated bloat rather than corruption. The
checker therefore splits its report: `errors` (a **referenced** rel whose
target is missing, uncovered parts, broken slide list, malformed XML) fail the
test; `knownIssues` are reported but tolerated, mirroring the Tier-2 allowlist
principle — escalate a class to an error only together with the library fix.
Current known-issue classes, all candidates for Phase 7 cleanup work:

- stale rel entries left behind when copied content is re-targeted (slide
  copy, media dedup, `setRelationTarget` swaps) — the copied slide keeps its
  source-template rels although the blips point at the new `-created` ones;
- notesSlides rels reference `notesMasters/notesMaster1.xml`, but a notesMaster
  is never copied to the output;
- slide/media/chart parts orphaned by `removeExistingSlides`/`removeSlide`
  without `cleanup`.

### Tier 2 — OOXML schema validation (CI job) — the "never see the repair prompt again" gate — ✅ done 2026-08-12

The repair-prompt bug class (e.g. `a:pPr` child order, see the HTML→text track)
is invisible to tiers 0/1/3. The genuine oracle is the Open XML SDK validator.

**Implemented Docker-first** (no .NET on any contributor machine or in CI —
same reasoning as Tier 3's pinned renderer): `tools/validate-pptx/` holds the
console app + a multi-stage Dockerfile (`sdk:8.0-alpine` build →
`runtime:8.0-alpine`, image builds in seconds, cached afterwards).
`yarn validate:pptx` builds the image and validates
`__tests__/pptx-templates` **and** `__tests__/pptx-output`; the `validate-pptx`
CI job runs `yarn test` to generate fresh outputs, then the same command.
Allowlist lives in `tools/validate-pptx/allowlist.json` (description + optional
part-URI substring match, comments allowed): baseline template noise (all MS
chart-extension elements, scoped to `/ppt/charts/`), plus **documented library
bugs** (see the bug track below) and upstream pptxgenjs chart quirks — each
library-bug entry must be removed together with its fix. On landing day:
161 files, 0 new errors, 573 allowlisted.

Original tool sketch (kept for reference; the implemented version adds the
allowlist and directory expansion):

  ```csharp
  // tools/validate-pptx/Program.cs
  using DocumentFormat.OpenXml;
  using DocumentFormat.OpenXml.Packaging;
  using DocumentFormat.OpenXml.Validation;

  var errors = 0;
  foreach (var path in args) {
    using var doc = PresentationDocument.Open(path, false);
    var validator = new OpenXmlValidator(FileFormatVersions.Microsoft365);
    foreach (var e in validator.Validate(doc)) {
      errors++;
      Console.WriteLine($"{path} :: [{e.ErrorType}] {e.Description}" +
                        $" @ {e.Part?.Uri} {e.Path?.XPath}");
    }
  }
  return errors == 0 ? 0 : 1;
  ```

  `validate-pptx.csproj` references the `DocumentFormat.OpenXml` NuGet package.
  CI: `actions/setup-dotnet` → `yarn test` → `dotnet run --project
  tools/validate-pptx -- __tests__/pptx-output/*.pptx`.
- ⚠ **Baseline first**: run the validator over `__tests__/pptx-templates/*.pptx`
  before gating — real-world templates often carry pre-existing warnings, and
  the validator flags some things PowerPoint tolerates. Ship a small allowlist
  (per error description/part) seeded from the template baseline; fail only on
  *new* errors. Tighten the allowlist over time.
- Alternative if .NET in CI is unwanted: a `python-pptx` round-trip open is a
  weaker, dependency-lighter proxy (catches package-level breakage, not schema
  order).

### Tier 3 — visual regression via pptx-thumbnailer (curated, separate CI job / nightly)

✅ Done 2026-08-12 (first 3 decks): `yarn test:visual` (Docker only) runs
`__tests__/visual/*.deck.ts` via `jest.visual.config.ts`; its globalSetup
builds `tools/render-pptx/` (digest-pinned `node:22-bookworm-slim` +
`libreoffice-impress` + poppler + Liberation/Carlito, `pptx-thumbnailer@0.1.0`
exact — no published image existed, so the renderer is pinned via base digest
plus exact npm version instead) and starts it on a free port; pixelmatch
threshold 0.1, ≤0.1% differing pixels, failure evidence in
`__tests__/visual-output/` (gitignored). CI job `visual-regression` gates and
uploads actual+diff PNGs as artifact. Decks 1 (partially, via multitext), 2, 3,
4 covered by `multitext-html`, `tables`, `chart-bars`; the other 9 below are
open.

Uses [pptx-thumbnailer](https://github.com/singerla/pptx-thumbnailer)
(headless LibreOffice → PDF → pdftocairo) as a **change detector, not a
correctness oracle**: LibreOffice's <100% PowerPoint fidelity doesn't matter
for regression testing — determinism does. The baseline PNG means "what the
pinned renderer showed when a human last blessed it", not "what PowerPoint
shows". Never conclude PowerPoint-correctness from green pixels.

- 🔧 Mechanics: run the thumbnailer as a service container **pinned by image
  digest** (renderer/font/poppler drift invalidates baselines), fixed `dpi`,
  fonts limited to those shipped in the image (Liberation/Carlito) for the
  golden templates. Compare with a perceptual diff (`odiff` or `pixelmatch`)
  with a small threshold — never byte-equality, antialiasing guarantees noise.
  Baselines live in `__tests__/visual-baselines/<deck>/<slide>.png` (400px
  wide → small enough to commit; GitHub renders image diffs in PRs).
  Update flow: `UPDATE_BASELINES=1 yarn test:visual` regenerates; the PR shows
  before/after PNGs; review is a glance, not a PowerPoint session.
- 🧪 **Do not render all 94 suites** (LibreOffice, concurrency 1 → coffee-break
  CI). Instead ~12 **golden decks**, each a small scripted build (1–5 slides)
  from existing `__tests__/pptx-templates/`, deliberately covering one feature
  area:

  | # | Deck | Exercises |
  |---|---|---|
  | 1 | `text-basics` | setText, replaceText, setBulletList, TextStyle (bold/italic/color/size) |
  | 2 | `multitext-html` | setMultiText + htmlToMultiText (nested lists, `<ol>`, styles) — pins current behavior; the HTML→text feature track updates these baselines intentionally |
  | 3 | `tables` | setTableData grow/shrink rows+cols, cell styles, borders |
  | 4 | `chart-bars` | setChartData add/remove series+categories on bar charts |
  | 5 | `chart-types` | pie, line, scatter, bubble, combo |
  | 6 | `chart-styling` | setAxisRange, setLegendPosition, data labels incl. removal |
  | 7 | `images` | image copy across templates, setPosition/setSize, duplicate-media dedup |
  | 8 | `masters-layouts` | addMaster, useSlideLayout, placeholder normalization |
  | 9 | `generate-bridge` | PptxGenJS-generated shapes next to template shapes |
  | 10 | `hyperlinks` | external + internal (slide-target) links |
  | 11 | `remove-and-order` | removeElement, removeSlide/slide order, normalizePresentation |
  | 12 | `mixed-report` | 3-slide realistic report combining text+table+chart+image |

  A red golden deck answers "which feature area changed visually" in one glance.
- Free byproduct: "LibreOffice converted the deck at all" is itself a weak
  validity smoke test.

### Rollout order

1. ✅ Tier 1 invariants + `expectXml` helper (immediate value, no new infra) —
   done 2026-08-12, incl. AGENTS.md "Testing rules" update.
2. ✅ Tier 2 validator with template-derived allowlist → CI gate — done
   2026-08-12, Dockerized (`yarn validate:pptx`, `validate-pptx` CI job).
   **After this, the repair prompt is a CI failure, not a customer report** —
   for everything not yet on the library-bug allowlist.
3. Tier 0 snapshots retro-fitted per bug fix / feature PR (rule: every bug fix
   adds the XML assertion that would have caught it).
4. ✅ Tier 3 golden decks, starting with `chart-bars`, `tables`,
   `multitext-html` (the areas where visual breakage historically happens) —
   done 2026-08-12; the remaining ~9 decks from the table land incrementally.
5. 📖 As each tier lands, update **AGENTS.md → "Testing rules"** to point at the
   tier model (especially: every bug fix adds the tier-0 assertion that would
   have caught it; never render all suites in tier 3), so agents working in the
   repo follow it automatically.
- 🔧 Wire coverage into CI once the glob is fixed (Phase 0). Coverage is a
  tier-0 metric only — tiers 1–3 verify output, not code paths.

## Phase 6 — Docs & AI enablement

Current state: `README.md` is a 1526-line monolith with 126 code blocks,
`AI-INSTRUCTOR.md` (19 KB) targets AI *consuming* the library, `AGENTS.md`
(12 KB) targets AI *contributing to* the repo. Typedoc is configured but has
never been run (no script, no CI step). Nothing verifies that any documented
example still compiles.

**Guiding principle: one corpus, two renderings.** Human and AI readers agree on
~75% of what they need — the facts: signatures, working examples, constraints,
error semantics. They diverge on *structure*, not content:

| | Human reader | AI reader |
|---|---|---|
| Access pattern | non-linear; TOC, search, returns later | one retrieved chunk, once, no memory |
| Cross-references | follows "see above" fine | usually cannot resolve them |
| Redundancy | irritating | **required** — each chunk needs its own imports and preconditions |
| Negative constraints | nice-to-have footnote | highest-value content; prevents hallucinated APIs |
| Navigation chrome, badges, screenshots | essential | pure token overhead |

The consequence: the AI-facing artifacts must be **generated from** the
human-facing corpus, never hand-maintained beside it. Today's maintenance rule
for `AI-INSTRUCTOR.md` ("update it in the same PR as any `modify.*` change") is
unenforced human discipline, and a stale AI instructor is worse than none — it
teaches wrong APIs authoritatively.

### 6.1 Kill documentation drift first (before any website)

Highest leverage in the phase, and independent of everything below.

- 🧪 **Compile every documented example in CI.** Add
  `__tests__/docs-examples.test.ts`: extract all fenced ` ```ts ` blocks from
  `README.md`, `docs/**/*.md` and `AI-INSTRUCTOR.md`, wrap each in a synthetic
  module, and typecheck the batch with the TS compiler API. Blocks that are
  deliberately partial get ` ```ts ignore `. This turns 126 prose snippets into
  tests and fixes drift for both audiences with one mechanism.
- 🔧 **`typedoc.json` writes into `docs/`**, which is not gitignored and is
  exactly where the docs-site sources want to live. Change `out` to
  `website/docs/api` (generated, gitignored) before the split starts.
- 🔧 Add scripts: `docs:api` (typedoc), `docs:build`, `docs:serve`.
- 📖 **Document the deferred-execution model explicitly** — the single biggest
  user surprise: callbacks run at `write()`, not at `addSlide()`. It belongs in
  the concepts page, restated in the intro of every modifier page (redundancy is
  correct here — see the principle above), and in the rules list.

### 6.2 Split the README

Keep `README.md` to pitch, badges, install, one runnable example, and links out.
Everything else moves to `docs/`, one page per feature area:

| Docs page | Absorbs from README |
|---|---|
| `getting-started` | Installation (package + cloned repo), Basic Example |
| `concepts` | deferred execution, one-instance rule, template/root model, 1-based numbering |
| `selectors` | How to select slides and shapes, creationId, Find and Modify Shapes |
| `text` | Modify Text, Text helpers (MultiText/HTML) |
| `tables` | Modify Tables, Table helpers |
| `charts` | Modify Charts, Extended Charts, Additional chart modifiers, Read chart data |
| `images` | Modify Images, Image helpers |
| `masters-layouts` | Slide Masters and Layouts, Import and modify slide Masters |
| `generation` | Generate shapes with PptxGenJS, hyperlinked shapes |
| `hyperlinks` | Hyperlink Management (add, update, remove) |
| `slide-management` | Remove elements, Sort output slides, loop through slides, get slide numbers |
| `helpers` | Shape/Cleanup/Unit/Generic/Advanced XML helpers |
| `output` | write/stream/getJSZip, StatusTracker, cleanup flags |
| `limitations` | Requirements and Limitations (shape types, chart types, animations, PowerPoint version) |
| `troubleshooting` | Troubleshooting, repair-prompt bisection, Testing |
| `api/` | generated by typedoc |

Do this **after** 6.1, so the example test moves with the content and catches
anything broken in transit.

### 6.3 Docs site

- 🔧 **Docusaurus** in `website/`, chosen over Starlight/VitePress for one
  reason: `docusaurus-plugin-typedoc` folds the existing typedoc config into the
  sidebar, so guides and API reference become one searchable site. Versioning is
  built in and matters for a pre-1.0 library with API churn — cut a `0.8` version
  at the first release after the site lands, keep `next` from `main`.
- 🔧 Search: Pagefind or the built-in local index, built at compile time. No
  hosted search service for a site this size.

### 6.4 AI rendering (generated, not written)

- 📖 `llms.txt` at the site root: flat plaintext index of every page with a
  one-line description. `llms-full.txt`: the whole corpus concatenated. Both
  generated at build time by a Docusaurus plugin — never edited by hand.
- 📖 Serve a `.md` twin of every HTML page at the same path + `.md`.
- 📖 `AI-INSTRUCTOR.md` stays hand-written **only** for the parts that have no
  human equivalent: the mental model, the rules list, and the minimal complete
  example. Everything else in it becomes an include from the docs corpus, and the
  whole file is covered by the 6.1 compile test. Keep shipping it in the npm
  `files` array — that is the offline channel for AI readers.
- 📖 `AGENTS.md` stays in the repo root and off the site. It addresses agents
  editing this repository, not users of the library, and root is where they look
  for it.

### 6.5 Publishing — GitHub Pages *and* self-hosted

Ship both. GitHub Pages is zero-ops and gets the site live in an afternoon;
the self-hosted container is the branded canonical URL and keeps the option of
serving other things from the same box.

- 🔧 **Runtime rule:** the container serves *static files*. No docs framework, no
  Node in the production image. Multi-stage build → `nginx:alpine`, ~50 MB.

  ```dockerfile
  FROM node:22-alpine AS build
  WORKDIR /app
  COPY package.json yarn.lock ./
  RUN yarn install --frozen-lockfile
  COPY . .
  RUN yarn docs:build

  FROM nginx:alpine
  COPY --from=build /app/website/build /usr/share/nginx/html
  ```

- 🔧 **TLS/proxy:** Caddy in front on a shared docker network — automatic Let's
  Encrypt, three lines of Caddyfile for `docs.ensembl.io`. Traefik instead if
  ensembl.io already runs it for other services.
- 🔧 **CI:** on push to `main`, one workflow builds the site twice (different
  `baseUrl`), deploys one build to Pages via `actions/deploy-pages`, and pushes
  the image to `ghcr.io`; the VPS pulls. Two builds are needed because Docusaurus
  bakes `url`/`baseUrl` in — `/pptx-automizer/` for
  `singerla.github.io/pptx-automizer`, `/` for `docs.ensembl.io`. Cheap: a
  matrix job.
- 🔧 **Canonical:** `docs.ensembl.io` is canonical; the Pages build sets
  `rel=canonical` at it and is the mirror/fallback. Avoids splitting search
  ranking across two identical sites.

### Rollout order

1. 6.1 — docs-example test + typedoc `out` fix. Standalone value, no site needed.
2. 6.2 — README split, page by page, example test green throughout.
3. 6.3 — Docusaurus + typedoc, deployed to **GitHub Pages only** first. Prove the
   pipeline on the zero-ops target.
4. 6.5 — add the Docker target and `docs.ensembl.io`, flip canonical.
5. 6.4 — `llms.txt` generation and the `AI-INSTRUCTOR.md` include refactor, once
   the corpus is stable enough to generate from.

## Phase 7 — Feature debt already noted in code (post-refactor)

Collected from TODOs and limitations, ordered by user value:

1. creationIds for slideMasters (`has-shapes.ts` TODO).
2. Modify-in-place for existing slides of the root template (currently
   "add-only"; README documents loop-over-all-slides workaround).
3. `removeDuplicatePlaceholders` can over-remove when >2 placeholders share a
   type (TODO at `has-shapes.ts:1234`) — match by id.
4. `Slide.remove()` is documented broken (`targetNumber` undefined ToDo).
5. Media/video/audio shape support (README "Limitations").
6. Animation id synchronization (README "Limitations").
7. `getLocation`/template resolution for URL sources; browser build is
   explicitly out of scope for now.
8. ~~HTML → text conversion (`modify.htmlToMultiText`) is incomplete and partly
   incorrect~~ — done, see the feature track below.

---

## Feature track — HTML → PPTX text (`htmlToMultiText`) — ✅ done 2026-08-11

Audit date: 2026-08-11. Scope: `src/helper/html-to-multitext-helper.ts`
(HTML → `MultiTextParagraph[]`) and `src/helper/multitext-helper.ts`
(→ DrawingML). Independent of the refactor phases; can proceed in parallel.

**Outcome** (prioritized on customer request): all 10 bugs below are fixed and
the coverage gaps closed. The converter is a single-pass `walk()` over
`(blockCtx, inlineStyle)` accumulators, the CSS subset lives in the standalone
`helper/css-style-parser.ts`, and schema-order insertion is now a shared
primitive (`XmlHelper.insertInSchemaOrder` / `sortChildrenBySchema`) used for
both `a:pPr` and `a:rPr`. New public surface (all additive): `TextStyle` gained
`isStrike`/`fontFamily`/`highlight`, `MultiTextParagraph.paragraph` gained
`bulletType`/`bulletChar`/`autoNumberType`, and a text run may be
`{ break: true }`. Tests: `__tests__/html-to-multitext-converter.test.ts`
(32 conversion-rule assertions) and a rewritten
`__tests__/replace-multi-text-html.test.ts` (6 output-XML suites, replacing the
`// TODO` that asserted nothing).

Two behavior changes worth knowing about:

- `extractDefaultStyle` no longer inherits bold/italic from the template's first
  run (bug 10), which changed one `a:br` assertion in
  `replace-multi-text-linebreaks.test.ts`.
- Paragraph alignment is only written when the HTML specifies `text-align`;
  previously every paragraph was forced to `algn="l"`, overriding the layout.

**Resolved scoping question:** the input contract is *WYSIWYG editor output*
(CKEditor/TinyMCE). No `htmlparser2` dependency was added — `@xmldom/xmldom` in
`text/html` mode turned out to handle the cases that mattered (`&nbsp;`, bare
`&`, unclosed `<br>`); CKEditor's invalid sibling-nested lists are handled in
the converter instead. If arbitrary web tag soup ever becomes a requirement,
swapping in a forgiving parse layer is a change to `run()` alone.

**Guiding principle:** PPTX text is strictly flat and two-level — a `txBody` is
a flat list of `<a:p>` paragraphs, each a flat list of `<a:r>` runs. There is no
nesting anywhere. All HTML hierarchy must be *projected* onto this:

- nested inline tags (`<strong><em>…`) → one run with **accumulated** character
  properties (the current style-accumulation approach is conceptually right);
- nested lists → flat paragraphs with a 0-based `lvl` attribute (0–8) — PPTX has
  no "list object", a list item is just a paragraph with bullet properties;
- block-inside-block (`<li><p>…</p>`, `<blockquote><p>…`) → the **innermost**
  block opens a new flat paragraph; ancestor blocks only contribute properties
  (indent level, spacing), never structure.

### Bugs — converter (`html-to-multitext-helper.ts`)

1. 🐛 **Level off-by-one**: top-level `<ul>` items get `level: 1`, but PPTX
   `lvl` is 0-based → every bullet renders one indent level too deep relative to
   the layout's `lstStyle`. Plain `<p>` gets `level: 0`, colliding with what
   should be first-level bullets.
2. 🐛 **Two divergent list mechanisms**: sibling-nested lists
   (`<ul><li/><ul>…` — CKEditor-style invalid nesting, what the test uses) go
   through a shared mutable `bulletLevel` counter; properly nested HTML
   (`<li>text<ul>…</ul></li>`) goes through `level + 1` recursion with a freshly
   created counter in `processListItem`. Both input shapes must normalize to the
   same output; today they diverge.
3. 🐛 **Dropped content**: bare text directly in `<body>` or `<div>` is silently
   discarded (the `default:` branch of `processNode` recurses but never emits);
   `<li><p>text</p></li>` loses the `<p>`'s block semantics.
4. 🐛 **Color passed verbatim**: `color: #ff0000` lands as
   `<a:srgbClr val="#ff0000">` — the leading `#` is invalid OOXML. Named colors
   and `rgb()` are unhandled.
5. 🐛 **px treated as pt**: `font-size: 12px` → `12 * 100`; 1px = 0.75pt at
   96dpi. Only integer `px` matches — `pt`, decimals, `em` don't.
6. 🐛 **`<sub>`/`<sup>` never mapped** although `TextStyle.isSubscript` /
   `isSuperscript` exist.
7. 🐛 **`<ol>` renders as `•` bullets** — should emit `a:buAutoNum`
   (`arabicPeriod` etc., type varying per nesting depth).

### Bugs — renderer (`multitext-helper.ts`)

8. 🐛 **`a:pPr` child order violates the OOXML schema**: `buChar`/`buNone` are
   appended *before* `lnSpc`/`spcBef`/`spcAft`; the schema requires spacing
   first, then `bu*`. This is the "PowerPoint repair prompt" bug class. Audit
   `a:rPr` child ordering while there.
9. 🐛 **Indentation constant regardless of level**: `marL`/`indent` fixed at
   228600 — nested bullets render at the same x-position when the target shape's
   layout has no per-level `lstStyle`. Use `marL = (level + 1) * 228600`,
   `indent = -228600` (or make it configurable).
10. ⚠ **`extractDefaultStyle` bleeds template styling**: size/color/bold/italic
    of the template's *first run* are merged into every generated run — a bold
    placeholder makes all HTML text bold. Restrict to size/color or make opt-in.

### Missing coverage

- **Tags**: `<br>` (needs an `a:br` concept in the run model — e.g.
  `{ break: true }` in `MultiTextParagraph.textRuns` — plus `MultiTextHelper`
  support), `<u>` (only `<ins>` is mapped), `<s>`/`<strike>`/`<del>` → strike,
  `h1`–`h6` (paragraph + size/bold), `<code>`/`<pre>` (monospace), `<blockquote>`.
- **CSS**: `text-align` on `p`/`li` (alignment is hardcoded `'l'`),
  `font-weight`/`font-style`/`text-decoration`, `font-family`,
  `background-color` (highlight).
- **Whitespace**: no CSS-style collapsing; literal `\n` inside a paragraph ends
  up inside `a:t`.

### Plan (ordered)

1. 🏗 **Rewrite the converter traversal as a single pass**:
   `walk(node, blockCtx, inlineStyle)` with two accumulators — an inline style
   context (bold/italic/color/…) and a block context (list depth, list type,
   alignment). Inline tags/CSS only extend `inlineStyle`; `ul`/`ol` push list
   depth+type; `p`/`li`/`h*`/`div`/`blockquote` flush the current paragraph and
   start a new one derived from `blockCtx`. Fixes bugs 1–3 structurally, handles
   both list-nesting shapes identically. Keep `MultiTextParagraph` as the output
   contract (public API).
2. 🏗 **Widen the mapping tables** (mechanical once 1 is in): tag map, CSS
   parser (font-size px/pt, color normalization to 6-digit hex incl. names and
   `rgb()`, text-align, font-weight/style/decoration, font-family), `<ol>` →
   `buAutoNum`, whitespace collapsing. Fixes 4–7 + coverage gaps.
3. 🏗 **Fix the renderer**: `a:pPr` child ordering, per-level `marL`, `a:br`
   run support, tame `extractDefaultStyle`. Fixes 8–10.
4. 🔧 **Parser hardening + tests**: `@xmldom/xmldom` is an XML parser —
   `parseFromString(html, 'text/html')` breaks on real-world tag soup (unclosed
   `<br>`, unquoted attributes, `&nbsp;`). Decide: add `htmlparser2` (small,
   forgiving) as the parse layer, or document a strict-XHTML input contract.
   Add XML-level assertions via the Phase 5 output-assertion helper — the
   existing `replace-multi-text-html.test.ts` writes a file with a `// TODO`
   instead of asserting, which is the biggest gap for iterating safely here.

Open scoping question: **where does the HTML come from** (CKEditor/TinyMCE
output vs. arbitrary user HTML vs. an own generator)? That decides how defensive
the parser must be and which CSS subset is worth supporting.

---

## Bug track — chart data-point styling creates `<c:dPt>` that erase line charts

Audit date: 2026-08-12. Found while investigating an ensemblio report where all
lines of a 9-series line chart disappeared. Independent of the refactor phases;
the behaviour is old (`XmlElements.dataPoint()` dates back to 2021) and
reproduces identically on `0.8.2` and on the current refactor branch — so this
is a genuine library bug, not refactor fallout.

**Symptom.** Every data point of every series gets a fabricated

```xml
<c:dPt><c:idx val="0"/><c:spPr>
  <a:solidFill><a:srgbClr val="CCCCCC"/></a:solidFill>
  <a:ln><a:noFill/></a:ln>
  <a:effectLst/>
</c:spPr></c:dPt>
```

On a bar/pie chart that only turns the points grey; on a **line** chart
`<a:ln><a:noFill/>` per point removes the line segments, and the chart renders
as data labels floating in empty space.

**Cause chain.**

1. `ModifyChart.chartPoint()` (`modify/modify-chart.ts`) decides "this point is
   styled" from *key presence* — `if (!style?.color && !style?.border &&
   !style?.marker) return;`. A caller-supplied `{ marker: { color: undefined } }`
   (or `{ color: { value } }` without `type`) is truthy, so a `c:dPt`
   modification is emitted although none of `chartPointFill/Border/Marker`
   yields an applicable tag.
2. `ModifyXmlHelper.assertElement` finds no `c:dPt` and calls
   `XmlElements.dataPoint()`, which appends `idx` + a **default** `spPr` from
   `XmlElements.spPr()` — grey `solidFill`, `a:ln` `noFill`, `effectLst`. A
   data point that carries no explicit style should inherit the series
   formatting, so fabricating one is wrong regardless of chart type.
3. `chartPointMarker` targets a `c:marker` *inside* the new `c:dPt`;
   `ModifyXmlHelper.createElement` does not handle that tag, so the
   modification is silently dropped and the point keeps the pure default.

**Fixes.**

1. `chartPoint()`: build the child tags first and return `undefined` when
   `chartPointFill`/`chartPointBorder`/`chartPointMarker` all produced nothing,
   instead of testing key presence. No empty style may reach the XML layer.
2. `XmlElements.dataPoint()`: create the `c:dPt` **empty** (`c:idx` only, plus
   `c:bubble3D` for schema order) and let the actual modifications add `c:spPr`
   via the existing `createElement('c:spPr')` path. Nothing should be styled
   that the caller did not ask for.
3. `chartPointMarker`: either support creating `c:marker` inside a `c:dPt`
   (`XmlElements.marker()` + a `createElement` case) or document that per-point
   marker styling is unsupported and drop the tag — today it is a silent no-op
   whose only effect is bug 2.
4. `Modification.isRequired: false` currently only suppresses a (commented-out)
   warning; `assertElement` creates the element either way. Make it actually
   mean "modify if present, never create" — `seriesStyle()` already passes it
   for `c:marker`/`c:spPr` in that intent.

**Guard.** Tier-0 assertion (Phase 5): modify a line chart with per-point styles
and assert no `c:ser` gains a `c:dPt` containing `<a:ln><a:noFill/>` unless the
caller asked for it; plus a tier-3 golden deck for a multi-series line chart,
which is exactly the failure mode a pixel diff catches and an XML diff hides.

**Related, already fixed.** `ModifyColorHelper.normalizeColorObject` used to
throw a `TypeError` on a `Color` without `value` (`color.value.indexOf`), which
aborted `setPointStyles` after the first point. Phase 1's hardening removed the
crash — but that means such a style now silently reaches *all* points instead of
one. Worth a CHANGELOG note when the refactor is published.

---

## Bug track — schema violations found by the Tier-2 validator

Audit date: 2026-08-12, first full run of `yarn validate:pptx` over all test
outputs. Each class is allowlisted in `tools/validate-pptx/allowlist.json` so
CI can gate on *regressions* today; **fixing a class means: fix the XML, remove
the allowlist entry, add the tier-0 assertion that would have caught it.**
Ordered by user impact:

1. 🐛 **Dangling relationships make the Open XML SDK refuse to open ~18 output
   files** (`OpenFailed: Specified part does not exist in the package`). These
   are the Tier-1 `knownIssues` classes — stale rels after copy/dedup/
   re-target, the never-copied notesMaster, orphaned parts without `cleanup`.
   PowerPoint tolerates them; the SDK (and probably other strict consumers,
   e.g. Apache POI) does not. Fixing this unlocks real Tier-2 coverage for
   those files and lets Tier 1 escalate its `knownIssues` to errors — the
   single highest-leverage cleanup in this list.
2. 🐛 **Chart modifiers insert children out of schema order**: `c:dPt` after
   `c:dLbls` in `c:ser` (`modify-existing-chart-styled`, related to the
   `<c:dPt>` bug track above), `c:tx` misplaced inside `c:dLbl`
   (`modify-chart-datalabels-text`), `c:dLbls` misplaced in scatter `c:ser`
   and `a:solidFill` misplaced in `c:dLbl`/`c:spPr`
   (`modify-chart-datalabels`). The shared `XmlHelper.insertInSchemaOrder`
   primitive from the HTML→text track is the intended fix vehicle.
3. 🐛 **Table cell borders written in caller order**: `a:tcPr` requires
   `lnL → lnR → lnT → lnB`; `setTable`/`setTableData` append in the order the
   caller lists them (`modify-existing-table-format-cells`).
4. 🐛 **Bullet-list replacement leaves `a:bodyPr` after `a:p`** in `p:txBody`
   (`replace-bullet-text`, `replace-nested-bullets`) — `bodyPr` must remain
   the first child.
5. 🐛 **Master auto-import registers slideLayout rels on
   `ppt/presentation.xml`** (`modify-master-add-external-image`) — layouts may
   only be related to their slideMaster.
6. 🐛 **Copied diagrams get the wrong content type**: the drawing part is
   re-registered as `application/vnd.openxmlformats-officedocument.…` although
   source templates (and the spec) use
   `application/vnd.ms-office.drawingml.diagramDrawing+xml`
   (`add-slide-diagrams`) — likely a hardcoded type in the content-type
   registration instead of copying the source template's.
7. ⚠ **Upstream, not ours**: pptxgenjs generates line charts with
   `c:varyColors` before `c:grouping` and `c:invertIfNegative` in line series
   (`generate-pptxgenjs-charts`). Revisit the allowlist entries on pptxgenjs
   upgrades.

---

## Suggested sequencing

```
Week 1:   Phase 0 (all) + CI green on main
Weeks 2-3: Phase 1 (errors/logging) — small PRs, mechanical
Weeks 3-6: Phase 2 (HasShapes decomposition + PptPaths) — one extraction per PR,
           integration suite as the safety net
Then:     Phase 3 (globals) → Phase 4 (templates/strict) as background chores
Early:    Phase 5 tiers 1+2 (invariants + schema validator) right after CI exists —
          they harden every later refactor PR at fixed cost
Ongoing:  Phase 5 tier-0 assertions added with every bug fix; tier-3 golden
          decks feature area by feature area
Early:    Phase 6.1 (docs-example compile test) alongside Phase 5 tier 1 — it is
          the same kind of cheap always-on gate, and it must exist before the
          README split moves 126 examples around
Then:     Phase 6.2-6.5 docs split → site on GitHub Pages → self-hosted mirror
```

Rule of thumb for every PR during the refactor: **no public `modify.*` signature
changes**, integration suite stays green, and any behavior change (especially
error behavior) gets a CHANGELOG note.
