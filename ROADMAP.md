# ROADMAP — Architecture audit & improvement plan

Audit date: 2026-08-11 (v0.8.2). Goal: consolidate architecture **before**
adding further functionality. Findings are ordered into phases; each phase is
shippable on its own and phases 0–2 unblock the rest.

Legend: 🐛 bug · 🔧 tooling · 🏗 architecture · 🧪 testing · 📖 docs · ⚡ performance

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

### 6.1 Kill documentation drift first (before any website) — ✅ done 2026-08-13

Highest leverage in the phase, and independent of everything below.

- 🧪 ✅ **Compile every documented example in CI** —
  `__tests__/docs-examples.test.ts` extracts all fenced ` ```ts ` blocks from
  `README.md`, `AI-INSTRUCTOR.md` and `docs/**/*.md`, wraps each in a synthetic
  module (`export {};` prefix — block-local scope + top-level await) and
  typechecks the batch in **one** TS program, so src's type graph is built once
  (~2 s inside the normal `yarn test`). Deliberately partial blocks opt out via
  ` ```ts ignore ` — currently **zero** need it: the docs' conventional context
  names (`pres`, `slide`, `modify`, the `Modify*Helper` classes, …) are ambient
  globals in `__tests__/helpers/docs-example-context.d.ts`, typed straight from
  src, and snippet-local declarations shadow them. Baseline: 73 blocks, 12 were
  broken. Genuine drift fixed: `replaceText` documented a fictional
  regex/whole-word/case API; `setSolidFill` documented a color argument it
  never had (both README and AI-INSTRUCTOR); `ModifyCleanupHelper.clearTextColor`
  was shown bare in a callback array, where the `relation` element lands in its
  `color` parameter at runtime; `expand.mode: 'col'` vs `'column'`;
  `dLblPos: 'outEnd'` vs the `LabelPosition` enum; internal import paths
  (`./types/types`, `./helper/modify-presentation-helper`,
  `pptx-automizer/src/...`) that resolve for no npm consumer. The last class is
  fixed properly: `ModifyPresentationHelper`, `XmlSlideHelper`,
  `XmlRelationshipHelper` and `FindElementSelector` are now exported from
  `src/index.ts` (additive).
- 🔧 ✅ `typedoc.json` `out` → `website/docs/api` (gitignored); first actual
  typedoc run works (0 errors, 51 warnings about unexported referenced types —
  a cleanup candidate for 6.3).
- 🔧 ✅ Script `docs:api` (typedoc) added; `docs:build`/`docs:serve` deferred to
  6.3 — they need the Docusaurus scaffold to exist.
- 📖 ✅ **Deferred-execution model documented** where users look today: new
  "Deferred execution" section in README (before the basic example);
  AI-INSTRUCTOR already taught it ("Key rule" in the mental model). Restating it
  on every modifier page happens with the 6.2 split. Also fixed stale AGENTS.md
  claims while there: callback errors reject `write()` since Phase 1 (not
  "swallowed"), CI exists, tier 2 catches schema-order bugs.

### 6.2 Split the README — ✅ done 2026-08-14

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

Implementation notes (done after 6.3, since the scaffold didn't need the
content): content moved verbatim wherever possible — the docs-examples test
covers every page, 77 blocks green. Deviations from the table above:
"Find all text elements on a slide" landed in `selectors` (it is about
addressing shapes), and the Table/Text/Image helper subsections landed on their
feature pages with pointers from `helpers` (one topic, one page). `concepts`
additionally states the one-instance rule and 1-based numbering in
AI-INSTRUCTOR's wording — the README never said them explicitly. `docs/index.md`
is now the real landing page (pitch + quick example + nav + ecosystem), no
longer a stub. README kept pitch/badges (npm, CI, license — newly added)/
install/the quick example/Special Thanks and links every docs page. The
`troubleshooting` page gained the repair-prompt bisection steps from
AI-INSTRUCTOR rule 7. Added beyond the table (2026-08-14, user request): a
`testing` page documenting the Phase-5 tooling for humans — `yarn test` +
automatic invariants, `yarn validate:pptx` (Docker OOXML schema validation),
`yarn test:visual` (pptx-thumbnailer visual regression) — previously only in
the agent-facing AGENTS.md, hence invisible to the site search.

### 6.3 Docs site — ✅ done 2026-08-14

- 🔧 ✅ **Docusaurus** in `website/` (own yarn workspace, Docusaurus 3.10),
  chosen over Starlight/VitePress for one reason: `docusaurus-plugin-typedoc`
  folds the typedoc config into the sidebar, so guides and API reference become
  one searchable site. Versioning is built in and matters for a pre-1.0 library
  with API churn — cut a version at the first release after the site lands,
  keep `next` from `main`. Implementation notes:
  - **The docs corpus stays at repo-root `docs/`** (Docusaurus points its docs
    plugin at `../docs`, docs-only mode, `routeBasePath: '/'`): it keeps the
    6.1 compile test's glob unchanged and the pages readable on GitHub. The 6.2
    split lands there, one sidebar id per page in `website/sidebars.ts`.
  - Typedoc now emits **markdown** (`typedoc-plugin-markdown` +
    `typedoc-docusaurus-theme`) into `docs/api` (gitignored, excluded from the
    docs-examples walk as generated content). Root `typedoc.json` is gone —
    the options live in the plugin block of `website/docusaurus.config.ts`,
    and root `yarn docs:api` delegates to `docusaurus generate-typedoc`.
    Generation also runs automatically on every `docs:start`/`docs:build`.
  - `.md` is parsed as **CommonMark, not MDX** (`markdown.format: 'detect'`):
    src doc comments and future 6.2 pages cite raw OOXML tags (`<a:ln>`,
    `Promise<TemplateInfo[]>`) that MDX rejects as malformed JSX. Opt into MDX
    per-file via `.mdx`. `sanitizeComments: true` escapes them in typedoc
    output for good measure.
  - The 51 typedoc warnings from 6.1 are gone in the markdown pipeline; the two
    real ones left (`@order` instead of `@param` in
    `modify-presentation-helper.ts`) fixed in src.
  - `docs/index.md` is a **temporary landing page** (install + quick example,
    compile-tested like everything else) that links to the README until the
    6.2 split replaces it.
- 🔧 ✅ Search: `@easyops-cn/docusaurus-search-local` — local index built at
  compile time. No hosted search service for a site this size.
- 🔧 ✅ CI: `.github/workflows/docs.yml` deploys to **GitHub Pages** on push to
  `main` (`actions/deploy-pages`; needs Pages → Source: "GitHub Actions" in the
  repo settings, one-time). `url`/`baseUrl` default to
  `singerla.github.io/pptx-automizer` and are env-overridable
  (`DOCS_URL`/`DOCS_BASE_URL`) for builds targeting another host.

### 6.4 AI rendering (generated, not written) — ✅ done 2026-08-14

- 📖 ✅ `llms.txt` at the site root: flat plaintext index of every page with a
  one-line description. `llms-full.txt`: the whole corpus concatenated. Both
  generated at build time by a Docusaurus plugin — never edited by hand.
- 📖 ✅ Serve a `.md` twin of every HTML page at the same path + `.md`.
- 📖 ✅ `AI-INSTRUCTOR.md` stays hand-written **only** for the parts that have no
  human equivalent: the mental model, the rules list, and the minimal complete
  example. Everything else in it becomes an include from the docs corpus, and the
  whole file is covered by the 6.1 compile test. Keep shipping it in the npm
  `files` array — that is the offline channel for AI readers.
- 📖 ✅ `AGENTS.md` stays in the repo root and off the site. It addresses agents
  editing this repository, not users of the library, and root is where they look
  for it. (Nothing to do.)

Implementation notes (2026-08-14):

- **Local plugin, no new dependency**: `website/plugins/llms-txt.ts`, a
  `postBuild` hook reading the docs plugin's loaded content — routes, titles
  and sidebar order come from Docusaurus itself, so a new page only needs its
  frontmatter and a sidebar id. Twins mirror the `trailingSlash: false` route
  scheme (`/charts` → `charts.md`; the category-index route `/api` →
  `api.md`, and relative `…/index.md` links are rewritten accordingly).
  Guides + all 65 typedoc pages get twins; `llms-full.txt` is guides-only
  (the generated API reference is indexed via `llms.txt` instead). A guide
  page without a frontmatter `description` **fails the build** — that
  one-liner is the llms.txt payload. Deploy needed no CI change (the Pages
  workflow uploads `website/build` wholesale).
- **AI-INSTRUCTOR.md is now a build artifact**: `tools/ai-instructor/build.ts`
  assembles it from `template.md` (mental model, rules, minimal example) plus
  `<!-- include: docs/<page>.md -->` whole-page and `… § <Heading> -->`
  section directives (heading levels shifted fence-aware, relative docs links
  absolutized to the Pages twins). `yarn docs:ai` regenerates;
  `__tests__/ai-instructor.test.ts` fails `yarn test` on drift, and the 6.1
  compile test covers the generated file's blocks as before. The file grew
  ~500 → ~1900 lines: the full feature pages replace the condensed cheat
  sheet.
- **Corpus became the single source first**: facts that existed only in
  AI-INSTRUCTOR moved into the docs pages — raw-XML-callback rules + worked
  outline example + units reference → `helpers.md`; `Buffer` template loading
  → `getting-started.md`; `getAllElements`/`getDimensions` → `selectors.md`;
  PptxGenJS inches note → `generation.md`; `loadMediaBuffer` → `images.md`.
  Two stale claims died in the process (CHANGELOG'd): "callback exceptions
  are swallowed" (rejects `write()` since Phase 1) and "don't run two builds
  concurrently" (`concepts.md` still said it; concurrency is supported and
  regression-tested since Phase 3).

### 6.5 Publishing — ✅ resolved 2026-08-14: GitHub Pages only

Originally planned as GitHub Pages *plus* a self-hosted branded mirror.
**Dropped 2026-08-14 by decision:** the docs stay on GitHub Pages exclusively,
and the open-source project is not linked to any commercial offering — no
second host, no canonical flip. `singerla.github.io/pptx-automizer` is the one
and only docs URL. The Pages deploy has been live since 6.3; nothing further to
do. (`DOCS_URL`/`DOCS_BASE_URL` env overrides remain in
`website/docusaurus.config.ts` as a generic escape hatch for anyone building
the site for another host — unused by CI.)

### Rollout order

1. ✅ 6.1 — docs-example test + typedoc `out` fix. Standalone value, no site
   needed. Done 2026-08-13.
2. ✅ 6.2 — README split, page by page, example test green throughout. Done
   2026-08-14 (after 6.3).
3. ✅ 6.3 — Docusaurus + typedoc, deployed to **GitHub Pages only** first. Prove
   the pipeline on the zero-ops target. Done 2026-08-14, ahead of 6.2 — the
   scaffold doesn't need the split content, the split pages drop into `docs/`
   and `website/sidebars.ts` as they land.
4. ✅ 6.5 — resolved 2026-08-14: GitHub Pages stays the only host, the
   self-hosted mirror is dropped (see above).
5. ✅ 6.4 — `llms.txt` generation and the `AI-INSTRUCTOR.md` include refactor,
   once the corpus is stable enough to generate from. Done 2026-08-14 — with
   this, Phase 6 is complete.

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

## Bug track — chart data-point styling creates `<c:dPt>` that erase line charts — ✅ fixed 2026-08-14

Audit date: 2026-08-12. Found while investigating a downstream report where all
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

1. ✅ (2026-08-14) `chartPoint()`: build the child tags first and return
   `undefined` when fill/border/marker all produced nothing, instead of
   testing key presence. No empty style may reach the XML layer —
   `setPointStyles` skips the `modify()` call entirely when nothing remains.
2. ✅ (2026-08-12, with the Modification-contract track) `XmlElements.dataPoint()`
   creates the `c:dPt` **empty** (`c:idx` plus `c:bubble3D`); modifications add
   `c:spPr` via the `createElement('c:spPr')` path.
3. ✅ (2026-08-14) `chartPointMarker`: per-point marker styling is
   **modify-if-present** (`isRequired: false` on `c:marker`, honored since the
   contract track) and gated on a typed color — a `marker` key without one
   contributes nothing instead of emitting a no-op tag whose only effect was
   bug 2.
4. ✅ (2026-08-12, contract track) `isRequired: false` means "modify if
   present, never create" in `assertElement`.

Closing the track (2026-08-14) surfaced that the **remaining payload lived in
`createElement('c:spPr')` itself**: `XmlElements.shapeProperties()` built the
grey `solidFill` + `<a:ln><a:noFill/>` + `effectLst` blob, so even a genuine
per-point fill erased its line segment on creation. It now creates an **empty**
`c:spPr` in schema position (new `C_DPT_CHILD_ORDER` — note `c:bubble3D`
sorts *before* `c:spPr` inside `c:dPt`, unlike `c:ser` — and reusing
`C_SER_CHILD_ORDER` for series). This also covers the series-level twin:
`seriesStyle`'s default-required `c:spPr` fabricated the same blob into any
styled series lacking one. Per-point borders no longer depend on the blob's
`a:ln`: `chartPointBorder` nests inside the point's single `c:spPr` entry
(`chartPointSpPr`) and a new `createElement('a:ln')` case inserts a bare
`<a:ln>` per `A_SPPR_CHILD_ORDER`. The dead `XmlElements.spPr()` is deleted.

**Guard.** ✅ Tier-0: `modify-chart-point-styles-line.test.ts` — no template
ships a line chart (why the bug class stayed invisible), so it generates a
3-series one via pptxgenjs and re-loads it through `setChartData`: a real
style materializes exactly one `c:dPt` with the requested fill, pseudo-styles
(`{ marker: {} }`, a color without `type`) produce none, and no `c:dPt`
carries an `<a:noFill/>` the caller didn't ask for. ✅ Tier-3 golden deck
`chart-line-points`: 5-series line chart with sparse real + pseudo point
styles — the baseline pins "all lines visible", exactly the failure mode a
pixel diff catches and an XML diff hides.

**Related, already fixed.** `ModifyColorHelper.normalizeColorObject` used to
throw a `TypeError` on a `Color` without `value` (`color.value.indexOf`), which
aborted `setPointStyles` after the first point. Phase 1's hardening removed the
crash — but that means such a style now silently reaches *all* points instead of
one. ✅ noted in `CHANGELOG.md` under Unreleased → Changed.

---

## Bug track — silent `Modification` contract violations in chart modifiers — ✅ fixed 2026-08-12

Audit date: 2026-08-12. Found by diffing two renderings of the *same*
report against **byte-identical input data** —
an older deck versus one generated on `0.8.2`. Two independent, customer-visible
regressions, both in chart modifiers, both silent, both already released. They
are separate defects with separate fixes, but they share one root cause with
each other and with the `<c:dPt>` track above: the `Modification` type's two
control fields — `index` and `isRequired` — have **no documented contract and no
enforcement**, so a plausible-looking edit silently changes what the XML layer
does.

Both were verified by reproduction against the customer's own templates; the
evidence below is reproducible from `__tests__` alone.

### A. 🐛 `setPointStyles` addresses `c:dPt` / `c:dLbl` by category index

Regression introduced in `3d6c452` (2026-06-24, "feat(chart): add functionality
to remove data labels from charts"), released in **v0.8.2**. The commit is
otherwise a *stop fabricating data-label XML* cleanup; this hunk is unrelated to
that goal and is not mentioned in the message.

```diff
-          count[s] = !count[s] ? 0 : count[s];
-          labelCount[s] = !labelCount[s] ? 0 : labelCount[s];
           this.chart.modify(
-            this.series(s, this.chartPoint(count[s], c, style)),
+            this.series(s, this.chartPoint(c, c, style)),
           );
           if (style.label) {
             this.chart.modify(
-              this.series(s, this.chartPointLabel(labelCount[s], c, style.label)),
+              this.series(s, this.chartPointLabel(c, c, style.label)),
             );
-            labelCount[s]++;
           }
-          count[s]++;
```

**Symptom.** Per-point styles are silently dropped and replaced by a duplicate
of the previously styled point. In the customer deck: a red point at category 0
and a green point at category 3 produce **two red `c:dPt`, both `c:idx val="0"`**
— the green is gone. Deck-wide, **879 of 1224 charts** carry duplicated `c:dPt`
and **60 of 2033 series** duplicated `c:dLbl` indices.

**Why the change looked right.** `chartPoint(count[s], c, style)` passes two
different numbers for what reads as one concept ("the point being styled"), so
collapsing them to `c` looks like deleting bookkeeping cruft — and the diff gets
shorter. The type says only `index?: number` with the comment *"Specify an index
if not 0"*. There was also a real bug next door: `chartPointLabel` addresses
`c:dLbl` elements that a template usually already carries for *specific* points,
so a counter over *styled* points edits an unrelated label and rewrites its
`c:idx` — an index mismatch for which "use the category index" is the natural
conclusion.

**Why it is wrong.**

1. `c:dPt` / `c:dLbl` are **sparse** in OOXML: one element per *explicitly
   styled* point, not one per category. Sibling position carries no meaning;
   `c:idx` is the payload naming the category.
2. `Modification.index` is **positional over existing siblings** — the same
   meaning it has for `c:ser`. `chartPoint(count[s], c, …)` was correct by
   construction: *take or create the next `c:dPt` slot, stamp `c:idx = c`*.
3. `assertElement` can only grow a collection **by one clone per call**. A
   positional index is therefore only satisfiable if the caller walks slots
   0, 1, 2, … in order — exactly the invariant `count[s]` encoded. Passing `c`
   breaks it as soon as the styled categories are not contiguous from 0.

**Why nothing caught it.** The failure is silent (`assertElement` returns
`false`, `modify()` swallows it, the diagnostic is a commented-out `vd(…)`), and
the clone it already inserted stays — so the output *grows*, which does not read
like a dropped modification. The only test exercising point styles,
`modify-existing-chart-styled.test.ts`, asserts `expect(result.charts).toBe(3)`
and nothing about XML. Its own demo data is sparse (`styles: [null, {…}]`), so
**the library's showcase output is wrong today**:

```
ser 0  [(idx 0, 333333)]                     ← correct
ser 1  [(idx 0, CCCCCC)]                     ← should be efefef at idx 1
ser 2  [(idx 0, CCCCCC), (idx 0, CCCCCC)]    ← should be eecc00 @1 and eeccff @3
```

Restoring the counters makes that output correct again — but see the fix below,
a plain revert re-introduces the `c:dLbl` mis-targeting the commit was chasing.
PowerPoint collapses duplicate `c:dPt` with equal `c:idx` on save, which is why
the artifact survives review even when a broken deck is opened.

### B. 🐛 `seriesDataLabel` fabricates a `c:dLbls` in every styled series

Regression introduced in `0a45454` (2026-03-18, "custom data label suffixes and
alpha transparency"), released in **v0.8.2**. Independent of A and three months
older.

**Symptom.** A radar chart whose template deliberately carries *no* data labels
gains one `c:dLbls` per series — ten of them — each containing the hardcoded
blob from `src/helper/xml/dLbl.ts` verbatim, fingerprint included:

```xml
<c:dLbl><c:idx val="0"/>
  …<a:defRPr sz="1400" …><a:solidFill><a:schemeClr val="accent1"/>…
  <c:showVal val="1"/>
  <c16:uniqueId val="{00000001-04B4-49A4-AD60-4DBFE8A0F479}"/>
```

Result: a 14 pt accent1 label switched on for category 0 of every series, over
a template that asked for none.

**Cause.** `ModifyChart.seriesDataLabel()` has declared its intent since
`f68ca16` (2022-05-18):

```js
'c:dLbls': { isRequired: false, children: { 'a:pPr': … } }
```

*Modify label formatting if the template has labels; never create them.* But
`isRequired: false` has never been honoured (fix 4 of the `<c:dPt>` track above).
That stayed harmless only because `createElement()` had no case for the tag —
the assert failed silently and the chart was left alone. `0a45454` added

```diff
+      case 'c:dLbls':
+        new XmlElements(parent).dataPointLabels();
```

which turns the ignored flag into its opposite: *never create* now means *create
an opinionated default*.

**Reproduction** (clean, from the customer template but the shape is generic):
modify a radar chart with `series[].style.label` set → 10 fabricated `c:dLbls`;
without it → 0. Matches the before/after decks exactly.

### C. 🏗 Shared root cause — make the `Modification` contract explicit

Both regressions, plus fixes 2 and 4 of the `<c:dPt>` track, are the same
weakness: `ModifyXmlHelper` silently *creates* and silently *fails*, and its
contract lives only in the heads of the call sites.

- `index` is positional-over-existing-siblings and can only grow a collection by
  one per call. Nothing says so.
- `isRequired: false` reads as "modify if present, never create" at every call
  site and does nothing.
- A failed assert produces no diagnostic (the `vd(…)` is commented out) and may
  leave a stray clone behind.
- `createElement()` injects opinionated defaults (grey `spPr` for `c:dPt`,
  14 pt/accent1/`showVal=1` for `c:dLbls`) rather than empty shells.

**Fixes, in implementation order** — 1 is the smallest change with the widest
effect and unblocks the rest:

1. ✅ **Honour `isRequired: false` in `assertElement`**: when false and the element
   is absent, return `false` *without* creating or cloning. Fixes B outright,
   fixes `<c:dPt>` track item 4, and stops fabrication generally. Audit every
   existing `isRequired: false` call site first (`seriesDataLabel`,
   `seriesStyle`, `setDataPointLabelAttributes`'s nine `c:show*`/`c:numFmt`/
   `c:dLblPos` entries) — some of them may today *depend* on the creation they
   ask not to have. *Audit outcome: none depended on creation (`c:marker`,
   `a:defRPr`, `c15:datalabelsRange` and the nine `c:show*` entries have no
   `createElement` case). Two call sites got the flag added because their
   default-required children fabricated content into label-less series:
   `setDataLabelAttributes`' `c:dLbls` wrapper and `setDataPointLabelAttributes`'
   `c:spPr` (grey-blob creation).*
2. ✅ **Re-enable the diagnostic**: the "Could not assert required tag" path logs
   at `warn` for `isRequired: true` and `debug` for skipped optional tags. The
   default test run stays silent.
3. ✅ **Fix `setPointStyles` addressing (A)** — `Modification.matchIdx` resolves
   `c:dPt` / `c:dLbl` by **matching the child `c:idx` value**, creating one in
   `c:idx` order when absent (`ModifyXmlHelper.assertElementByIdx`); `c:dLbl`
   creation clones the clean `fromIndex` template, whose cache key is now truly
   per-parent (`getParentIndex` ranks over the whole document, so the `c:dLbls`
   of different `c:ser` no longer collide).
4. ✅ **Empty shells in `XmlElements`**: `dataPoint()` → `c:idx` + `c:bubble3D`
   only; `dataPointLabels()` → empty `c:dLbls`; `dataPointLabel()` → `c:idx`
   plus an unopinionated `c:txPr` scaffold. All three insert via
   `insertInSchemaOrder` (`C_SER_CHILD_ORDER` / `C_DLBLS_CHILD_ORDER`), which
   also fixed the `c:dPt`-after-`c:dLbls` class of the schema-violation track —
   both allowlist entries removed. The hardcoded `dLbl.ts` blob is deleted.
5. ✅ **Document the contract** on the `Modification` type: what `index` indexes,
   the one-grow-per-call limit, what `isRequired` guarantees, `matchIdx`.

**Guards** (Phase 5, in place):

- 🧪 ✅ tier-0: point styles on **non-contiguous** categories (0 and 3 of 15) →
  exactly two `c:dPt`, `c:idx` 0 and 3, two distinct colors
  (`modify-chart-point-styles-sparse.test.ts`).
- 🧪 ✅ tier-0: a series `style.label` on a template whose series carry **no**
  `c:dLbls` → no `c:dLbls` created (same suite + styled test). Covers B.
- 🧪 ✅ tier-1 (stronger than tier-0 — checked on *every* written archive): the
  `c16:uniqueId` fingerprint `{00000001-04B4-49A4-AD60-4DBFE8A0F479}` is an
  `errors`-class invariant in `pptx-invariants.ts`; no shipped template
  contains it (verified 2026-08-12), so any occurrence is fabricated XML.
- 🧪 ✅ `modify-existing-chart-styled.test.ts` asserts the exact `c:dPt`/`c:dLbl`
  sets per series via `expectXml` — its data was already the sparse case, and
  the showcase output is correct again (`333333@0` / `efefef@1` / `eecc00@1` +
  `eeccff@3` + one styled label at idx 3).
- 🧪 ✅ tier-3 golden deck `chart-radar-labels`: 5-series pptxgenjs-generated
  radar chart with styled series labels re-loaded through `setChartData` — the
  baseline pins "no labels appear". (One upstream pptxgenjs allowlist entry
  added: `c:invertIfNegative` in radar series cascades into a `c:axId` report.)

**Downstream note.** A downstream api service consumes these paths through
`ShapeModifiersChartService` → `modify.setChartData`; the api side was verified
correct in both cases (metas, style matching and the `categories[].styles` array
all carry the right values), so no api change is needed — but regenerating
affected customer decks is part of "done".

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
2. 🐛 **Chart modifiers insert children out of schema order**: ~~`c:dPt` after
   `c:dLbls` in `c:ser`~~ (✅ fixed 2026-08-12 with the Modification-contract
   track — `c:dPt`/`c:dLbls` creation goes through `insertInSchemaOrder` now,
   both allowlist entries removed), `c:tx` misplaced inside `c:dLbl`
   (`modify-chart-datalabels-text`), `c:dLbls` misplaced in scatter `c:ser`
   and `a:solidFill` misplaced in `c:dLbl`/`c:spPr`
   (`modify-chart-datalabels`). The shared `XmlHelper.insertInSchemaOrder`
   primitive from the HTML→text track is the intended fix vehicle for the
   rest.
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

## Performance track — unbounded memory growth on large decks — ✅ done 2026-08-14

Investigated 2026-08-12 (v0.8.2), deferred, then **implemented 2026-08-14**
(fixes 1–4 below; only the optional item 5 remains open — see the outcome
block at the end of this track). Reported
symptom: on big decks the run gets slower and slower, with
`Cleaning unsupported tag p:custDataLst` as the last log line before each stall,
eventually dying.

**That log line is a red herring.** It comes from
`PlaceholderNormalizer.removeUnsupportedTags` (`placeholder-normalizer.ts:126`),
which is the *last* step of `Slide.append()` — so the growing gap between two
consecutive lines is the whole per-slide pipeline slowing down, not the cleanup.
Measured, the cleanup is 3–4 % of runtime.

**Reproduction.** Synthetic Think-cell-like template: 600 shapes/slide, each
with a `p:custDataLst`, ~310 KB slide XML, appended N times onto `RootTemplate`.

| slides | heap | RSS | ms/slide |
|--------|------|-----|----------|
| 20     | 212 MB | 332 MB | 67 |
| 60     | 534 MB | 675 MB | 56 |
| 120    | 979 MB | 1159 MB | 57 |

Growth is linear and permanent: **~8 MB of heap per slide, never released.**
With `--max-old-space-size=900` and N=400, per-slide time climbs 15 ms → 29 ms
as the limit nears, then `FATAL ERROR: Reached heap limit` at ~slide 365. The
"getting slower" is GC thrash on a heap that only ever grows.

### Root cause

1. ⚡🏗 **`Archive.buffer` never evicts** (`src/helper/archive/archive.ts:14`).
   `readXml()` (`archive-jszip.ts:151`, `archive-fs.ts:187`) parses a part into
   an xmldom `XmlDocument` and pushes it into `buffer`; `writeXml()` only
   re-buffers the same object. Serialization happens exclusively in
   `writeBuffer()` at `output()`/`stream()` time. By the end of a run **every
   slide, rels, layout and master DOM of the whole output deck is live at
   once** — and an xmldom DOM costs roughly **25× its XML source size**
   (310 KB slide → ~8 MB heap).

   Verified by experiment: serializing each finished slide back into the zip and
   dropping it from `buffer` after each append, same 120 slides →
   buffered entries 243 → 3, final heap **1019 MB → 67 MB** with no upward
   trend, total time 7.7 s → 6.3 s (less GC pressure more than pays for the
   extra serialization).

### Secondary findings (real, none dominant today)

2. ⚡ **`fromBuffer` is a linear scan** (`archive.ts:64`) run on every
   `readXml`/`writeXml` — ~20 calls per slide, so O(n²) in parts: 643 k string
   comparisons at 300 slides. Wants a `Map<string, ArchivedFile>`.
3. ⚡ **XML parsing dominates per-slide CPU** (~48 %, attributed to
   `parsePlaceholders` only because that is where the target slide is first
   read). Inside it, `@xmldom/xmldom` 0.9.10 compiles a **fresh RegExp for every
   closing tag** (`lib/sax.js:151`,
   `g.reg('^', g.QName_group, g.S_OPT, '$')`) — 19 % of total runtime in the CPU
   profile. Memoizing `grammar.reg` gave a reproducible **~20 % end-to-end
   speedup** (3.55 s → 2.99 s over 60 slides). Upstream bug; 0.9.10 is latest.
4. ⚡ **`addToPresentation` is quadratic in slide count** — 12 / 66 / 196 ms for
   100 / 400 / 800 slides, from `getMaxId` + append over the ever-growing
   `presentation.xml.rels` and `[Content_Types].xml`. Negligible below ~1000
   slides.
5. ⚡ **`removeUnsupportedTags` is superlinear per slide**:
   `XmlHelper.sliceCollection` (`xml-helper.ts:625`) reads `collection.length`
   every iteration, and that is a live `LiveNodeList` getter which **re-walks
   the entire document** after each removal (same pattern in
   `modifyCollection`, `xml-helper.ts:806`). Plus `parents.includes()` is
   O(k²). Measured 0.6 ms/slide at 150 tags, 3.5 ms at 600 — i.e. ~3.5 % of
   runtime, *not* the bottleneck despite being the code behind the log line.

### Constraint any fix must respect

`Archive.toBuffer` (`archive.ts:44`) **silently ignores the write when the path
is already buffered** — callers rely on mutating the same document object in
place, and `writeXml()` can therefore never replace a buffered document. An
eviction scheme has to fix that too. Re-reading an evicted part re-parses it
from the zip, which is correct as long as the flush wrote the serialized
content back first (verified: output file stays valid).

### Proposed fix, in order

1. ⚡🏗 ✅ **Flush + evict a part's DOM once it is finished** — for slides after
   `cleanSlide()`, for masters/layouts after their append. Turns O(deck) memory
   into O(1) and is by far the biggest win. Needs `toBuffer` to actually
   replace (see constraint above).
2. ⚡ ✅ **`buffer` → `Map`**, killing the linear `fromBuffer` scan.
3. ⚡ ✅ **Memoize `grammar.reg`** in a small local patch and report upstream to
   `@xmldom/xmldom`. ~20 % for a few lines.
4. ⚡ ✅ **Snapshot live NodeLists into plain arrays** in `sliceCollection` /
   `modifyCollection` before mutating; `parents` → `Set`.
5. ⚡ ✅ Optional, later: `addToPresentation` should cache the max rId / content-type
   state instead of rescanning the growing parts (item 4 above).

**Guard.** ✅ New Tier-1-style test: append N large slides and assert
`process.memoryUsage().heapUsed` stays under a ceiling (and that
`archive.buffer.size` stays bounded) — the current behavior fails it by an order
of magnitude, so it is a genuine regression gate once fixed.

### Outcome (2026-08-14)

- **Fix 1**: `Slide`/`Master`/`Layout.append()` end with
  `flushTargetXml()` — the new `IArchive.flushXml(file)` serializes the
  buffered DOM back into the archive and drops it; a slide also flushes its
  rels and copied notesSlide (same target number). `toBuffer` now
  **replaces** a buffered entry (last write wins), which resolves the
  constraint above; anything re-reading a flushed part (e.g. `cleanup` at
  write time) re-parses it from the serialized content.
- **Fix 3** is a guarded runtime patch (`src/helper/xmldom-sax-patch.ts`,
  applied on archive-module load): memoizes `grammar.reg` on the exports
  object `sax.js` calls through — safe because `reg` is pure and emits no
  `g` flag. No-op if a future xmldom version blocks the deep import.
  ⏳ Reporting it upstream to `@xmldom/xmldom` is still open — issue text is
  drafted (with a self-contained repro: 50k-element parse, 461 → 352
  ms/parse from hoisting the one closing-tag RegExp; verified no existing
  upstream issue as of 2026-08-18), needs a logged-in GitHub account to
  file.
- **Fix 4** also corrected the iteration semantics live lists silently had:
  a callback removing an element made the loop skip the next one
  (`normalizePlaceholderShapes`); snapshots process every element
  (CHANGELOG'd, suite unchanged).
- Measured (same synthetic 600-shape deck): 60 slides **566 → 54 MB** heap,
  124 → 4 buffered DOMs, 57 → 46 ms/slide; 120 slides **979 → 69 MB**, no
  upward trend; the N=400 / `--max-old-space-size=900` run that died at
  ~slide 365 finishes in 20 s at 149 MB. Remaining growth is the serialized
  XML strings inside the output zip (≈ source size, not 25×) — inherent
  until an fs-backed output is used.
- **Guard**: `__tests__/memory-bounded-large-deck.test.ts` builds the
  synthetic deck in a child process (`--expose-gc`, deterministic heap
  numbers) and asserts ≤ 10 buffered parts (pre-fix: `2·slides + 4`) and
  < 200 MB heap at 60 slides (pre-fix: ~566 MB).

### Outcome, fix 5 (2026-08-18)

- **Fix 5**: `addToPresentation` no longer rescans the growing parts per
  slide. `ContentTypeRegistry` keeps two WeakMap caches keyed by the parsed
  document (so an evicted/re-parsed part rebuilds them with one scan): the
  max rId of `presentation.xml.rels` (safe upper bound — the registry is the
  only appender to that part; `Template.removeSlideRelations` only removes),
  and the `p:sldIdLst`/`p:sldMasterIdLst` elements of `presentation.xml`
  (xmldom's live `getElementsByTagName` re-walks the whole document per
  access). `createContentTypeChild`/`createRelationshipChild` and
  `appendToSlideRel` now use `xml.documentElement` directly — `Types` and
  `Relationships` *are* the document elements, so the per-append live-list
  walk of `[Content_Types].xml` and every rels part was pure waste.
- Measured (empty slide appended N times, total `write()` time): 3200
  slides **4.30 s → 2.17 s**; ms/slide 800→3200 was climbing 0.96 → 1.35,
  now flat 0.82 → 0.68. Existing `remove-existing-slides` suite covers the
  truncate-then-append rId interplay; suite green (311 tests).

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
Urgent:   "Modification contract violations" bug track — both regressions are
          released in 0.8.2 and corrupt customer decks silently. Order:
          isRequired → diagnostic → setPointStyles addressing → empty shells,
          each with its tier-0 guard written first
Early:    Phase 6.1 (docs-example compile test) alongside Phase 5 tier 1 — it is
          the same kind of cheap always-on gate, and it must exist before the
          README split moves 126 examples around
Then:     Phase 6.2-6.5 docs split → site on GitHub Pages (only host; the
          self-hosted mirror was dropped 2026-08-14)
Done:     Performance track (archive buffer eviction) — postponed 2026-08-12,
          implemented 2026-08-14; addToPresentation caching added 2026-08-18;
          open: upstream report of the xmldom RegExp recompilation
          (issue text drafted, needs a logged-in GitHub account to file)
```

Rule of thumb for every PR during the refactor: **no public `modify.*` signature
changes**, integration suite stays green, and any behavior change (especially
error behavior) gets a CHANGELOG note.
