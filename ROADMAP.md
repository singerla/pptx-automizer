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

## Phase 1 — Error handling & logging (foundation for everything else)

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

## Phase 2 — Decompose `HasShapes` and centralize OOXML paths

`src/classes/has-shapes.ts` is a 1300-line base class mixing at least six
concerns. It's where most future features will land, so pay this debt first.

- 🏗 Extract collaborators (composition over inheritance):
  - `ElementImporter` — queueing + `getElementInfo`/`findElementOnSlide`/dispatch
    (the `importedSelectedElements` switch)
  - `RelatedContentCopier` — `copyRelatedContent` (charts/images/diagrams/OLE/hyperlinks)
  - `SlideNotesCopier` — the notesSlide number remapping trio
  - `PlaceholderNormalizer` — `removeDuplicatePlaceholders`/`normalizePlaceholderShapes`/`cleanSlide`
  - `ContentTypeRegistry` — the `appendToContentType`/`appendToSlideList` family
- 🏗 **Central `PptPaths` helper**: `ppt/slides/slide${n}.xml` and friends are
  string-built in ~20 places (has-shapes, slide, shape, file-helper, ole, …).
  One typo-prone convention, zero reuse. A tiny module with
  `slide(n)`, `slideRels(n)`, `master(n)`, `layout(n)`, `chart(n)`, `notesSlide(n)`
  removes a whole bug class and makes the `fs`-archive mode auditable.
- 🏗 Replace the `analyzeElement` if-chain with a **shape-type detector registry**
  (`[{ match: (el) => el.getElementsByTagName('c:chart').length, type: Chart, relType: 'chart' }, …]`).
  New shape types (video/audio are on the wish list) become one registry entry
  instead of another branch in a 110-line method.
- 🏗 The dynamic dispatch `new Chart(info, …)[info.mode](…)` (calling `append`/
  `modify`/`remove` by string) defeats type checking. Give shapes an explicit
  interface: `IShapeAction { append(...); modify(...); remove(...) }` and call
  methods directly.

## Phase 3 — Kill global state (enables concurrent instances)

- 🏗 `contentTracker` is a module-level singleton (`content-tracker.ts:295`,
  reset via `Tracker.reset()` in `automizer.ts:finalizePresentation`). Two
  `Automizer` instances running concurrently — the obvious server use case, and
  Ensemblio's — share and corrupt each other's tracking state. The code already
  has two TODOs for this. Move the instance onto the root `Template` (it already
  exists as `automizer.content`; the remaining offenders are the direct
  `contentTracker` imports in `xml-helper.ts` and `file-helper.ts` — thread it
  through instead).
- 🏗 Same for `Logger.verbosity` (process-global).
- 🧪 Add a regression test: two Automizer instances built in parallel
  (`Promise.all([presA.write(...), presB.write(...)])`) produce valid output.

## Phase 4 — Template & archive layer clarity

- 🏗 `Template` plays two roles decided by whether `params.name` is set
  (`Template.import`), then casts to `PresTemplate` or `RootPresTemplate`.
  `isPresTemplate` tests `'name' in template`. Split into two classes
  (`SourceTemplate`, `OutputTemplate` extends shared base) and delete the casts —
  half the fields on `Template` are `undefined` in one of the roles today.
- 🏗 `template.archive` is typed `IArchive` but is `await`ed all over
  (`await this.archive`) as if it were a promise — it isn't; initialization is
  lazy inside `ArchiveJszip.read`. Make initialization explicit
  (`async open()`), type honestly, drop the fake awaits.
- 🏗 `Template.file: any`, `zipCopyWithRelations(parentClass, …)` untyped
  "parentClass" params, 13 `: any` — tighten while touching these files.
- 🔧 Turn on `strict` incrementally: start with `noImplicitAny` +
  `strictBindCallApply`, then `strictNullChecks` (the big one — the pipeline has
  many "returns undefined on failure" paths that Phase 1 converts to throws,
  which makes strictNullChecks feasible).
- 🔧 Modernize the build: `"lib": ["es2020","dom"]` pulls browser DOM types into
  a Node library and shadows xmldom types (e.g. `XMLDocument` in `slide.ts` is the
  browser type). Remove `dom`, use xmldom's types consistently. Consider dual
  CJS+ESM output (tsup or tsc twice + `exports` map) — CJS-only is increasingly
  painful for consumers.
- 🏗 Consider making `pptxgenjs` an **optional peer dependency**, lazily imported
  by the generator bridge. It's a heavy dependency that pure "modify existing
  pptx" users never need. (`runExternalGenerator` currently instantiates the
  bridge unconditionally on every write.)

## Phase 5 — Testing strategy

The 94-suite integration harness is a real asset (it exercises actual OOXML end
to end). Its weakness: assertions are almost all `expect(result.slides).toBe(n)` —
the content of the produced XML is unverified, so regressions that corrupt output
without crashing pass green.

- 🧪 Add an **output-assertion helper**: open the written .pptx with jszip, fetch
  a part, and assert on XML (`expectXml(output, 'ppt/slides/slide1.xml').toContainElement('a:t', 'my text')`).
  Retro-fit onto the highest-value suites first (charts, tables, text).
- 🧪 Add unit tests for pure helpers (`modify-text-helper`, `modify-color-helper`,
  `cell-id-helper`, `general-helper`) — they're pure functions on DOM nodes,
  cheap to test, currently only covered incidentally.
- 🧪 Optional: a validity smoke check (open output with jszip and verify all
  `[Content_Types].xml` overrides and rel targets resolve to existing parts) —
  a cheap proxy for "PowerPoint can open this", catching the most common
  corruption class (dangling relationships).
- 🔧 Wire coverage into CI once the glob is fixed (Phase 0).

## Phase 6 — Docs & AI enablement

- 📖 README is a 1500-line monolith. Keep a short README (pitch, install, basic
  example, links) and split the rest into `docs/` (selectors, text, tables,
  charts, images, masters/layouts, generation, output modes, troubleshooting).
- 📖 Publish typedoc (dev dep already present, no script) via GitHub Pages.
- 📖 **AI instructor** (done in this pass): `AI-INSTRUCTOR.md` — a self-contained
  guide users can hand to Claude/other assistants to generate working
  pptx-automizer code. Maintenance rule: update it in the same PR as any
  `modify.*` API change. Ship it in the npm package (`files` array) and consider
  publishing it as `llms.txt` on the docs site / repo root so AI tools find it
  automatically.
- 📖 Document the deferred-execution model explicitly (biggest user surprise:
  callbacks run at `write()`, not at `addSlide()`).

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
8. HTML → text conversion (`modify.htmlToMultiText`) is incomplete and partly
   incorrect — see the dedicated feature track below.

---

## Feature track — HTML → PPTX text (`htmlToMultiText`)

Audit date: 2026-08-11. Scope: `src/helper/html-to-multitext-helper.ts`
(HTML → `MultiTextParagraph[]`) and `src/helper/multitext-helper.ts`
(→ DrawingML). Independent of the refactor phases; can proceed in parallel.

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

## Suggested sequencing

```
Week 1:   Phase 0 (all) + CI green on main
Weeks 2-3: Phase 1 (errors/logging) — small PRs, mechanical
Weeks 3-6: Phase 2 (HasShapes decomposition + PptPaths) — one extraction per PR,
           integration suite as the safety net
Then:     Phase 3 (globals) → Phase 4 (templates/strict) as background chores
Ongoing:  Phase 5 assertions added with every bug fix; Phase 6 docs split
```

Rule of thumb for every PR during the refactor: **no public `modify.*` signature
changes**, integration suite stays green, and any behavior change (especially
error behavior) gets a CHANGELOG note.
