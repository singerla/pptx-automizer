# Changelog

All notable changes to this project are documented here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/); versions follow
semver, with the pre-1.0 convention that **breaking changes bump the minor**.

## [0.9.3] — 2026-08-22

Completes the certified 0.9.2 release candidate: the certification sweep that
cleared 0.9.2 ran with both `fix/hyperlink-rel-type-guards` **and**
`fix/master-import-ole-and-image-rels` applied, but only the former was merged
before tagging. This release adds the missing branch; there are no other
changes.

### Fixed

- Copying a slide, layout, or master no longer leaves orphaned image
  relationships behind. `Image.modifyOnAddedSlide()` appended a fresh
  relationship for every image on the copied part and left the source-named
  relationship in the `.rels` file; whenever that orphan's Target did not
  exist in the output archive (same basename, different extension in root vs
  source), PowerPoint prompted to repair the file. The existing
  relationship's Target is now pointed at the copied media file in place —
  the rId stays stable, the element XML no longer needs rewriting, and no
  orphaned relationships remain.
- `autoImportSlideMasters` no longer fails with "Could not find file
  ppt/slides/slide<n>.xml" when an imported slideLayout or slideMaster
  carries an OLE object (e.g. think-cell data objects).
  `OLEObject.updateSlideXml()` built a hardcoded slide path from the target
  number, which is a layout/master number on that path; it now uses the
  target-type-aware slide file the Shape base class already derives. Covered
  by a regression test with a schema-faithful OLE fixture on a slideLayout.

## [0.9.2] — 2026-08-22

### Fixed

- Modifying a template shape that already carries a hyperlink
  (`slide.modifyElement()` on a shape classified as `ElementType.Hyperlink`)
  could corrupt the output file: PowerPoint reported "found a problem with
  content" and offered to repair, silently dropping the shape.
  `Shape.modifySlideTree()` unconditionally rewrote the shape's
  `<a:hlinkClick r:id>` to a separately precomputed id *after*
  `editTargetHyperlinkRel()` had already verified the existing relationship
  was correct — leaving an `r:id` with no matching `<Relationship>` entry in
  the slide's `.rels`. The rewrite is removed from the modify path;
  `Hyperlink.modify()` owns its relationship via `editTargetHyperlinkRel()`.
  On the append path (`slide.addElement()`), the rewrite to the freshly
  reserved id stays — it is what protects a cloned shape from colliding with
  an unrelated relationship already declared under its source rId on the
  target slide — and `ModifyHyperlinkHelper.addHyperlink()` now creates the
  matching relationship explicitly: it checks whether an existing
  `<a:hlinkClick>`'s `r:id` resolves to a relationship *of hyperlink or
  slide Type* before skipping (the genuine pptxgenjs-authored case). If the
  id is unmatched it creates the relationship for it; if it resolves to a
  relationship of the wrong Type it allocates a fresh id instead of silently
  pointing the hyperlink at an unrelated part. Contributed in part by
  [#205](https://github.com/singerla/pptx-automizer/pull/205).
- Importing a shape whose `a:hlinkClick r:id` is stale — pointing at a
  *structural* relationship of the source slide (classically `rId1` =
  slideLayout, `rId2` = notesSlide) — corrupted the package:
  `HyperlinkProcessor.copyMultipleHyperlinks()` (the GenericShape append
  path for multi-hyperlink shapes, and the fallback for hyperlink shapes
  whose rel lookup fails) cloned whatever relationship sat at that id, Type
  and all, giving the target slide a second slideLayout/notesSlide
  relationship with a source-numbered target — an OPC singleton violation
  ("can only have one instance of relationship that targets part") observed
  in field-generated decks. Only relationships of hyperlink/slide Type are
  copied now, and the two hyperlink id spaces are kept apart: ids present on
  the unmutated source element are imports and resolve against the source
  slide's rels, while ids added by modification callbacks during `prepare()`
  (e.g. `modify.htmlToMultiText()` links on an added element) already have
  their relationships on the target slide's rels and are left untouched.
  Hyperlinks that resolve in neither id space are stripped from the imported
  element with a warning, since a dangling `r:id` is itself a repair
  trigger. Downstream, the pre-fix behavior surfaced either as the repair
  prompt or as **silent link loss** — link text rendered, hyperlink gone or
  pointing at an unrelated part — depending on what sat at the colliding id,
  which made this fix release-blocking for correctness rather than an edge
  hardening. A certification sweep over field-generated decks that previously
  failed to open under the Open XML SDK opens cleanly with these guards in
  place. The Tier-1 package invariants (checked for every
  archive written by the test suite) now also fail on duplicate
  slideLayout/notesSlide relationships and on any
  `a:hlinkClick`/`a:hlinkHover` resolving to a structural relationship.

## [0.9.1] — 2026-08-20

Documentation, performance and dependency-security housekeeping on top of
0.9.0, plus chart and text-rendering fixes. No API breaking changes; one
visible behavior change: `htmlToMultiText` output now carries the vertical
spacing the HTML implies, so decks regenerated from HTML render with the
paragraph gaps they previously lacked (see **Fixed**).

### Added

- Documentation site at <https://singerla.github.io/pptx-automizer> (ROADMAP
  Phases 6.2/6.3): the README's feature documentation, split into per-feature
  pages under `docs/`, published with a generated typedoc API reference and
  full-text search; the README is reduced to pitch, install and one example.
- `MultiTextParagraph.paragraph.lineSpacing`/`spaceBefore`/`spaceAfter` accept
  `{ percent }` in addition to points: rendered as `<a:spcPct>`, spacing that
  scales with the paragraph's line height instead of a fixed point value.
- `ModifyPresentationHelper`, `XmlSlideHelper`, `XmlRelationshipHelper` and the
  `FindElementSelector` type are exported from the package root. The README
  documented imports for them (`pptx-automizer/src/...`, `./helper/...`,
  `./types/types`) that no npm consumer could actually resolve — only `dist/`
  ships.
- Every fenced ```ts example in `README.md` and `AI-INSTRUCTOR.md` is compiled
  against the current source on every `yarn test`
  (`__tests__/docs-examples.test.ts`, ROADMAP Phase 6.1) — documented API drift
  now fails CI.
- `yarn docs:api` generates the typedoc API reference into `website/docs/api`
  (gitignored).
- The docs site ships AI renderings, generated at build time (ROADMAP 6.4):
  `llms.txt` (per-page index with one-line descriptions), `llms-full.txt`
  (the guide corpus in one file), and a Markdown twin of every HTML page —
  guides and API reference — served at the page's URL plus `.md`.
- `AI-INSTRUCTOR.md` (shipped in the npm package) is now **generated from the
  documentation corpus** (`yarn docs:ai`; hand-written parts: mental model,
  rules list, minimal example) and covers every feature page instead of a
  condensed cheat sheet. A test fails `yarn test` whenever the committed file
  drifts from the docs it is built from.

### Changed

- **Memory no longer grows with deck size** (ROADMAP performance track).
  Every appended slide, master and layout (plus its rels and notesSlide) is
  serialized back into the archive and its parsed DOM dropped from the
  buffer as the last step of its append — previously every part's xmldom
  document (~25× the XML source size) stayed live until `write()`, which on
  large decks meant GC thrash and eventually `FATAL ERROR: Reached heap
  limit`. Measured on a synthetic 600-shapes-per-slide deck: 60 slides
  566 MB → 54 MB heap; a 400-slide run under `--max-old-space-size=900`
  used to die at ~slide 365 and now finishes with 149 MB. Guarded by
  `__tests__/memory-bounded-large-deck.test.ts`. Two internal surfaces
  moved with this: `IArchive` gained `flushXml(file)`, and the archive
  `buffer` is a `Map<string, ArchivedFile>` (was an array with a linear
  scan per `readXml`/`writeXml` — O(n²) in part count). Writing XML for an
  already-buffered path now *replaces* the buffered document (last write
  wins); before, the first buffered DOM silently won.
- `@xmldom/xmldom`'s SAX parser compiles a fresh RegExp for every closing
  tag it parses (~19 % of total runtime on parse-heavy runs, upstream issue
  in 0.9.10). A guarded runtime patch memoizes the grammar's RegExp
  compilation — ~20 % end-to-end speedup on template-heavy workloads; a
  no-op if a future xmldom version blocks or fixes it.
- Registering an appended slide/master/layout with the presentation is no
  longer quadratic in deck size (ROADMAP performance track, fix 5).
  `ContentTypeRegistry` caches the max rId of `presentation.xml.rels` and
  the `p:sldIdLst`/`p:sldMasterIdLst` elements per parsed document instead
  of rescanning the growing parts on every append, and appends to
  `[Content_Types].xml` and rels parts address the container via
  `documentElement` instead of a live `getElementsByTagName` walk of the
  whole part. Appending 3200 empty slides: 4.3 s → 2.2 s, with per-slide
  time now flat instead of climbing.
- `XmlHelper.sliceCollection`/`modifyCollection` snapshot xmldom's live
  node lists before iterating. Besides removing superlinear re-walks of the
  document (visible as ever-slower `Cleaning unsupported tag` steps on
  large slides), this fixes the iteration semantics when a callback removes
  elements: placeholder normalization previously skipped the element
  following each removed one.

### Fixed

- `htmlToMultiText` packed every paragraph flush against the previous one: the
  browser's default vertical margins (`<p>`, headings, list edges) were lost,
  and a trailing `<br/>` inside a block — invisible in HTML — became a real
  empty line. The converter now emits one collapsed gap
  (`spaceBefore: { percent: 100 }`) at every block boundary — between block
  paragraphs, at list edges including two adjacent lists, never between items
  of the same list — and drops exactly one trailing `<br/>` per block (plus
  the collapsed whitespace around it); a deliberate `<br/><br/>` keeps its
  empty line. Decks regenerated from unchanged HTML will render with these
  gaps where they previously had none.
- Chart per-point styling (`setChartData` with `categories[].styles`) no
  longer fabricates default-styled `<c:dPt>` elements. Previously every
  styled-looking point — including pseudo-styles such as `{ marker: {} }` or a
  color without `type` — produced a data point carrying a default grey
  `<a:solidFill>` plus `<a:ln><a:noFill/>`, which on **line charts** erased
  the line segments (the chart rendered as floating labels). Now a style that
  yields no applicable modification produces no `<c:dPt>` at all, and a
  created `c:spPr` starts empty — points inherit the series formatting except
  for exactly what the caller asked for. The same applies to a series
  `style.color` on a template series without `<c:spPr>`. Per-point border
  styles on template-less points now materialize as a bare `<a:ln>` in schema
  position instead of silently relying on the fabricated default (per-point
  *marker* styling remains modify-if-present).
- A created `<c:dPt>` now pins `<c:invertIfNegative val="0"/>`. OOXML
  defaults the absent element to *true*, so PowerPoint inverted the fill of
  negative-value **bar** points — rendering them white with a border and
  silently overriding the very fill the point was styled for. The pre-0.9.0
  code never hit this because it cloned the template's existing data point
  (inheriting its `invertIfNegative`); the empty-shell creation lost the flag.
  LibreOffice ignores `invertIfNegative`, which is why the tier-3 golden
  decks could not catch it.
- Documentation only — examples that documented APIs which never existed:
  `replaceText` has no regex/whole-word/case options (its options are the
  `openingTag`/`closingTag` tag delimiters, default `{{`/`}}`);
  `ModifyShapeHelper.setSolidFill` takes no color argument (it sets theme color
  accent6); `ModifyCleanupHelper.clearTextColor` must be wrapped in a callback
  (passed bare, the relationship element lands in its `color` parameter);
  table `expand.mode` is `'column'`, not `'col'`; `dLblPos` takes the
  `LabelPosition` enum. Added a "Deferred execution" section to the README.
- Documentation only — two claims that 0.9.0 had made stale: callback
  exceptions are no longer "swallowed" (a throwing callback or unresolvable
  selector rejects `write()` with a typed error unless
  `continueOnError: true`), and concurrent Automizer instances in one process
  are supported (the global state was removed in 0.9.0, with a regression
  test). `AI-INSTRUCTOR.md` and `docs/concepts.md` said otherwise.
- fs mode: `compressFolder` returned before the output stream finished and
  swallowed write errors, so callers could clean up the work directory while
  the zip was still being written. It now awaits stream completion and
  rethrows failures. (#202)

### Security

For the full review record of the changes below — findings, threat-model
discussion and open questions — see
[`reviews/pr-202-security-deps.md`](./reviews/pr-202-security-deps.md).

- Removed the `extract-zip` dependency (unmaintained since 2020;
  CVE-2026-19693: arbitrary file write via crafted zip symlink entries, no
  patched release). fs mode extracts via an internal jszip-based
  `extractToFolder` that rejects absolute and zip-slip entry paths and skips
  symlink entries. Note: the whole archive is held in memory while
  extracting, unlike the streaming extractor it replaces. (#202)
- Removed the `image-size` dependency (repo archived upstream;
  CVE-2025-71329/71330: infinite-loop DoS in its ICNS/JXL/HEIF parsers, no
  patched release) in favor of an internal loop-safe parser for PNG, JPEG,
  GIF, BMP, WebP and SVG. TIFF and ICO dimensions are no longer detected;
  cover cropping (`setRelationTargetCover`) falls back to default dimensions
  for such media with a warning. `pptxgenjs` still pins a vulnerable
  `image-size` transitively; the yarn resolution to the newest release
  remains in place as mitigation. (#202)

## [0.9.0] — 2026-08-12

The architecture audit of `ROADMAP.md` (Phases 0–5). Consumers should read
**Breaking changes** before upgrading: the `modify.*` call signatures are
unchanged, but the *types* behind them and the default error behavior are not.

### Breaking changes

- **The xml layer is typed with `@xmldom/xmldom`, not the TypeScript `dom`
  lib.** `tsconfig.json` no longer includes `"dom"` in `lib`, and the exported
  `XmlDocument` / `XmlElement` / `XmlElementCollection` now resolve to
  xmldom's `Document` / `Element` / `LiveNodeList<Element>` instead of the
  browser globals. This is what the library always parsed at runtime — only
  the types were wrong.

  Any consumer that types its callbacks against the dom globals will fail to
  compile with errors of the shape *"Type 'Element' is missing the following
  properties from type 'Element': getQualifiedName, isSupported"*. Fix by
  taking the aliases from the library instead of redeclaring them:

  ```diff
  - type XmlDocument = XMLDocument;   // TypeScript dom lib
  - type XmlElement = Element;
  + import type { XmlDocument, XmlElement } from 'pptx-automizer';
  ```

  Casts through the dom `Node` (`... as unknown as Node`) need to go as well;
  the parsed element is already the node type xmldom's `replaceChild` expects.

- **A throwing modification callback or an unresolvable element selector now
  fails the run** with a typed error, instead of logging and continuing. Pass
  `continueOnError: true` in the `Automizer` options to restore the previous
  lenient behavior.

- **`verbosity` semantics changed.** `0` now means *errors only* (it used to
  mean no output at all); `1` (default) adds warnings, `2` adds info. To
  silence the library completely, inject the new `NullLogger`.

- **`PlaceholderInfo.id` is typed `string`**, not `number`. The value has
  always been the raw `ph` attribute string; only the declaration was wrong.

- **Modification callbacks may return `Promise<void>`.** Signatures that
  previously declared `=> void` now declare `=> void | Promise<void>`, and
  async callbacks are awaited (they were silently not awaited before, which
  made cover images replace against a stale element).

- **Media files are deduplicated in the output** (#145). Identical images are
  written once and shared; decks that relied on a particular media part name
  or count will see different `ppt/media/*` entries.

### Added

- `logger` option plus `ILogger`, `ConsoleLogger`, `NullLogger` exports — route
  library output into your own logging stack.
- Typed errors: `AutomizerError` (base), `TemplateNotFoundError`,
  `SlideNotFoundError`, `ElementNotFoundError`, `ArchiveError`, `OutputError`,
  `CallbackError`. A missing template now names the directories that were
  searched instead of failing as `File not found: undefined` deep inside
  `FileHelper`.
- `modify.setOutline` for shape outlines (width, dash style, color) — see #188.
- `PtToEmu` / `EmuToPt` unit helpers, alongside the existing `CmToDxa` /
  `DxaToCm`.
- `ChartInfo`, `WorkbookData` and `ShapeOutline` are exported types.
- `htmlToMultiText` understands inline CSS, and handles soft line breaks in
  `setMultiText` (#186).
- `Modification.matchIdx` — address `c:dPt` / `c:dLbl` by their `<c:idx>`
  payload rather than by sibling position. See the chart fixes below.
- Concurrent `Automizer` instances are supported: `ContentTracker` and the
  logger are per-instance / async-scoped instead of module globals.

### Fixed

- **Per-point chart styles on non-contiguous categories were silently dropped**
  (regression in `3d6c452`, released in 0.8.2). `setPointStyles` addressed the
  `c:dPt` / `c:dLbl` *element slot* by category index, but the index is
  positional over existing siblings and the assert can only grow a collection
  by one per call. Styling categories 0 and 3 produced two `c:dPt` with
  `c:idx val="0"` — a duplicate of the first styled point — and lost the second
  style. Sparse style sets now resolve by `<c:idx>` and are created in
  ascending idx order. Affected every caller styling individual data points;
  in one production deck, 879 of 1224 charts carried duplicated `c:dPt`.

- **A styled series label fabricated a `c:dLbls` in every series** (regression
  in `0a45454`, released in 0.8.2). `seriesDataLabel` has marked its `c:dLbls`
  `isRequired: false` since 2022, meaning *modify if present, never create* —
  but the flag was never honored, and once `createElement` learned to build
  `c:dLbls` it began injecting a hardcoded label blob (14 pt, `accent1`,
  `showVal=1`) into templates that deliberately carried no data labels.
  `isRequired: false` now does what it says, and the hardcoded `dLbl.ts`
  template is deleted.

- `isRequired: false` is honored throughout `ModifyXmlHelper.assertElement`:
  when the element is absent it returns without creating or cloning.
- Failed asserts are reported again (warn for required tags, debug for skipped
  optional ones) — the diagnostic had been commented out, which is why both
  chart regressions shipped unnoticed.
- `XmlElements` builds empty shells and inserts them in schema order: `c:dPt`
  is `c:idx` + `c:bubble3D`, `c:dLbls` is empty, `c:dLbl` is `c:idx` plus an
  unopinionated `txPr` scaffold. Nothing is styled that the caller did not ask
  for, and the `c:dPt`-after-`c:dLbls` schema violation is gone.
- Chart worksheets handle empty cells and no longer emit duplicate addresses
  (#39).
- `write()` reported its duration with a `/600` divisor — summaries were wrong
  by ~40%. It is seconds now.
- Generic shapes copy their image relationships when imported.
- Hyperlink handling is more robust and validates its targets.
- Async shape modification callbacks are awaited, so cover images replace
  correctly when the placeholder is larger.

### Changed

- **A `Color` without a `value` no longer aborts point styling.**
  `ModifyColorHelper.normalizeColorObject` used to throw a `TypeError`
  (`color.value.indexOf`), which ended `setPointStyles` after the first point.
  Phase 1's hardening removed the crash — so a malformed color that previously
  affected one point now reaches *all* of them. Check callers that were
  (knowingly or not) relying on that abort.
- `pptxgenjs` is loaded lazily and only when `generate()` is used.
- `HasShapes` is decomposed, OOXML paths are centralized in `PptPaths`, and
  shape dispatch is typed. `Template` is split into `SourceTemplate` and
  `OutputTemplate` — internal classes, not part of the public surface.
- `IArchive` opens idempotently and no longer declares fake `async` methods.

### Internal

- CI on Node 20/22 (typecheck, lint, test); ESLint 9 flat config;
  `noImplicitAny` and `strictBindCallApply` enabled.
- Four-tier test model (`ROADMAP.md` Phase 5): tier-0 XML assertions via
  `expectXml`, tier-1 package invariants on every written archive, tier-2
  Dockerized OOXML schema validation (`yarn validate:pptx`), tier-3 visual
  regression against golden decks via `pptx-thumbnailer`.
- Guards for the two chart regressions above: a sparse point-style assertion
  (categories 0 and 3 of 15), a no-fabrication assertion for label-less
  series, a tier-1 invariant on the deleted label fingerprint, and a
  `chart-radar-labels` golden deck. The showcase test
  `modify-existing-chart-styled` asserts XML instead of a chart count.

---

## [0.8.2] and earlier

Not covered by this file; see the git history and the release notes on GitHub.
