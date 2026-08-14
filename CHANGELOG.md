# Changelog

All notable changes to this project are documented here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/); versions follow
semver, with the pre-1.0 convention that **breaking changes bump the minor**.

## [Unreleased]

### Added

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

### Fixed

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
