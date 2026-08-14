---
title: Testing and Validation Tools
description: The test suite, the Docker-based OOXML schema validator (validate:pptx), and visual regression via pptx-thumbnailer.
---

The repository ships with a layered test system that goes beyond unit tests: every archive written by a test is invariant-checked, all templates and outputs can be validated against the OOXML schema, and curated decks are rendered and pixel-diffed against committed baselines. These tools live in the [repository](https://github.com/singerla/pptx-automizer) (`tools/`), not in the npm package — clone the repo to use them (see [Getting started](./getting-started.md#as-a-cloned-repository)).

## Unit and integration tests

```
yarn test
yarn test-coverage
```

Tests are integration-style: they build a real presentation from `__tests__/pptx-templates/`, write it to `__tests__/pptx-output/`, and assert on the result — including XML-level assertions via the `expectXml` helper (`__tests__/helpers/expect-xml.ts`).

Two gates run automatically inside `yarn test`:

- **Archive invariants**: every archive any test writes is validated on the fly — referenced relationships resolve, all parts are covered by `[Content_Types].xml`, the slide list is intact, all XML is well-formed. If this fails your test, the change produced a broken deck.
- **Documented examples compile**: every fenced TypeScript block in this documentation, the README and AI-INSTRUCTOR.md is typechecked (`__tests__/docs-examples.test.ts`), so the docs cannot drift from the actual API.

## OOXML schema validation (`validate:pptx`)

```
yarn validate:pptx
```

Validates **all template and output .pptx files** with the [Open XML SDK](https://github.com/dotnet/Open-XML-SDK), packaged in a Docker image (`tools/validate-pptx/`) so you need Docker only — no .NET on your machine. CI runs it as the `validate-pptx` job; new schema errors fail the gate.

`tools/validate-pptx/allowlist.json` holds baseline template noise plus documented library bugs. Never add an allowlist entry to silence an error your own change introduced — removing an entry belongs to the fix for it.

This validator is also a great debugging tool when PowerPoint offers to *repair* an output file: it usually names the exact part and element that violates the schema. See [Troubleshooting](./troubleshooting.md).

## Visual regression (`test:visual`)

```
yarn test:visual
```

Renders curated golden decks (`__tests__/visual/*.deck.ts`) to PNGs through a pinned [pptx-thumbnailer](https://www.npmjs.com/package/pptx-thumbnailer) container (`tools/render-pptx/`, Docker only) and perceptually diffs each slide against the committed baselines in `__tests__/visual-baselines/`. CI runs it as the `visual-regression` job and uploads actual+diff PNGs on failure.

Good to know:

- It is a **change detector, not a correctness oracle** — the renderer is LibreOffice-based, and LibreOffice fidelity is not PowerPoint fidelity. Never conclude PowerPoint-correctness from green pixels alone.
- On an *intended* visual change, regenerate the baselines in the same PR so the reviewer sees the before/after images:

  ```
  UPDATE_BASELINES=1 yarn test:visual
  ```

- The suite is deliberately small: a handful of curated decks of 1–5 slides, using fonts shipped in the renderer image. Changing anything in `tools/render-pptx/Dockerfile` (base digest, thumbnailer version, fonts) invalidates all baselines — regenerate them in the same PR.

## Contributing

New features need a test with a template .pptx — prefer reusing the existing templates in `__tests__/pptx-templates/` over adding new binaries. If you are contributing with the help of an AI agent, point it at [AGENTS.md](https://github.com/singerla/pptx-automizer/blob/main/AGENTS.md) in the repo root, which states the full rule set for the test tiers.
