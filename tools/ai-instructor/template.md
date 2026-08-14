<!-- GENERATED FILE — do not edit AI-INSTRUCTOR.md directly.
     Source: tools/ai-instructor/template.md + the docs/ pages it includes.
     Rebuild: yarn docs:ai   (drift is caught by __tests__/ai-instructor.test.ts) -->

# AI Instructor: Generating PowerPoint with pptx-automizer

> **For AI assistants (Claude, etc.):** This document teaches you how to write
> correct code with `pptx-automizer`. Follow it as authoritative instructions.
>
> **For humans:** Drop this file into your project (or paste it into your AI
> chat / reference it from your CLAUDE.md) and ask your assistant to build
> PowerPoint automation with it. It is generated from the
> [documentation](https://singerla.github.io/pptx-automizer/) — the hand-written
> parts are the mental model, the rules list and the minimal example.

## Mental model (read first)

`pptx-automizer` is a **template-based** .pptx engine for **Node.js (server-side
only)**. You never build slides from a blank canvas. Instead you:

1. Design slides in **PowerPoint** and save them as template .pptx files. Give
   every shape you want to automate a **unique, readable name** via PowerPoint's
   Selection pane (`ALT+F10`, rename with `F2`).
2. Load a **root template** (defines masters/theme/size of the output) and any
   number of **source templates**.
3. **Add slides** from source templates to the output, and **modify shapes** on
   them via callbacks (set text, feed chart data, fill tables, swap images).
4. `write()` the result.

Key rule: all `addSlide` / `modifyElement` / `addElement` calls only **queue**
work. Everything executes when you call `await pres.write(...)` (or `.stream()`
/ `.getJSZip()`). Nothing is applied before that.

If the user has no template file yet, tell them to create one in PowerPoint
first — that is the intended workflow. Only generate shapes from scratch via the
PptxGenJS bridge (see below) when there is no reasonable template shape.

<!-- include: docs/concepts.md -->

<!-- include: docs/getting-started.md § Basic Example -->

<!-- include: docs/selectors.md -->

<!-- include: docs/text.md -->

<!-- include: docs/tables.md -->

<!-- include: docs/charts.md -->

<!-- include: docs/images.md -->

<!-- include: docs/masters-layouts.md -->

<!-- include: docs/generation.md -->

<!-- include: docs/hyperlinks.md -->

<!-- include: docs/slide-management.md -->

<!-- include: docs/helpers.md -->

<!-- include: docs/output.md -->

<!-- include: docs/limitations.md -->

<!-- include: docs/troubleshooting.md -->

## Rules for the AI assistant

1. **Always `await` the output call** (`write`/`stream`/`getJSZip`) — nothing
   happens without it. One output call per Automizer instance.
2. **One Automizer instance per presentation build**, a fresh instance per
   output file. Running several instances concurrently in the same process is
   supported.
3. Shape callbacks receive raw XML elements for advanced use, but **prefer the
   `modify.*` helpers** — only fall back to XML manipulation (`XmlHelper`,
   `modify.dump(element)` for debugging) when no helper exists.
4. Slide numbers are **1-based**. Template files are addressed by filename or
   the alias given to `.load()`.
5. `templateDir`/`outputDir` must exist; media loading requires the root
   template to be loaded first.
6. Charts must be **native charts** in the template, tables native tables. A
   grouped-shape "fake table" cannot be filled with `setTable`.
7. If the output opens with a PowerPoint repair prompt: retry with
   `cleanup: false` and `cleanupPlaceholders: false`, remove exotic elements
   (animations, embedded videos, think-cell remnants) from the template, and
   isolate the failing slide by bisecting `addSlide` calls.
8. When a modification doesn't apply, don't guess selector variations:
   `write()` rejects with a typed error (`ElementNotFoundError`,
   `CallbackError`) naming slide and element — list the real shape names via
   `pres.getInfo()` and fix the selector. With `continueOnError: true`,
   failures are only logged and skipped, so check the console output then.
9. Keep data preparation (fetching/transforming) separate from the
   Automizer chain; build plain `ChartData`/`TableData` objects first, then
   apply them in modifiers.
10. Suggested skeleton for any task: *load root → load templates → loop over
    data → addSlide with modifiers → write → log summary*.

## Minimal complete example

```ts
import Automizer, { modify } from 'pptx-automizer';

const automizer = new Automizer({
  templateDir: `${__dirname}/templates`,
  outputDir: `${__dirname}/output`,
  removeExistingSlides: true,
  autoImportSlideMasters: true,
});

const regions = [
  { name: 'North', q1: 100, q2: 120 },
  { name: 'South', q1:  80, q2: 140 },
];

const pres = automizer
  .loadRoot('Root.pptx')
  .load('Report.pptx', 'report');

for (const region of regions) {
  pres.addSlide('report', 1, (slide) => {
    slide.modifyElement('Title', modify.setText(`Sales: ${region.name}`));
    slide.modifyElement('Chart', modify.setChartData({
      series: [{ label: 'Revenue' }],
      categories: [
        { label: 'Q1', values: [region.q1] },
        { label: 'Q2', values: [region.q2] },
      ],
    }));
  });
}

const summary = await pres.write('sales-report.pptx');
console.log(`Done: ${summary.slides} slides in ${summary.file}`);
```

More examples: the `__tests__/` directory of the repo covers nearly every
feature (94 runnable scenarios). Full documentation:
https://singerla.github.io/pptx-automizer/ — every page is also served as
Markdown at its URL plus `.md`, indexed in
https://singerla.github.io/pptx-automizer/llms.txt.
Repo: https://github.com/singerla/pptx-automizer
