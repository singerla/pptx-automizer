# AI Instructor: Generating PowerPoint with pptx-automizer

> **For AI assistants (Claude, etc.):** This document teaches you how to write
> correct code with `pptx-automizer`. Follow it as authoritative instructions.
>
> **For humans:** Drop this file into your project (or paste it into your AI
> chat / reference it from your CLAUDE.md) and ask your assistant to build
> PowerPoint automation with it.

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

## Setup

```ts
import Automizer, { modify } from 'pptx-automizer';

const automizer = new Automizer({
  templateDir: `${__dirname}/templates`, // where template .pptx files live
  outputDir: `${__dirname}/output`,      // where write() puts results
  removeExistingSlides: true,            // start output with 0 slides (usual)
  autoImportSlideMasters: true,          // bring masters/layouts along (usual)
  cleanup: false,
  verbosity: 1,                          // 0 silent … 2 chatty
});
```

Load files (chainable, synchronous):

```ts
const pres = automizer
  .loadRoot('RootTemplate.pptx')     // theme/masters/slide size of the output
  .load('ContentSlides.pptx', 'content')  // 2nd arg = handy alias
  .load('Charts.pptx', 'charts');
```

Buffers work too: `automizer.load(buffer, 'name')` (name required for buffers);
`loadRoot(buffer)` for the root.

## The core loop: addSlide + modify

```ts
pres.addSlide('content', 1, (slide) => {
  // `slide` = copy of slide #1 of ContentSlides.pptx, now part of the output.
  slide.modifyElement('Title', modify.setText('Quarterly Report'));

  // Multiple modifiers: pass an array
  slide.modifyElement('Subtitle', [
    modify.setText('FY 2026'),
    modify.setPosition({ x: 1000000, y: 500000 }), // EMU/DXA units
  ]);

  // Pull a shape in from ANOTHER template's slide:
  slide.addElement('charts', 2, 'PieChart', modify.setChartData(myData));

  // Remove a shape:
  slide.removeElement('DraftWatermark');
});

const summary = await pres.write('report.pptx'); // don't forget the await!
```

- `slide.modifyElement(selector, cb)` — shape already on this slide.
- `slide.addElement(templateNameOrAlias, slideNumber, selector, cb?)` — copy a
  shape from any loaded template onto this slide.
- Selectors are shape **names** (Selection pane). If names may be duplicated:
  `{ name: 'Box', nameIdx: 1 }` (0-based) targets the second 'Box'.
  With `useCreationIds: true` you can use PowerPoint creationIds instead:
  `{ creationId: '{E43D…}', name: 'Box' }` (name = fallback).

### Discovering what's in a template (do this when unsure)

```ts
const info = await pres.getInfo();
const slides = info.slidesByTemplate('content'); // slide numbers + elements
const shape  = info.elementByName('content', 1, 'Title');
// Or per slide inside a callback:
// const elements = await slide.getAllElements();
// const dims = await slide.getDimensions();
```

Prefer inspecting over guessing shape names — a wrong selector currently fails
**silently** (logged to console, deck still written). After generating code,
always check the console output and open the .pptx in PowerPoint.

## Modifier cheat sheet (`import { modify } from 'pptx-automizer'`)

### Text

```ts
modify.setText('Hello');                             // replace all text
modify.setBulletList(['One', 'Two', ['Nested']]);    // bullet lists
modify.replaceText(
  [{ replace: 'client', by: { text: 'ACME Corp' } }],   // {{client}} tags
  { openingTag: '{{', closingTag: '}}' },
);
// Rich multi-paragraph text (MultiTextParagraph[]):
modify.setMultiText([
  {
    paragraph: { alignment: 'l', bullet: false },
    textRuns: [
      { text: 'Bold red ', style: { isBold: true, color: { type: 'srgbClr', value: 'FF0000' } } },
      { text: 'plain rest' },
    ],
  },
  { paragraph: { level: 1, bullet: true }, textRuns: [{ text: 'a bullet' }] },
]);
modify.htmlToMultiText('<p>Simple <b>HTML</b> → text</p>');
```

**`htmlToMultiText` caveats (current state):** it supports a limited subset —
`<p>`, `<ul>`/`<ol>` (both render as `•` bullets), `<li>`, `<strong>`/`<b>`,
`<em>`/`<i>`, `<ins>` (underline), `<a href="url">` / `<a href="3">` (slide
link), and `<span style="font-size: Npx; color: …">`. Input must be
**well-formed XHTML wrapped in `<html><body>…</body></html>`** (the parser is
an XML parser: close every tag, no `&nbsp;`). Quirks to work around: write
colors as 6-digit hex **without `#`** (`color: FF0000`) and note that `font-size`
px values are applied as points. Not supported yet: `<br>`, `<u>`, `<s>`,
`<sub>`/`<sup>`, headings, `text-align`, numbered-list rendering. For anything
beyond this subset, build the paragraphs with `setMultiText` directly.

`TextStyle`: `{ size?, color?: {type: 'srgbClr'|'schemeClr', value}, isBold?,
isItalics?, isUnderlined?, hyperlink? … }`. Colors are hex without `#`
(`'FF0000'`) or scheme names (`{ type: 'schemeClr', value: 'accent1' }`).

### Position & style

```ts
import { CmToDxa } from 'pptx-automizer';
modify.setPosition({ x: CmToDxa(2), y: CmToDxa(1), w: CmToDxa(10), h: CmToDxa(5) });
modify.updatePosition({ x: CmToDxa(1) });   // only given props, relative-safe
modify.rotateShape(45);
modify.setSolidFill({ type: 'srgbClr', value: '00FF00' });
```

### Tables

Template needs a real PowerPoint table shape. Data shape:

```ts
const data = {
  body: [
    { label: 'Berlin',  values: ['Berlin', 120, 130] },
    { label: 'Munich',  values: ['Munich', 100,  90] },
  ],
};
slide.modifyElement('MyTable', modify.setTable(data));
// The table grows/shrinks to fit data. Styling per row:
// { values: [...], styles: [null, { isBold: true, background: {type:'srgbClr', value:'EEEEEE'} }] }
```

Also: `modify.adjustHeight/adjustWidth`, `modify.updateColumnWidth(idx, size)`,
`modify.updateRowHeight(idx, size)`.

### Charts

Template needs a native PowerPoint chart (not a picture of one!). The chart
keeps its template styling; you only feed data:

```ts
const data = {
  series: [{ label: '2025' }, { label: '2026' }],
  categories: [
    { label: 'Q1', values: [100, 150] },
    { label: 'Q2', values: [120, 180] },
  ],
};
slide.modifyElement('ColumnChart', modify.setChartData(data));
```

- Scatter: categories' `values` are `{x, y}` points + `modify.setChartScatter`.
- Bubbles: `{x, y, size}` + `modify.setChartBubbles`.
- Combo: `modify.setChartCombo`. Waterfall/extended: `modify.setExtendedChartData`.
- Styling per point: `categories[i].styles: [{ color?, background?, label? }]`.
- Extras: `modify.setAxisRange({ axisIndex, min, max, formatCode })`,
  `modify.setChartTitle('T')`,
  `modify.setLegendPosition({ x, y, w, h })` (shares of chart area, e.g. `w: 0.3`),
  `modify.removeChartLegend()`, `modify.setPlotArea({...})`,
  `modify.setDataLabelAttributes({ showVal: true, … })`, `modify.removeDataLabels()`.
- Read existing chart data: `import { read } from 'pptx-automizer'` →
  `read.readChartInfo` / `read.readWorkbookData` as modifyElement callbacks.

### Images

```ts
// Swap an existing image for a newly loaded media file:
automizer.loadMedia('newLogo.png', `${__dirname}/media`); // before addSlide chain
slide.modifyElement('LogoShape', modify.setRelationTarget('newLogo.png'));
// Buffers: automizer.loadMediaBuffer('chart.png', myBuffer)
// Recolor: modify.setDuotoneFill({ color: { type: 'srgbClr', value: '336699' } })
```

### Hyperlinks

```ts
modify.setHyperlinkTarget('https://new-url.example');
modify.addHyperlink('https://example.com');   // external
modify.addHyperlink(3);                       // internal → slide 3
modify.removeHyperlink();
```

## Generating shapes from scratch (PptxGenJS bridge)

For content with no template shape (rare — prefer templates):

```ts
pres.addSlide('content', 1, (slide) => {
  slide.generate((pptxGenJSSlide) => {
    pptxGenJSSlide.addText('Generated!', {
      x: 1, y: 1, w: 4, h: 1, color: '363636',
    });
    // Full PptxGenJS API: addChart, addImage, addTable, addShape …
  }, 'myGeneratedText'); // 2nd arg: name for the generated object
});
```

Note: inside `generate`, units are **inches** (PptxGenJS convention), unlike
`modify.setPosition` which uses DXA/EMU.

## Masters, layouts, output

```ts
// Usually just set autoImportSlideMasters: true and forget about it.
// Manual control:
pres.addMaster('OtherTemplate.pptx', 1);                  // import master #1
slide.useSlideLayout(4);                                  // or by name: useSlideLayout('Title and Content')
```

Output options:

```ts
const summary = await pres.write('out.pptx');             // → outputDir/out.pptx
const stream  = await pres.stream({ compressionOptions: { level: 9 } });
const jszip   = await pres.getJSZip();                    // → base64/blob/etc.
const b64     = await jszip.generateAsync({ type: 'base64' });
```

## Rules for the AI assistant

1. **Always `await` the output call** (`write`/`stream`/`getJSZip`) — nothing
   happens without it. One output call per Automizer instance.
2. **One Automizer instance per presentation build**; don't run two builds
   concurrently in the same process (shared internal state, known limitation).
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
8. When something doesn't apply, check the console: "Can't find element on
   slide …" means a wrong shape name — list names via `pres.getInfo()` and fix
   the selector rather than guessing variations.
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
feature (94 runnable scenarios), and the README documents each modifier in
detail. Repo: https://github.com/singerla/pptx-automizer
