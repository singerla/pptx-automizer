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
  verbosity: 1,                          // 0 errors only … 2 chatty
  // logger: myILogger,                  // inject custom logging (NullLogger = silent)
  // continueOnError: true,              // log & skip failing modifications
  //                                     // instead of rejecting write()
});
```

Error handling: a throwing modification callback or an unresolvable element
selector rejects `write()` with a typed error (`CallbackError`,
`ElementNotFoundError`, both `instanceof AutomizerError`) unless
`continueOnError: true` is set.

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
const slides = info.slidesByTemplate('content'); // visible slides, in deck order
const shape  = info.elementByName('content', 1, 'Title');
// slide.number is the number of the slide file (use it with addSlide),
// not the position in the deck.
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

Inside a `setMultiText` run, `\n` and `\v` (U+000B — what PowerPoint itself
returns for a Shift+Enter) become a **soft line break** (`<a:br/>`) within the
same paragraph. Use a new `MultiTextParagraph` when you want a real paragraph
(with its own bullet, level and alignment) instead.

**`htmlToMultiText` — input contract:** wrap the markup in
`<html><body>…</body></html>` (without a `<body>` you get an empty result and an
error log). The parser is an XML parser, so the input should be the well-formed
markup a WYSIWYG editor produces — quote attributes, close container tags.
`&nbsp;`, `&amp;` and an unclosed `<br>` do work.

Supported: `<p>`, `<div>`, `<h1>`–`<h6>`, `<blockquote>`, `<pre>` (paragraphs);
`<ul>`/`<ol>` + `<li>`, nested either properly (`<li>a<ul>…</ul></li>`) or in
CKEditor's sibling form (`<ul><li/><ul>…</ul></ul>`) — both give the same
result, with `<ol>` rendering as real automatic numbering (1. / a. / i. per
level); `<strong>`/`<b>`, `<em>`/`<i>`, `<u>`/`<ins>`, `<s>`/`<strike>`/`<del>`,
`<sub>`, `<sup>`, `<code>`/`<kbd>` (monospace), `<mark>` (highlight), `<br>`,
`<a href="url">` / `<a href="3">` (slide link), `<font color size face>`.
CSS on *any* element: `font-size` (px→pt at 96dpi, or `pt`), `color` and
`background-color` (hex, `rgb()`, `rgba()`, named), `font-weight`, `font-style`,
`text-decoration`, `font-family`, `text-align`.

Notes: colors may be written any CSS way (`#f00`, `red`, `rgb(255,0,0)`) — they
are normalized to OOXML's 6-digit hex. Relative font sizes (`em`, `%`) are
*ignored* rather than guessed, so the size stays inherited. Whitespace is
collapsed like a browser does; `&nbsp;` survives. Text alignment is only set
when the HTML says so, otherwise the shape's layout decides. Unsupported CSS is
skipped silently. For anything beyond this, build paragraphs with `setMultiText`
directly.

Two hyperlink gotchas: `<a href="4">` is a slide number in the **finished output
deck** (existing root-template slides included, not your `addSlide()` order), and
the slide must exist — otherwise the relationship dangles and the text renders
underlined but unlinked, with no error. And a `color` on an `<a>` has no visible
effect: PowerPoint always paints links in the theme's `<a:hlink>` color.

`TextStyle`: `{ size?, color?: {type: 'srgbClr'|'schemeClr', value}, isBold?,
isItalics?, isUnderlined?, isStrike?, isSubscript?, isSuperscript?, fontFamily?,
highlight?, hyperlink? … }`. `size` is in 1/100 pt (`1400` = 14pt). Colors are
hex without `#` (`'FF0000'`) or scheme names
(`{ type: 'schemeClr', value: 'accent1' }`).

A `MultiTextParagraph`'s `paragraph` takes `{ level (0-based, 0-8), bullet,
bulletType: 'char'|'number', bulletChar, autoNumberType, alignment, lineSpacing,
spaceBefore, spaceAfter, indent, marginLeft }`. A run may carry
`{ break: true }` instead of text to emit a soft line break explicitly.

### Position & style

```ts
import { CmToDxa } from 'pptx-automizer';
modify.setPosition({ x: CmToDxa(2), y: CmToDxa(1), w: CmToDxa(10), h: CmToDxa(5) });
modify.updatePosition({ x: CmToDxa(1) });   // only given props, relative-safe
modify.rotateShape(45);
modify.setSolidFill({ type: 'srgbClr', value: '00FF00' });

// Outline/border of a shape: width (EMU, 1pt = 12700), color, dash style.
// Only given props are touched; an a:ln is created if the shape has none.
import { PtToEmu } from 'pptx-automizer';
modify.setOutline({
  weight: PtToEmu(2),
  color: { type: 'srgbClr', value: 'FF0000' },
  type: 'sysDash', // any a:prstDash value: solid, dot, dash, lgDashDot, …
});
```

If a shape's outline was switched off in PowerPoint, `weight` alone stays
invisible — pass a `color` too. `ModifyCleanupHelper.removeBorder` removes an
outline entirely.

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

## Escape hatch: raw XML callbacks

Not every OOXML property has a `modify.*` helper (outline weight, `cap`/`cmpd`
line attributes, custom geometry, `a:effectLst`, …). Any callback you pass to
`modifyElement`/`addElement` receives the shape's XML element — the `<p:sp>`,
`<p:pic>` or `<p:graphicFrame>` node — so you can edit it directly:

```ts
import { XmlElement } from 'pptx-automizer';

slide.modifyElement('MyBox', (element: XmlElement) => {
  // ...manipulate the DOM node...
});
```

Rules for hand-written callbacks:

1. **Scope your lookups.** `element.getElementsByTagName('a:ln')` searches *all*
   descendants of the shape, including text run properties (`a:rPr`) — where
   `a:ln` also occurs. Reach the container first
   (`element.getElementsByTagName('p:spPr')[0]`), then prefer a **direct child**
   scan over another `getElementsByTagName`.
2. **A missing element is the normal case.** If a property was never overridden
   in PowerPoint, it is inherited from the theme/shape style and simply absent
   from the slide XML. Always handle "modify existing" *and* "create new".
3. **Child order follows the schema, not your call order.** Appending in the
   wrong order is what makes PowerPoint show the "repair" prompt on open. For
   `p:spPr` the sequence is `a:xfrm` → geometry (`a:prstGeom`/`a:custGeom`) →
   fill → `a:ln` → `a:effectLst` → `a:scene3d` → `a:sp3d` → `a:extLst`; inside
   `a:ln` it is fill → `a:prstDash` → join (`a:round`/`a:bevel`/`a:miter`) →
   `a:headEnd` → `a:tailEnd`.
4. **Inspect before you guess:** `slide.modifyElement('MyBox', modify.dump)`
   prints the shape's current XML to the console. Do that first when unsure what
   the template actually contains.
5. **Exceptions inside callbacks are swallowed** (logged via `console.warn`), and
   the deck is still written. A silent no-op means: read the console.
6. Use `XmlHelper` (exported) for common DOM chores: `XmlHelper.remove(node)`,
   `insertAfter(new, ref)`, `getClosestParent('p:sp', node)`,
   `appendClone(node, parent)`, `dump(node)`.

### Worked example: shape outline (weight + color)

> Outlines now have a dedicated modifier — use `modify.setOutline(...)` from the
> cheat sheet above for real work. The example is kept because it shows the
> general technique on a realistic property.

```ts
import Automizer, { XmlElement } from 'pptx-automizer';

// p:spPr children that must stay AFTER a:ln
const AFTER_LN = ['a:effectLst', 'a:effectDag', 'a:scene3d', 'a:sp3d', 'a:extLst'];
const childByName = (parent: XmlElement, names: string[]) =>
  Array.from(parent.childNodes as any).find((n: any) =>
    names.includes(n.nodeName),
  ) as XmlElement | undefined;

const setOutline =
  (outline: { weight?: number; color?: string }) =>
  (element: XmlElement) => {
    const spPr = element.getElementsByTagName('p:spPr')[0];
    if (!spPr) return;

    // Direct child only — a:ln also lives inside text run properties.
    let ln = childByName(spPr, ['a:ln']);
    if (!ln) {
      ln = spPr.ownerDocument.createElement('a:ln');
      const anchor = childByName(spPr, AFTER_LN);
      anchor ? spPr.insertBefore(ln, anchor) : spPr.appendChild(ln);
    }

    if (outline.weight !== undefined) {
      // a:ln/@w is EMU: 1pt = 12700
      ln.setAttribute('w', String(Math.round(outline.weight * 12700)));
    }

    if (outline.color) {
      const solidFill = ln.ownerDocument.createElement('a:solidFill');
      const srgbClr = ln.ownerDocument.createElement('a:srgbClr');
      srgbClr.setAttribute('val', outline.color.replace('#', '')); // no '#'!
      solidFill.appendChild(srgbClr);

      // Fill is the FIRST child of a:ln — replace whatever fill is there.
      const currentFill = childByName(ln, [
        'a:noFill', 'a:solidFill', 'a:gradFill', 'a:pattFill',
      ]);
      currentFill
        ? ln.replaceChild(solidFill, currentFill)
        : ln.insertBefore(solidFill, ln.firstChild);
    }
  };

slide.modifyElement('MyBox', setOutline({ weight: 2, color: 'FFFFFF' }));
```

Produces `<a:ln w="25400"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln>`.
Caveat worth knowing: if the shape's existing `a:ln` contains `<a:noFill/>`
(outline explicitly turned off in PowerPoint), setting only the weight yields
`<a:ln w="…"><a:noFill/></a:ln>` — a thick *invisible* line. Set a color too, or
replace the `a:noFill` node.

### Units reference

OOXML uses no single unit. When writing raw XML:

| What | Unit | Conversion |
|---|---|---|
| Position/size (`a:off`, `a:ext`), line width `a:ln/@w`, corner radius | EMU | 1 cm = 360000 · 1 inch = 914400 · 1 pt = 12700 |
| `modify.setPosition` / `updatePosition` / `setOutline` | same EMU values | helpers `CmToDxa(cm)` / `DxaToCm(v)` (name says Dxa, value is EMU), `PtToEmu(pt)` / `EmuToPt(v)` |
| Rotation (`a:xfrm/@rot`) | 1/60000 degree | 45° = 2700000 |
| Font size (`a:rPr/@sz`, `TextStyle.size`) | 1/100 pt | 18pt = 1800 |
| Percentages (`a:alpha/@val`, `a:lumMod`, …) | 1/1000 % | 50% = 50000 |
| Colors (`a:srgbClr/@val`) | 6-digit hex | `'FF0000'`, never `'#FF0000'` |
| PptxGenJS `slide.generate(...)` | inches | (different world — see below) |

If you write a raw callback that turns out to be generally useful, it is a good
candidate for a `modify.*` helper — see AGENTS.md in the repository.

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
