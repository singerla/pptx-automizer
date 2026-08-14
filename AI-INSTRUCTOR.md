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

## Concepts

Four things explain almost everything about how `pptx-automizer` behaves. Read this page before anything else.

### Deferred execution

The single most important thing to know: calls like `addSlide()`, `addElement()`
and `modifyElement()` only **queue** work — including every modification
callback you pass. Nothing is applied to the output presentation until you call
`await pres.write(...)` (or `.stream()` / `.getJSZip()`), which executes the
whole queue in order. Two practical consequences:

- A throwing callback or an unresolvable element selector surfaces as a rejected
  `write()` (with a typed error such as `CallbackError`), not at the line where
  you queued it. Set `continueOnError: true` to log a warning and skip the
  failing modification instead.
- Anything you compute *inside* a callback runs at `write()` time; variables it
  closes over will have their values from that moment, not from when the
  callback was queued.

### The template/root model

Every build starts from a **root template** (`loadRoot`) that the output
presentation is based on, plus any number of further templates (`load`) that
serve as shape and slide sources. This is how it works internally:

- Load a root template to append slides to
- (Probably) load root template again to modify slides
- Load other templates
- Append a loaded slide to (probably truncated) root template
- Modify the recently added slide
- Write root template and appended slides as output presentation

`pptx-automizer` is currently limited to _adding_ things to the output
presentation. If you require the ability to, for instance, modify a specific
element on a slide within an existing presentation and leave the rest
untouched, you will need to include all the other slides in the process — see
[looping through the slides of a presentation](https://singerla.github.io/pptx-automizer/slide-management.md#loop-through-the-slides-of-a-presentation).

### 1-based numbering

Slide numbers are **1-based**: `addSlide('shapes', 1)` takes the first slide of
the template labelled `shapes`. The number addresses the slide file inside the
.pptx — it is not necessarily the position in the final deck. Template files
are addressed by filename or the label given to `.load()`.

### One instance, one output

Use **one Automizer instance per presentation build**, and call the output
method (`write`/`stream`/`getJSZip`) **once** per instance. If you need
several output files, create a fresh `new Automizer(...)` per file. Separate
instances are isolated from each other — running several builds concurrently
in the same process (e.g. `Promise.all` in a server) is supported and covered
by a regression test.

## Basic Example

This is a basic example on how to use `pptx-automizer` in your code. Before you dive in, make sure you know about [deferred execution](https://singerla.github.io/pptx-automizer/concepts.md) — nothing is applied until `write()` is called. Most of the examples in these docs make use of [these template files](https://github.com/singerla/pptx-automizer/blob/main/__tests__/pptx-templates).

```ts
import Automizer from 'pptx-automizer';

// First, let's set some preferences!
const automizer = new Automizer({
  // this is where your template pptx files are coming from:
  templateDir: `my/pptx/templates`,

  // use a fallback directory for e.g. generic templates:
  templateFallbackDir: `my/pptx/fallback-templates`,

  // specify the directory to write your final pptx output files:
  outputDir: `my/pptx/output`,

  // turn this to true if you want to generally use
  // PowerPoint's creationIds instead of slide numbers
  // or shape names:
  useCreationIds: false,

  // Always use the original slideMaster and slideLayout of any
  // imported slide:
  autoImportSlideMasters: true,

  // truncate root presentation and start with zero slides
  removeExistingSlides: true,

  // activate `cleanup` to eventually remove unused files:
  cleanup: false,

  // Set a value from 0-9 to specify the zip-compression level.
  // The lower the number, the faster your output file will be ready.
  // Higher compression levels produce smaller files.
  compression: 0,

  // Please note: An image that is imported more than once (e.g. a logo on
  // several slides) is stored only once in the output file. There is no
  // setting for that, identical media files are always shared.

  // You can enable 'archiveType' and set mode: 'fs'.
  // This will extract all templates and output to disk.
  // It will not improve performance, but it can help debugging:
  // You don't have to manually extract pptx contents, which can
  // be annoying if you need to look inside your files.
  // archiveType: {
  //   mode: 'fs',
  //   baseDir: `${__dirname}/../__tests__/pptx-cache`,
  //   workDir: 'tmpWorkDir',
  //   cleanupWorkDir: true,
  // },

  // use a callback function to track pptx generation process.
  // statusTracker: myStatusTracker,

  // Console log verbosity: 0 shows errors only, 1 adds warnings (default),
  // 2 adds info & debug output. You can also inject a custom logger
  // implementing ILogger (or NullLogger for complete silence):
  // logger: new NullLogger(),
  verbosity: 1,

  // By default, a throwing modification callback or an unresolvable
  // element selector rejects write() with a typed error (CallbackError,
  // ElementNotFoundError). Set to true to log a warning and skip the
  // failing modification instead:
  // continueOnError: false,

  // Remove all unused placeholders to prevent unwanted overlays:
  cleanupPlaceholders: false,

  // Use a customized version of PptxGenJS if required:
  // pptxGenJs: PptxGenJS,
});

// Now we can start and load a pptx template.
// With removeExistingSlides set to 'false', each addSlide will append to
// any existing slide in RootTemplate.pptx. Otherwise, we are going to start
// with a truncated root template.
let pres = automizer
  .loadRoot('RootTemplate.pptx')
  // We want to make some more files available and give them a handy label.
  .load('SlideWithShapes.pptx', 'shapes')
  .load('SlideWithGraph.pptx', 'graph')
  // Skipping the second argument will not set a label.
  .load('SlideWithImages.pptx');

// Get useful information about loaded templates:
/*
const presInfo = await pres.getInfo();
// The visible slides of a template, in presentation order. Slides that were
// removed from the presentation (e.g. by 'removeExistingSlides') are skipped,
// even if their slide file is still contained in the .pptx.
const mySlides = presInfo.slidesByTemplate('shapes');
// 'slide.number' is the number of the slide file inside the .pptx and can be
// passed to addSlide(); it is not necessarily the position in the deck.
const mySlide = presInfo.slideByNumber('shapes', 2);
const myShape = presInfo.elementByName('shapes', 2, 'Cloud');
*/

// addSlide takes two arguments: The first will specify the source
// presentation's label to get the template from, the second will set the
// slide number to require.
pres
  .addSlide('graph', 1)
  .addSlide('shapes', 1)
  .addSlide('SlideWithImages.pptx', 2);

// Finally, we want to write the output file.
pres.write('myPresentation.pptx').then((summary) => {
  console.log(summary);
});
```

Templates don't have to come from disk: `load` and `loadRoot` also accept a
`Buffer` (e.g. from a database or an upload). A name is required for buffered
templates:

```ts
declare const rootBuffer: Buffer;
declare const contentBuffer: Buffer;

const pres = automizer
  .loadRoot(rootBuffer)
  .load(contentBuffer, 'content');
```

See [Output](https://singerla.github.io/pptx-automizer/output.md) for the other output formats — a `ReadableStream` or the underlying JSZip instance.

## Selecting slides and shapes

`pptx-automizer` needs a selector to find the required shape on a template slide. While an imported .pptx file is identified by filename or custom label, there are different ways to address its slides and shapes.

### Select slide by number and shape by name

If your .pptx-templates are more or less static and you do not expect them to evolve a lot, it's ok to use the slide number and the shape name to find the proper source of automation.

```ts
// This will take slide #2 from 'SlideWithGraph.pptx' and expect it
// to contain a shape called 'ColumnChart':
pres.addSlide('SlideWithGraph.pptx', 2, (slide) => {
  // `slide` is slide #2 of 'SlideWithGraph.pptx'
  slide.modifyElement('ColumnChart', [
    /* ... */
  ]);
});

// This example will take slide #1 from 'RootTemplate.pptx' and place
// 'ColumnChart' from slide #2 of 'SlideWithGraph.pptx' on it.
pres.addSlide('RootTemplate.pptx', 1, (slide) => {
  // `slide` is slide #1 of 'RootTemplate.pptx'
  slide.addElement('SlideWithGraph.pptx', 2, 'ColumnChart', [
    /* ... */
  ]);
});

// In case you have two or more shapes on your template slide with the same name,
// you can address the target by nameIdx:
pres.addSlide('SlideWithGraph.pptx', 1, (slide) => {
  // Starting from 0, this will find the second shape named 'ColumnChart'
  // on the slide.
  slide.modifyElement(
    {
      name: 'ColumnChart',
      nameIdx: 1,
    },
    [
      /* ... */
    ],
  );
});
```

> You can display and manage shape names directly in PowerPoint by opening the "Selection"-pane for your current slide. Hit `ALT+F10` and PowerPoint will give you a (nested) list including all (grouped) shapes. You can edit a shape name by double-click or by hitting `F2` after selecting a shape from the list. [See MS-docs for more info.](https://support.microsoft.com/en-us/office/use-the-selection-pane-to-manage-objects-in-documents-a6b2fd3e-d769-46c1-9b9c-b94e04a72550)

But be aware: Whenever your template slides are rearranged or a template shape is renamed, you need to update your code as well.

Please also make sure that each shape to add or modify has a unique name on its slide. Otherwise, only the last matching shape will be taken as target.

### Select slides by creationId

Additionally, each slide and shape is stored together with a (more or less) unique `creationId`. In XML, it looks like this:

```xml
<p:cNvPr name="MyPicture" id="64">
    <a:extLst>
        <a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}">
            <a16:creationId id="{0980FF19-E7E7-493C-8D3E-15B2100EA940}" xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main"/>
        </a:ext>
    </a:extLst>
</p:cNvPr>
```

This is where `name` and `creationId` are coupled together for each shape.

While our shape could now be identified by both, `MyPicture` or by `{0980FF19-E7E7-493C-8D3E-15B2100EA940}`, `creationIds` for slides consist of an integer value, e.g. `501735012` below:

```xml
<p:extLst>
   <p:ext uri="{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}">
      <p14:creationId val="501735012" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"/>
   </p:ext>
</p:extLst>
```

You can add a simple code snippet to get a list of the available `creationIds` of your loaded templates:

```ts
const pres = automizer
  .loadRoot(`RootTemplate.pptx`)
  .load(`SlideWithShapes.pptx`, 'shapes');

const creationIds = await pres.setCreationIds();

// This is going to print the slide creationId and a list of all
// shapes from slide #1 in `SlideWithShapes.pptx` (aka `shapes`).
console.log(
  creationIds
    .find((template) => template.name === 'shapes')
    .slides.find((slide) => slide.number === 1),
);
// Find the corresponding slide-creationId and -number on top of this list.
```

If your templates are not final and if you expect to have new slides and shapes added in the future, it is worth the effort and use `creationId` in general:

```ts
const automizer = new Automizer({
  templateDir: `${__dirname}/pptx-templates`,
  outputDir: `${__dirname}/pptx-output`,
  // turn this to true and use creationIds for both, slides and shapes
  useCreationIds: true,
});
```

Regarding shapes, it is also possible to use a `creationId` and the shape name as a fallback. These are the different types of a `FindElementSelector`:

```ts
import { FindElementSelector } from 'pptx-automizer';

// This is default when set up with `useCreationIds: true`:
const myShapeSelectorCreationId: FindElementSelector =
  '{E43D12C3-AD5A-4317-BC00-FDED287C0BE8}';

// pptx-generator will try to find the shape even if one of the given keys
// won't match any shape on the target slide:
const myShapeSelectorFallback: FindElementSelector = {
  creationId: '{E43D12C3-AD5A-4317-BC00-FDED287C0BE8}',
  name: 'Drum',
};

// Use this only if `useCreationIds: false`:
const myShapeSelectorName: FindElementSelector = 'Drum';

// Whenever `useCreationIds` was set to true, you need to replace slide numbers
// by `creationId`, too:
await pres.addSlide('shapes', 4167997312, (slide) => {
  // slide is now #1 of `SlideWithShapes.pptx`
  slide.addElement('shapes', 273148976, {
    creationId: '{E43D12C3-AD5A-4317-BC00-FDED287C0BE8}',
    name: 'Drum',
  });
  // 'Drum' is from #2 of `SlideWithShapes.pptx`, see __tests__ dir for an
  // example.
});
```

If you decide to use the `creationId` method, you are safe to add, remove and rearrange slides in your templates. It is also no problem to update shape names, and you also don't need to pay attention to unique shape names per slide.

> Please note: PowerPoint is going to update a shape's `creationId` only in case the shape was copied & pasted on a slide with an already existing identical shape `creationId`. If you were copying a slide, each shape `creationId` will be copied, too. As a result, you have unique shape ids, but different slide `creationIds`. If you are now going to paste a shape an such a slide, a new creationId will be given to the pasted shape. As a result, slide ids are unique throughout a presentation, but shape ids are unique only on one slide.

### Find and Modify Shapes

There are basically two ways to access a target shape on a slide:

- `slide.modifyElement(...)` requires an existing shape on the current slide,
- `slide.addElement(...)` adds a shape from another slide to the current slide.

Modifications can be applied to both in the same way:

```ts
import { modify, CmToDxa } from 'pptx-automizer';

pres.addSlide('shapes', 2, (slide) => {
  // This will only work if there is a shape called 'Drum'
  // on slide #2 of the template labelled 'shapes'.
  slide.modifyElement('Drum', [
    // You can use some of the builtin modifiers to edit a shape's xml:
    modify.setPosition({
      // set position from the left to 5 cm
      x: CmToDxa(5),
      // or use a number in DXA unit
      h: 5000000,
      w: 5000000,
    }),
    // Log your target xml into the console:
    modify.dump,
  ]);
});

pres.addSlide('shapes', 1, (slide) => {
  // This will import the 'Drum' shape from
  // slide #2 of the template labelled 'shapes'.
  slide.addElement('shapes', 2, 'Drum', [
    // add modifiers as seen in the example above
  ]);
});
```

### Inspect a slide from inside the callback

Besides `pres.getInfo()` (see the [Basic Example](https://singerla.github.io/pptx-automizer/getting-started.md#basic-example)),
an added slide can describe itself at `write()` time. Prefer inspecting over
guessing shape names — an unresolvable selector rejects `write()` (or is
skipped with `continueOnError: true`), so listing what is actually there beats
retrying name variations:

```ts
pres.addSlide('myTemplate.pptx', 1, async (slide) => {
  // All shapes on this slide, with name, id, position and type:
  const elements = await slide.getAllElements();
  // Only certain tags, e.g. text-bearing shapes:
  const shapes = await slide.getAllElements(['sp']);
  // The slide's dimensions:
  const dimensions = await slide.getDimensions();
});
```

### Find all text elements on a slide

When processing an added slide, you might want to apply a modifier to any existing text element. Call `slide.getAllTextElementIds()` for this:

```ts
import Automizer, { modify } from 'pptx-automizer';

pres.addSlide('myTemplate.pptx', 1, async (slide) => {
  const elements = await slide.getAllTextElementIds();
  elements.forEach((element) => {
    // element has a text body:
    slide.modifyElement(element, [modify.setText('my text')]);
    // ... or use the tag replace function:
    slide.modifyElement(element, [
      modify.replaceText([
        {
          replace: 'TAG',
          by: {
            text: 'my tag text',
          },
        },
      ]),
    ]);
  });
});
```

## Modify Text

You can select and import generic shapes from any loaded template. It is possible to update the containing text in several ways:

```ts
import { ModifyTextHelper, XmlElement } from 'pptx-automizer';

pres.addSlide('SlideWithImages.pptx', 1, (slide) => {
  // You can directly modify the child nodes of <p:sp>
  slide.addElement('shapes', 2, 'Arrow', (element: XmlElement) => {
    element.getElementsByTagName('a:t').item(0).textContent = 'Custom content';
  });

  // You might prefer a built-in function to set text:
  slide.addElement('shapes', 2, 'Arrow', [
    ModifyTextHelper.setText('This is my text'),
  ]);
});
```

### Replace tagged text

`pptx-automizer` also provides a powerful helper to replace tagged text. You can use e.g. `{{myTag}}` on your slide and apply a modifier to insert dynamic text. Font style can be inherited from template or updated by the modifier.

```ts
import { modify } from 'pptx-automizer';

pres.addSlide('TextReplace.pptx', 1, (slide) => {
  slide.modifyElement(
    // This is the name of the target element on slide #1 of
    // 'TextReplace.pptx
    'replaceText',
    // This will look for a string `{{replace}}` inside the text
    // contents of 'replaceText' shape
    modify.replaceText([
      {
        replace: 'replace',
        by: {
          text: 'Apples',
        },
      },
    ]),
  );
});
```

### MultiText: styled paragraphs and lists

You can use `modify.setMultiText` to replace all text contents of an existing textfield by styled paragraphs, bulleted lists and text runs:

```ts
import { modify } from 'pptx-automizer';

pres.addSlide('TextReplace.pptx', 1, (slide) => {
  slide.modifyElement(
    'setText',
    modify.setMultiText([
      {
        paragraph: {
          bullet: true,
          level: 0,
          marginLeft: 41338,
          indent: -87325,
          alignment: 'l',
        },
        textRuns: [
          {
            text: 'Bullet point level 0',
            style: {
              isItalics: true,
              color: {
                type: 'srgbClr',
                value: 'CCCCCC',
              },
            },
          },
        ],
      },
    ]),
  );
});
```

Within a text run, `\n` and `\v` (U+000B, the character PowerPoint itself uses for
a soft line break created with Shift+Enter) are converted into an `<a:br/>` line
break inside the same paragraph. Add another entry to the array if you need a real
paragraph with its own bullet, level and alignment.

### Convert HTML to text contents

It is also possible to directly convert an HTML page into pptx text contents. HTML code will be flattened and converted into a MultiText array.

```ts
import { modify } from 'pptx-automizer';

const html =
  '<html><body>' +
  '<h2 style="text-align: center">Quarterly report</h2>' +
  '<p>Plain text with <strong>bold</strong>, <em>italics</em> and ' +
  '<span style="color: #ff0000; font-size: 12pt">styling</span>.</p>' +
  '<ul>' +
  '<li>bullet level 0' +
  '<ul><li>bullet level 1</li></ul>' +
  '</li>' +
  '</ul>' +
  '<ol><li>numbered</li><li>list</li></ol>' +
  '<p><a href="https://example.com">external link</a> and ' +
  '<a href="3">a link to slide 3</a></p>' +
  '</body></html>';

pres.addSlide('TextReplace.pptx', 1, (slide) => {
  slide.modifyElement('setText', modify.htmlToMultiText(html));
});
```

#### What HTML is supported

PPTX text is strictly flat: a text body is a list of paragraphs, each a list of
text runs, with no nesting anywhere. HTML hierarchy is therefore *projected*
onto that — nested inline tags become one run with accumulated character
properties, nested lists become paragraphs with a 0-based level, and a block
inside a block (`<li><p>…</p></li>`) yields a single paragraph, with the
innermost block winning.

| | Supported |
|---|---|
| Paragraphs | `<p>`, `<div>`, `<h1>`–`<h6>`, `<blockquote>`, `<pre>`, `<section>` & friends |
| Lists | `<ul>`, `<ol>`, `<li>`, nested to 9 levels. `<ol>` renders as real automatic numbering (1. / a. / i. per level) |
| Inline | `<strong>`/`<b>`, `<em>`/`<i>`, `<u>`/`<ins>`, `<s>`/`<strike>`/`<del>`, `<sub>`, `<sup>`, `<code>`/`<kbd>`/`<samp>` (monospace), `<mark>`, `<br>`, `<span>`, `<a>`, `<font>` |
| Links | `<a href="https://…">` external, `<a href="3">` to slide 3 |
| CSS (on any element) | `font-size` (`px` converted at 96dpi, or `pt`), `color`, `background-color` (highlight), `font-weight`, `font-style`, `text-decoration`, `font-family`, `text-align` |

Both list-nesting styles work and produce identical output — properly nested
(`<li>text<ul>…</ul></li>`) and the sibling form CKEditor emits
(`<ul><li/><ul>…</ul></ul>`).

Good to know:

- The input has to be wrapped in `<html><body>…</body></html>`, and is parsed as
  XML: quote your attributes and close your container tags, as WYSIWYG editors
  do. `&nbsp;`, `&amp;` and an unclosed `<br>` are fine.
- Colors can be written in any CSS notation (`#f00`, `red`, `rgb(255,0,0)`) and
  are normalized to the 6-digit hex OOXML requires.
- Relative font sizes (`em`, `%`) are ignored rather than guessed, leaving the
  size inherited from the template.
- Whitespace collapses the way a browser collapses it; `&nbsp;` survives.
- Alignment is only written when the HTML asks for it — otherwise the target
  shape's layout keeps deciding.
- Font size and color of the target shape's existing text are used as the
  fallback style, so generated text keeps the template's look.
- `<table>` markup has no equivalent in a single text shape: the cell text is
  kept, but flattened into one paragraph. Use `modify.setTableData` for
  [tables](https://singerla.github.io/pptx-automizer/tables.md).
- `<a href="4">` is a slide **number in the finished output deck**, counting the
  root template's existing slides — not the index of your `addSlide()` calls.
  The target slide has to exist, or the relationship dangles and PowerPoint
  shows the text underlined but unlinked, without a warning.
- A `color` on an `<a>` element is written to the run, but PowerPoint paints
  hyperlink text in the theme's `<a:hlink>` color regardless. To restyle links,
  change that theme color; per-link colors are not achievable in PPTX.

### Text helpers (MultiText/HTML)

Generate complex text (multiple runs, links, bullets) either from a structured value or directly from HTML.

```ts
// From structured paragraphs
slide.modifyElement('TextBox', [
  ModifyTextHelper.setMultiText([
    {
      paragraph: { bullet: false },
      textRuns: [
        { text: 'Hello ', style: { isBold: true } },
        { text: 'World' },
      ],
    },
  ]),
]);

// From HTML - note the required <html><body> wrapper
const html =
  '<html><body><p><b>Bold</b> and ' +
  '<a href="https://example.com">link</a></p></body></html>';
slide.modifyElement('TextBox', [ModifyTextHelper.htmlToMultiText(html)]);
```

`HtmlToMultiTextHelper` and `MultiTextHelper` also support hyperlinks: an
external target (`<a href="https://...">`) or a slide number for an internal
link (`<a href="3">`). See [What HTML is supported](#what-html-is-supported)
above for the full tag and CSS coverage, and the tests:

- [Replace text by MultiText objects](https://github.com/singerla/pptx-automizer/blob/main/__tests__/replace-multi-text.test.ts)
- [Replace text by HTML](https://github.com/singerla/pptx-automizer/blob/main/__tests__/replace-multi-text-html.test.ts)
- [HTML conversion rules, unit level](https://github.com/singerla/pptx-automizer/blob/main/__tests__/html-to-multitext-converter.test.ts)

### Find out more

- [Replace and style by tags](https://github.com/singerla/pptx-automizer/blob/main/__tests__/replace-tagged-text.test.ts)
- [Modify text elements using getAllTextElementIds](https://github.com/singerla/pptx-automizer/blob/main/__tests__/get-all-text-element-ids.test.ts)
- [Replace text by multitext objects](https://github.com/singerla/pptx-automizer/blob/main/__tests__/replace-multi-text.test.ts)
- [Soft line breaks inside a text run](https://github.com/singerla/pptx-automizer/blob/main/__tests__/replace-multi-text-linebreaks.test.ts)
- [Replace text by HTML](https://github.com/singerla/pptx-automizer/blob/main/__tests__/replace-multi-text-html.test.ts)

## Modify Tables

You can use a PowerPoint table and add/modify data and style. It is also possible to add rows and columns and to style cells.

```ts
const pres = automizer
  .loadRoot(`RootTemplate.pptx`)
  .load(`SlideWithTables.pptx`, 'tables');

const result = await pres.addSlide('tables', 3, (slide) => {
  slide.modifyElement('TableWithEmptyCells', [
    modify.setTable({
      // Use an array of rows to insert data.
      // use `label` key for your information only
      body: [
        { label: 'item test r1', values: ['test1', 10, 16, 12, 11] },
        { label: 'item test r2', values: ['test2', 12, 18, 15, 12] },
        { label: 'item test r3', values: ['test3', 14, 12, 11, 14] },
      ],
    }),
  ]);
});
```

Note that the table has to be a **native table** in the template — a grouped-shape "fake table" cannot be filled with `setTable`.

### Table helpers

`ModifyTableHelper` provides rich control over existing tables.

- Fill table data and auto-adjust size

```ts
slide.modifyElement('MyTable', [
  ModifyTableHelper.setTable({
    body: [
      { label: 'r1', values: ['A', 1] },
      { label: 'r2', values: ['B', 2] },
    ],
  }),
]);
```

- Expand rows/columns by tag before filling

```ts
slide.modifyElement('MyTable', [
  ModifyTableHelper.setTable(
    {
      body: [ /* ... */ ],
    },
    {
      expand: [
        { tag: '<<ROW>>', count: 3, mode: 'row' },
        { tag: '<<COL>>', count: 2, mode: 'column' },
      ],
      adjustHeight: true,
      adjustWidth: true,
    }
  ),
]);
```

- Set fixed row heights / column widths

```ts
slide.modifyElement('MyTable', [
  ModifyTableHelper.updateRowHeight(0, CmToDxa(1)),
  ModifyTableHelper.updateColumnWidth(1, CmToDxa(3)),
]);
```

- Apply a table style and header/column banding flags

```ts
slide.modifyElement('MyTable', [
  ModifyTableHelper.setTableStyle('TableStyleMedium2', [
    'firstRow', 'bandRow',
  ]),
]);
```

Additional convenience methods:

- `ModifyTableHelper.setTableData(data)` – just set data without sizing
- `ModifyTableHelper.adjustHeight(data)` / `adjustWidth(data)` – recompute sizes only

### Find out more

- [Modify and style table cells](https://github.com/singerla/pptx-automizer/blob/main/__tests__/modify-existing-table.test.ts)
- [Insert data into table with empty cells](https://github.com/singerla/pptx-automizer/blob/main/__tests__/modify-existing-table-create-text.test.ts)

## Modify Charts

All data and styles of a chart can be modified. Please note that if your template contains more data than your data object, Automizer will remove these extra nodes. Conversely, if you provide more data, new nodes will be cloned from the first existing one in the template.

Note that the chart has to be a **native chart** in the template.

```ts
// Modify an existing chart on an added slide.
pres.addSlide('charts', 2, (slide) => {
  slide.modifyElement('ColumnChart', [
    // Use an object like this to inject the new chart data.
    // Additional series and categories will be copied from
    // previous sibling.
    modify.setChartData({
      series: [
        { label: 'series 1' },
        { label: 'series 2' },
        { label: 'series 3' },
      ],
      categories: [
        { label: 'cat 2-1', values: [50, 50, 20] },
        { label: 'cat 2-2', values: [14, 50, 20] },
        { label: 'cat 2-3', values: [15, 50, 20] },
        { label: 'cat 2-4', values: [26, 50, 20] },
      ],
    }),
  ]);
});
```

Find out more about modifying charts:

- [Modify chart axis](https://github.com/singerla/pptx-automizer/blob/main/__tests__/modify-chart-axis.test.ts)
- [Dealing with bubble charts](https://github.com/singerla/pptx-automizer/blob/main/__tests__/modify-chart-bubbles.test.ts)
- [Vertical line charts](https://github.com/singerla/pptx-automizer/blob/main/__tests__/modify-chart-vertical-lines.test.ts)
- [Style chart series and data points](https://github.com/singerla/pptx-automizer/blob/main/__tests__/modify-existing-chart-styled.test.ts)

### Modify Extended Charts

If you need to modify extended chart types, such like waterfall or map charts, you need to use `modify.setExtendedChartData`.

```ts
// Add and modify a waterfall chart on slide.
pres.addSlide('charts', 2, (slide) => {
  slide.addElement('ChartWaterfall.pptx', 1, 'Waterfall 1', [
    modify.setExtendedChartData(<ChartData>{
      series: [{ label: 'series 1' }],
      categories: [
        { label: 'cat 2-1', values: [100] },
        { label: 'cat 2-3', values: [50] },
        { label: 'cat 2-4', values: [-40] },
        // ...
      ],
    }),
  ]);
});
```

### Additional chart modifiers

Besides `modify.setChartData` and `modify.setExtendedChartData`, the `modify` namespace exposes a number of helpers to fine-tune chart appearance and special chart types. Each of them returns a modification callback that can be passed to `slide.modifyElement()`.

- Chart title (requires an already existing, manually edited title)

```ts
slide.modifyElement('ColumnChart', [modify.setChartTitle('My new title')]);
```

- Axis range and number format. Only manually scaled (non-"Auto") min/max values can be altered.

```ts
slide.modifyElement('ColumnChart', [
  modify.setAxisRange({
    axisIndex: 0, // index of c:valAx, defaults to 0
    min: 0,
    max: 100,
    majorUnit: 20,
    minorUnit: 5,
    formatCode: '0.0',
    sourceLinked: false,
  }),
]);
```

- Legend position and visibility. Legend coordinates are shares of the chart coordinates (e.g. `w: 0.5` means "half of chart width").

```ts
slide.modifyElement('ColumnChart', [
  // move/resize the legend
  modify.setLegendPosition({ x: 0.8, y: 0.1, w: 0.2, h: 0.3 }),
  // set legend coordinates to zero so a user can maximize it easily
  modify.minimizeChartLegend(),
  // completely remove the legend (PowerPoint will maximize the chart space)
  modify.removeChartLegend(),
]);
```

- Plot area. Requires a `c:manualLayout` element, which only exists if the plot area was edited manually in PowerPoint before. Coordinates are shares of the chart coordinates.

```ts
slide.modifyElement('ColumnChart', [
  modify.setPlotArea({ x: 0.1, y: 0.1, w: 0.8, h: 0.8 }),
]);
```

- Data labels

```ts
import { LabelPosition } from 'pptx-automizer';

slide.modifyElement('ColumnChart', [
  // format the data labels of all series (or a single series by index)
  modify.setDataLabelAttributes({
    applyToSeries: 0, // omit to apply to all series
    dLblPos: LabelPosition.OutsideEnd,
    formatCode: '0.0%',
    sourceLinked: false,
    showVal: true,
    showSerName: false,
    showCatName: false,
    showPercent: false,
    showLegendKey: false,
  }),
  // or remove all data labels
  modify.removeDataLabels(),
]);
```

- Waterfall total column. Mark the last column (or a specific index) of an extended waterfall chart as a total.

```ts
slide.modifyElement('Waterfall 1', [
  modify.setWaterFallColumnTotalToLast(), // or pass an index, e.g. (3)
]);
```

- Special chart types. Use dedicated modifiers for scatter, combo, bubble and vertical line charts. They accept the same `ChartData` object as `setChartData`.

```ts
slide.modifyElement('ScatterChart', [modify.setChartScatter(chartData)]);
slide.modifyElement('ComboChart', [modify.setChartCombo(chartData)]);
slide.modifyElement('BubbleChart', [modify.setChartBubbles(chartData)]);
slide.modifyElement('VerticalLineChart', [
  modify.setChartVerticalLines(chartData),
]);
```

### Read chart data

The `read` namespace provides callbacks to read information out of a chart without modifying it. They populate the object you pass in (see `__tests__/read-chart-data.test.js`).

```ts
import { read, WorkbookData, ChartInfo } from 'pptx-automizer';

const workbookData: WorkbookData = [];
const chartInfo: ChartInfo = { series: [] };

slide.modifyElement('ColumnChart', [
  // read the raw rows of the embedded workbook into workbookData
  read.readWorkbookData(workbookData),
  // read series colors and the detected chart type into chartInfo
  read.readChartInfo(chartInfo),
]);
```

## Modify Images

`pptx-automizer` can extract images from loaded .pptx template files and add to your output presentation. You can use shape modifiers (e.g. for size and position) on images, too. Additionally, it is possible to load external media files directly and update relation `Target` of an existing image. This works on both, existing or added images.

```ts
const automizer = new Automizer({
  // ...
  // Specify a directory to import external media files from:
  mediaDir: `path/to/media`,
});

const pres = automizer
  .loadRoot(`RootTemplate.pptx`)
  // load one or more files from mediaDir
  .loadMedia([`feather.png`, `test.png`] /* or use a custom dir */)
  // and/or use a custom dir
  .loadMedia(`icon.png`, 'path/to/icons')
  .load(`SlideWithImages.pptx`, 'images');

pres.addSlide('images', 2, (slide) => {
  slide.modifyElement('imagePNG', [
    // Override the original media source of element 'imagePNG'
    // by an imported file:
    ModifyImageHelper.setRelationTarget('feather.png'),

    // You might need to update size
    ModifyShapeHelper.setPosition({
      w: CmToDxa(5),
      h: CmToDxa(3),
    }),
  ]);
});
```

Media can also come from memory instead of `mediaDir`:
`automizer.loadMediaBuffer('chart.png', myImageBuffer)` registers a `Buffer`
under a filename, usable with `setRelationTarget` like any loaded file.

Note: media loading requires the root template to be loaded first. An image that is imported more than once (e.g. a logo on several slides) is stored only once in the output file — identical media files are always shared.

This will also auto-crop the image to the new width and height,
based on the container aspect ratio, derived from the original image
and using the new image width and height based on the files loaded into
the presentation media folder using .loadMedia():

```ts
pres.addSlide('images', 2, (slide) => {
  slide.modifyElement('Image Placeholder', [
    ModifyImageHelper.setRelationTargetCover('feather.png', pres),
  ]);
});
```

### Image helpers

Swap images and apply duotone overlays.

```ts
// Point an existing image to a different media file loaded via .loadMedia()
slide.modifyElement('Image Placeholder', [
  ModifyImageHelper.setRelationTarget('feather.png'),
]);

// Auto-crop to cover based on new image aspect ratio (uses template media)
slide.modifyElement('Image Placeholder', [
  ModifyImageHelper.setRelationTargetCover('feather.png', pres),
]);

// Duotone overlay on images/icons
slide.modifyElement('Icon', [
  ModifyImageHelper.setDuotoneFill({
    color: { type: 'srgbClr', value: '00AAFF' },
    tint: 70000, // 0-100000
    satMod: 90000, // 0-100000
  }),
]);
```

### Find out more

- [Add external image](https://github.com/singerla/pptx-automizer/blob/main/__tests__/add-external-image.test.ts)
- [Modify duotone color overlay for images](https://github.com/singerla/pptx-automizer/blob/main/__tests__/modify-image-duotone.test.ts)
- [Swap image source on a slide master](https://github.com/singerla/pptx-automizer/blob/main/__tests__/modify-master-external-image.test.ts)

## Slide Masters and Layouts

`pptx-automizer` supports importing slide masters and their associated slide layouts into the output presentation. It is important to note that you cannot add, modify, or remove individual slide layouts directly. However, you have the flexibility to modify the underlying slide master, which can serve as a workaround for certain changes — each modification on a slideMaster will appear on all related slideLayouts.

Please be aware that importing slide layouts containing complex content, such as charts and images, is currently not supported. For instance, if a slide layout includes an icon that is not present on the slide master, this icon will break when the slide master is auto-imported into an output presentation. To avoid this issue, ensure that all images and charts are placed exclusively on a slide master and not on a slide layout.

### Import and modify slide Masters

You can import, modify and use one or more slideMasters and the related slideLayouts.

To specify the target index of the required slide master to import, you need to count slideMasters in your _template_ presentation.
To specify another slideLayout for an added output slide, you need to count slideLayouts in your _output_ presentation.

To add and modify shapes on a slide master, please take a look at [Find and Modify Shapes](https://singerla.github.io/pptx-automizer/selectors.md#find-and-modify-shapes) — masters use the same selectors and modifiers as slides.

```ts
// Import another slide master and all its slide layouts.
// Index 1 means, you want to import the first of all masters:
pres.addMaster('SlidesWithAdditionalMaster.pptx', 1, (master) => {
  // Modify a certain shape on the slide master:
  master.modifyElement(
    `MasterRectangle`,
    ModifyTextHelper.setText('my text on master'),
  );
  // Add a shape from an imported templated to the current slideMaster.
  master.addElement('SlideWithShapes.pptx', 1, 'Cloud 1');
});
```

Any imported slideMaster will be appended to the existing ones in the root template. If you have already e.g. one master with five layouts, and you import a new master coming with seven slide layouts, the first new layout will be #6.

```ts
// Import a slideMaster and its slideLayouts:
pres.addMaster('SlidesWithAdditionalMaster.pptx', 1);

// Add a slide and switch to another layout:
pres.addSlide('SlidesWithAdditionalMaster.pptx', 3, (slide) => {
  // use another master, e.g. the imported one from 'SlidesWithAdditionalMaster.pptx'
  // You need to pass the index of the desired layout after all
  // related layouts of all imported masters have been added to rootTemplate.
  slide.useSlideLayout(12);
});

// It is also possible to use the original slideLayout of any added slide:
pres.addSlide('SlidesWithAdditionalMaster.pptx', 3, (slide) => {
  // To use the original master from 'SlidesWithAdditionalMaster.pptx',
  // we can skip the argument:
  slide.useSlideLayout();
  // This will also auto-import the original slideMaster, if not done already,
  // and look for the created index of the source slideLayout.
});
```

Please notice: If your root template and your imported slides have an equal structure of slideMasters and slideLayouts, it won't be necessary to add slideMasters manually.

If you have trouble with messed up slideMasters, and if you don't worry about the impact on performance, you can try and set `autoImportSlideMasters: true` to always import all required files:

```ts
import Automizer from 'pptx-automizer';

const automizer = new Automizer({
  // ...

  // Always use the original slideMaster and slideLayout of any
  // imported slide:
  autoImportSlideMasters: true,
  // ...
});
```

## Generate shapes with PptxGenJS

This library wraps around [PptxGenJS](https://github.com/gitbrent/PptxGenJS) to generate shapes from scratch. It is possible to use the `PptxGenJS` wrapper to generate shapes on a slide.

Here's an example of how to use `pptxGenJS` to add a text shape to a slide:

```ts
pres.addSlide('empty', 1, (slide) => {
  // Use PptxGenJS to add text from scratch:
  slide.generate((pptxGenJSSlide) => {
    pptxGenJSSlide.addText('Test 1', {
      x: 1,
      y: 1,
      h: 5,
      w: 10,
      color: '363636',
    });
  }, 'custom object name');
});
```

You can as well create charts with `pptxGenJS`:

```ts
const dataChartAreaLine = [
  {
    name: 'Actual Sales',
    labels: ['Jan', 'Feb', 'Mar'],
    values: [1500, 4600, 5156],
  },
  {
    name: 'Projected Sales',
    labels: ['Jan', 'Feb', 'Mar'],
    values: [1000, 2600, 3456],
  },
];

pres.addSlide('empty', 1, (slide) => {
  // Use PptxGenJS to add generated content from scratch:
  slide.generate((pSlide, pptxGenJs) => {
    pSlide.addChart(pptxGenJs.ChartType.line, dataChartAreaLine, {
      x: 1,
      y: 1,
      w: 8,
      h: 4,
    });
  });
});
```

Note: inside `generate`, coordinates and sizes are in **inches** (PptxGenJS
convention) — unlike `modify.setPosition`, which uses EMU/DXA values (see the
[units reference](https://singerla.github.io/pptx-automizer/helpers.md#units-reference)).

You can use the following functions to generate shapes with `pptxGenJS`:

- addChart
- addImage
- addShape
- addTable
- addText

### Create a new hyperlinked text shape

It is also possible to create a new hyperlink from scratch with the `pptxGenJS` wrapper. This is useful if you want to add hyperlinks to shapes that are not part of the template. (To manage hyperlinks on *existing* template shapes, see [Hyperlink Management](https://singerla.github.io/pptx-automizer/hyperlinks.md).)

```ts
// Generate a new text shape pointing to an external site
slide.generate((pptxGenJSSlide) => {
  pptxGenJSSlide.addText(`External Link`, {
    hyperlink: { url: 'https://github.com' },
    x: 1,
    y: 1,
    w: 2.5,
    h: 0.5,
    fontSize: 12,
  });
});

// Or generate an internal hyperlink
slide.generate((pptxGenJSSlide) => {
  pptxGenJSSlide.addText(`Go to slide 3`, {
    hyperlink: { slide: 3 },
    x: 1,
    y: 1,
    w: 2.5,
    h: 0.5,
    fontSize: 12,
  });
});
```

See a complete example on [how to add a chart from scratch](https://github.com/singerla/pptx-automizer/blob/main/__tests__/generate-pptxgenjs-charts.test.ts).

## Hyperlink Management

PowerPoint presentations often use hyperlinks to connect to external websites or internal slides. The `pptx-automizer` provides simple and powerful functions to manage hyperlinks in your presentations.

### Add Hyperlinks to existing shapes

You can add hyperlinks to template text shapes using the `addHyperlink` helper function. The function accepts either a URL string for external links or a slide number for internal slide links:

```ts
import { XmlElement } from 'pptx-automizer';

// Add an external hyperlink
slide.modifyElement('TextShape', modify.addHyperlink('https://example.com'));

// Add an internal slide link (to slide 3)
slide.modifyElement('TextShape', (element: XmlElement, relation?: XmlElement) => {
  modify.addHyperlink(3)(element, relation);
});
```

The `addHyperlink` function will automatically detect whether the target is an external URL or an internal slide number and set up the appropriate relationship type and attributes.

### Update or remove existing hyperlinks

Use `modify.setHyperlinkTarget` to change the target of a hyperlink that already exists on a shape. The second argument controls whether the new target is external (default, `true`) or an internal slide link (`false`):

```ts
// Point an existing hyperlink to a new external URL
slide.modifyElement('TextShape', modify.setHyperlinkTarget('https://example.com'));

// Point an existing hyperlink to an internal slide (e.g. slide 5)
slide.modifyElement('TextShape', modify.setHyperlinkTarget(5, false));
```

Use `modify.removeHyperlink` to strip the hyperlink from a shape while keeping its text:

```ts
slide.modifyElement('TextShape', modify.removeHyperlink());
```

### Related

- To create a **new** hyperlinked shape from scratch, use the PptxGenJS wrapper — see [Generate shapes with PptxGenJS](https://singerla.github.io/pptx-automizer/generation.md#create-a-new-hyperlinked-text-shape).
- Inline hyperlinks inside generated text (external targets or slide numbers) are also supported by the [MultiText/HTML text helpers](https://singerla.github.io/pptx-automizer/text.md#text-helpers-multitexthtml).

## Slide Management

### Remove elements from a slide

You can as well remove elements from slides.

```ts
// Remove existing charts, images or shapes from added slide.
pres
  .addSlide('charts', 2, (slide) => {
    slide.removeElement('ColumnChart');
  })
  .addSlide('images', 2, (slide) => {
    slide.removeElement('imageJPG');
    slide.removeElement('Textfeld 5');
    slide.addElement('images', 2, 'imageJPG');
  });
```

### Loop through the slides of a presentation

If you would like to modify elements in a single .pptx file, it is important to know that `pptx-automizer` is not able to directly "jump" to a shape to modify it — see [the template/root model](https://singerla.github.io/pptx-automizer/concepts.md#the-templateroot-model) for how it works internally.

In case you need to apply modifications to the root template, you need to load it as a normal template:

```ts
import Automizer, {
  CmToDxa,
  ISlide,
  ModifyColorHelper,
  ModifyShapeHelper,
  ModifyTextHelper,
} from 'pptx-automizer';

const run = async () => {
  const automizer = new Automizer({
    templateDir: `path/to/pptx-templates`,
    outputDir: `path/to/pptx-output`,
    // this is required to start with no slides:
    removeExistingSlides: true,
  });

  let pres = automizer
    .loadRoot(`SlideWithShapes.pptx`)
    // We load it twice to make it available for modifying slides.
    // Defining a "name" as the second parameter makes it a little easier
    .load(`SlideWithShapes.pptx`, 'myTemplate');

  // This is brand new: get useful information about loaded templates:
  const myTemplates = await pres.getInfo();
  const mySlides = myTemplates.slidesByTemplate(`myTemplate`);

  // Feel free to create some functions to pre-define all modifications
  // you need to apply to your slides.
  type CallbackBySlideNumber = {
    slideNumber: number;
    callback: (slide: ISlide) => void;
  };
  const callbacks: CallbackBySlideNumber[] = [
    {
      slideNumber: 2,
      callback: (slide: ISlide) => {
        slide.modifyElement('Cloud', [
          ModifyTextHelper.setText('My content'),
          ModifyShapeHelper.setPosition({
            h: CmToDxa(5),
          }),
          ModifyColorHelper.solidFill({
            type: 'srgbClr',
            value: 'cccccc',
          }),
        ]);
      },
    },
  ];
  const getCallbacks = (slideNumber: number) => {
    return callbacks.find((callback) => callback.slideNumber === slideNumber)
      ?.callback;
  };

  // We can loop all slides and apply the callbacks if defined
  mySlides.forEach((mySlide) => {
    pres.addSlide('myTemplate', mySlide.number, getCallbacks(mySlide.number));
  });

  // This will result in an output presentation containing all slides of "SlideWithShapes.pptx"
  pres.write(`myOutputPresentation.pptx`).then((summary) => {
    console.log(summary);
  });
};

run().catch((error) => {
  console.error(error);
});
```

### Quickly get all slide numbers of a template

When calling `pres.getInfo()`, it will gather information about all elements on all slides of all templates. In case you just want to loop through all slides of a certain template, you can use this shortcut:

```ts
const slideNumbers = await pres
  .getTemplate('myTemplate.pptx')
  .getAllSlideNumbers();

for (const slideNumber of slideNumbers) {
  // do the thing
}
```

### Sort output slides

There are three ways to arrange slides in an output presentation.

1. By default, all slides will be appended to the existing slides in your root template. The order of `addSlide` calls will define slide sorting in the output presentation.

2. You can alternatively remove all existing slides by setting the `removeExistingSlides` flag to true. The first slide added with `addSlide` will be first slide in the output presentation. If you want to insert slides from root template, you need to load it a second time.

```ts
import Automizer from 'pptx-automizer';

const automizer = new Automizer({
  templateDir: `my/pptx/templates`,
  outputDir: `my/pptx/output`,

  // truncate root presentation and start with zero slides
  removeExistingSlides: true,
});

let pres = automizer
  .loadRoot(`RootTemplate.pptx`)
  // We load this twice to make it available for sorting slides
  .load(`RootTemplate.pptx`, 'root')
  .load(`SlideWithShapes.pptx`, 'shapes')
  .load(`SlideWithGraph.pptx`, 'graph');

pres
  .addSlide('root', 1) // First slide will be taken from root
  .addSlide('graph', 1)
  .addSlide('shapes', 1)
  .addSlide('root', 3) // Third slide from root will be appended
  .addSlide('root', 2); // Second and third slide will switch position

pres.write(`mySortedPresentation.pptx`).then((summary) => {
  console.log(summary);
});
```

3. Use `sortSlides`-callback
   You can pass an array of numbers and create a callback and apply it to `presentation.xml`.
   This will also work without adding slides.

Slides will be appended to the existing slides by slide number (starting from 1). You may find irritating results in case you skip a slide number.

```ts
import { ModifyPresentationHelper } from 'pptx-automizer';

//
// You may truncate root template or you may not
// ...

// It is possible to skip adding slides, try sorting an unmodified presentation
pres
  .addSlide('charts', 1)
  .addSlide('charts', 2)
  .addSlide('images', 1)
  .addSlide('images', 2);

const order = [3, 2, 4, 1];
pres.modify(ModifyPresentationHelper.sortSlides(order));
```

## Helper Utilities

This section documents handy helpers exposed by `pptx-automizer` that provide extra capabilities. You can import helpers directly from the package:

```ts
import {
  ModifyShapeHelper,
  ModifyTableHelper,
  ModifyCleanupHelper,
  ModifyTextHelper,
  ModifyImageHelper,
} from 'pptx-automizer';
```

Some helper families are documented with their feature area:

- [Table helpers](https://singerla.github.io/pptx-automizer/tables.md#table-helpers) (`ModifyTableHelper`)
- [Text helpers (MultiText/HTML)](https://singerla.github.io/pptx-automizer/text.md#text-helpers-multitexthtml) (`ModifyTextHelper`)
- [Image helpers](https://singerla.github.io/pptx-automizer/images.md#image-helpers) (`ModifyImageHelper`)

### Shape helpers

Use `ModifyShapeHelper` to quickly adjust common properties of shapes and text frames.

- Solid fill color

```ts
slide.modifyElement('MyShape', [
  // sets the shape's solid fill to theme color "accent6"
  ModifyShapeHelper.setSolidFill,
]);
```

- Outline (line) width, color and dash style

```ts
import { PtToEmu } from 'pptx-automizer';

slide.modifyElement('MyShape', [
  ModifyShapeHelper.setOutline({
    weight: PtToEmu(2), // EMU, 1pt = 12700
    color: { type: 'srgbClr', value: 'FF0000' },
    type: 'sysDash', // any a:prstDash value, e.g. solid, dash, dot
  }),
]);
```

Only the given properties are modified. An outline is created if the shape has
none, which is the case whenever it was never overridden in PowerPoint and is
inherited from the theme or shape style. If the outline was explicitly switched
off in PowerPoint, a `weight` alone stays invisible — pass a `color`, too. Use
`ModifyCleanupHelper.removeBorder` to remove an outline.

- Bullet list from strings

```ts
slide.modifyElement('MyTextBox', [
  ModifyShapeHelper.setBulletList(['Item 1', 'Item 2', 'Item 3']),
]);
```

- Replace tagged text (see [Replace tagged text](https://singerla.github.io/pptx-automizer/text.md#replace-tagged-text) for the
  tag concept; the default delimiters are `{{` and `}}`)

```ts
slide.modifyElement('MyTextBox', [
  ModifyShapeHelper.replaceText(
    [
      { replace: 'company', by: { text: 'Globex' } },
      { replace: 'year', by: { text: '2025', style: { isBold: true } } },
    ],
    { openingTag: '{{', closingTag: '}}' },
  ),
]);
```

- Position, size, rotation, rounded corners

```ts
import { CmToDxa } from 'pptx-automizer';

slide.modifyElement('MyShape', [
  // set absolute position/size
  ModifyShapeHelper.setPosition({ x: CmToDxa(2), y: CmToDxa(3), w: CmToDxa(6), h: CmToDxa(2) }),
  // update only some props, leave others untouched
  ModifyShapeHelper.updatePosition({ x: CmToDxa(4) }),
  // rotate clockwise in degrees
  ModifyShapeHelper.rotate(15),
  // rounded rectangle corners (0-100000)
  ModifyShapeHelper.roundedCorners(25000),
]);
```

### Cleanup helpers

`ModifyCleanupHelper` helps remove formatting noise from shapes when you need a clean base.

```ts
import { XmlElement } from 'pptx-automizer';

slide.modifyElement('MyShape', [
  ModifyCleanupHelper.removeBackground,
  ModifyCleanupHelper.removeBorder,
  ModifyCleanupHelper.removeEffects,
  // text-level cleanup
  ModifyCleanupHelper.clearTextUnderline,
  ModifyCleanupHelper.clearTextBold,
  ModifyCleanupHelper.clearTextSize,
  // remove all explicit text colors …
  (element: XmlElement) => ModifyCleanupHelper.clearTextColor(element),
  // … or pass a color to set a uniform one instead
  (element: XmlElement) =>
    ModifyCleanupHelper.clearTextColor(element, {
      type: 'srgbClr',
      value: 'FF0000',
    }),
]);
```

Other useful helpers include: `removeTextEffects`, `removeFillEffects`, `remove3dEffects`, `removeShadowEffects`, and `removeExtLst`.

### Unit conversion helpers

PowerPoint stores coordinates and sizes in the `dxa` (EMU) unit. Use the exported converters to work with centimeters instead:

```ts
import { CmToDxa, DxaToCm } from 'pptx-automizer';

// centimeters -> dxa (e.g. when setting position/size)
const widthInDxa = CmToDxa(6); // 2160000

// dxa -> centimeters (e.g. when reading shape coordinates)
const widthInCm = DxaToCm(2160000); // 6
```

Line weights are usually given in points, use `PtToEmu`/`EmuToPt` for those:

```ts
import { PtToEmu, EmuToPt } from 'pptx-automizer';

const weight = PtToEmu(1.5); // 19050
const inPoints = EmuToPt(19050); // 1.5
```

### Generic / debugging helpers

`ModifyHelper` (also available through the `modify` namespace) offers low-level callbacks that are handy for debugging or custom XML tweaks:

```ts
import { modify } from 'pptx-automizer';

slide.modifyElement('MyShape', [
  // print the element's XML to the console
  modify.dump,
  // print the related chart XML to the console
  modify.dumpChart,
  // set an attribute on the first matching tag (optionally by index)
  modify.setAttribute('a:off', 'x', 1000000),
]);
```

### Advanced XML helpers (power users)

For advanced scenarios, you can inspect slide XML and relationships. These are considered expert APIs and may change.

Import them from the package root:

```ts
import { XmlSlideHelper, XmlRelationshipHelper } from 'pptx-automizer';
```

Examples:

- Read all text element IDs on a slide: `new XmlSlideHelper(slideXml).getAllTextElementIds()`
- Get named elements: `new XmlSlideHelper(slideXml).getNamedElements(['p:sp'])`
- Table introspection: `XmlSlideHelper.readTableInfo(element)`
- Relationship targets by type or prefix: `new XmlRelationshipHelper(relsXml).getTargetsByType(type)`

See tests for practical usage:

- [Find all text elements on a slide](https://github.com/singerla/pptx-automizer/blob/main/__tests__/get-all-text-element-ids.test.ts)
- [Read shape/group info](https://github.com/singerla/pptx-automizer/blob/main/__tests__/read-shape-info.test.ts)
- [Read group info](https://github.com/singerla/pptx-automizer/blob/main/__tests__/read-group-info.test.ts)

### Writing raw XML callbacks

Not every OOXML property has a `modify.*` helper (`cap`/`cmpd` line attributes,
custom geometry, `a:effectLst`, …). Any callback you pass to
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
5. **A throwing callback rejects `write()`** with a `CallbackError` naming the
   slide and element — unless `continueOnError: true` is set, which logs a
   warning and skips the modification instead (see
   [deferred execution](https://singerla.github.io/pptx-automizer/concepts.md#deferred-execution)).
6. Use `XmlHelper` (exported) for common DOM chores: `XmlHelper.remove(node)`,
   `insertAfter(new, ref)`, `getClosestParent('p:sp', node)`,
   `appendClone(node, parent)`, `dump(node)`.

If a raw callback you wrote turns out to be generally useful, it is a good
candidate for a new `modify.*` helper — see
[AGENTS.md](https://github.com/singerla/pptx-automizer/blob/main/AGENTS.md) in
the repository.

#### Worked example: shape outline (weight + color)

> Outlines have a dedicated modifier — use `ModifyShapeHelper.setOutline` from
> the [shape helpers](#shape-helpers) for real work. The example is kept because
> it shows the general technique on a realistic property.

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

#### Units reference

OOXML uses no single unit. When writing raw XML:

| What | Unit | Conversion |
|---|---|---|
| Position/size (`a:off`, `a:ext`), line width `a:ln/@w`, corner radius | EMU | 1 cm = 360000 · 1 inch = 914400 · 1 pt = 12700 |
| `modify.setPosition` / `updatePosition` / `setOutline` | same EMU values | helpers `CmToDxa(cm)` / `DxaToCm(v)` (name says Dxa, value is EMU), `PtToEmu(pt)` / `EmuToPt(v)` |
| Rotation (`a:xfrm/@rot`) | 1/60000 degree | 45° = 2700000 |
| Font size (`a:rPr/@sz`, `TextStyle.size`) | 1/100 pt | 18pt = 1800 |
| Percentages (`a:alpha/@val`, `a:lumMod`, …) | 1/1000 % | 50% = 50000 |
| Colors (`a:srgbClr/@val`) | 6-digit hex | `'FF0000'`, never `'#FF0000'` |
| PptxGenJS `slide.generate(...)` | inches | (different world — see [Generate shapes](https://singerla.github.io/pptx-automizer/generation.md)) |

## Output

Calling an output method executes the whole queued modification chain (see [deferred execution](https://singerla.github.io/pptx-automizer/concepts.md#deferred-execution)) and produces the final archive. Always `await` the output call — nothing happens without it, and one output call per Automizer instance.

### write, stream, getJSZip

```ts
// Write the output file.
pres.write('myPresentation.pptx').then((summary) => {
  console.log(summary);
});

// It is also possible to get a ReadableStream.
// stream() accepts JSZip.JSZipGeneratorOptions for 'nodebuffer' type.
const stream = await pres.stream({
  compressionOptions: {
    level: 9,
  },
});
// You can e.g. output the pptx archive to stdout instead of writing a file:
stream.pipe(process.stdout);

// If you need any other output format, you can eventually access
// the underlying JSZip instance:
const finalJSZip = await pres.getJSZip();
// Convert the output to whatever needed:
const base64 = await finalJSZip.generateAsync({ type: 'base64' });
```

By default, a throwing modification callback or an unresolvable element selector rejects `write()` with a typed error (`CallbackError`, `ElementNotFoundError`). Set `continueOnError: true` to log a warning and skip the failing modification instead.

### Track status of automation process

When creating large presentations, you might want to have some information about the current status. Use a custom status tracker:

```ts
import Automizer, { StatusTracker } from 'pptx-automizer';

// If you want to track the steps of creation process,
// you can use a custom callback:
const myStatusTracker = (status: StatusTracker) => {
  console.log(status.info + ' (' + status.share + '%)');
};

const automizer = new Automizer({
  // ...
  statusTracker: myStatusTracker,
});
```

### Cleanup and compression flags

The `Automizer` constructor takes a few options that affect the output archive:

```ts
const automizer = new Automizer({
  templateDir: `my/pptx/templates`,
  outputDir: `my/pptx/output`,

  // activate `cleanup` to eventually remove unused files from the archive:
  cleanup: false,

  // Remove all unused placeholders to prevent unwanted overlays:
  cleanupPlaceholders: false,

  // Set a value from 0-9 to specify the zip-compression level.
  // The lower the number, the faster your output file will be ready.
  // Higher compression levels produce smaller files.
  compression: 0,
});
```

If an output file triggers PowerPoint's repair prompt, retry with `cleanup: false` and `cleanupPlaceholders: false` first — see [Troubleshooting](https://singerla.github.io/pptx-automizer/troubleshooting.md).

## Requirements and Limitations

This generator can only be used on the server-side and requires a [Node.js](https://nodejs.org/en/download/package-manager/) environment.

### Shape Types

At the moment, you might encounter difficulties with special shape types that require additional relations (e.g., hyperlinks, video and audio may not work correctly). However, most shape types, including connection shapes, tables, and charts, are already supported. If you encounter any issues, please feel free to [report any issue](https://github.com/singerla/pptx-automizer/issues/new).

### Chart Types

Extended chart types, like waterfall or map charts, are basically supported. You might need additional modifiers to handle extended properties, which are not implemented yet. Please help to improve `pptx-automizer` and [report](https://github.com/singerla/pptx-automizer/issues/new) issues regarding extended charts.

### Animations

Animations are currently out of scope of this library. You might get errors on opening an output .pptx when there are added or removed shapes. This is because `pptx-automizer` doesn't synchronize `id`-attributes of animations with the existing shapes on a slide.

### Slide Masters and Layouts

Slide masters and their layouts can be imported, but individual slide layouts cannot be added, modified or removed directly, and layouts must not carry complex content like charts and images. See [Slide Masters and Layouts](https://singerla.github.io/pptx-automizer/masters-layouts.md) for details and workarounds.

### Direct manipulation of elements

It is also important to know that `pptx-automizer` is currently limited to _adding_ things to the output presentation. If you require the ability to, for instance, modify a specific element on a slide within an existing presentation and leave the rest untouched, you will need to include all the other slides in the process. Find some workarounds in [Slide Management](https://singerla.github.io/pptx-automizer/slide-management.md#loop-through-the-slides-of-a-presentation).

### PowerPoint version

All testing focuses on PowerPoint 2019 .pptx file format.

## Troubleshooting

If you encounter problems when opening a `.pptx`-file modified by this library, you might worry about PowerPoint not giving any details about the error. It can be hard to find the cause, but there are some things you can check:

- **Broken relation**: There are still unsupported shape types and `pptx-automizer` will not copy required relations of those. You can inflate `.pptx` output and check `ppt/slides/_rels/slide[#].xml.rels` files to find possible missing files.
- **Unsupported media**: You can also take a look at the `ppt/media`-directory of an inflated `.pptx`-file. If you discover any unusual file formats, remove or replace the files by one of the [known types](https://github.com/singerla/pptx-automizer/blob/main/src/enums/content-type-map.ts).
- **Broken animation**: Pay attention to modified/removed shapes which are part of an animation. In case of doubt, (temporarily) remove all animations from your template. (see [#78](https://github.com/singerla/pptx-automizer/issues/78))
- **Proprietary/Binary contents** (e.g. ThinkCell): Walk through all slides, slideMasters and slideLayouts and seek for hidden Objects. Hit `ALT+F10` to toggle the sidebar.
- **Chart datasheet won't open** If you encounter an error message on opening a chart's datasheet, please make sure that the data table (blue bordered rectangle in worksheet view) of your template starts at cell A:1. If not, open worksheet in Excel mode and edit the table size in the table draft tab.

### Bisecting the repair prompt

If the output opens with a PowerPoint repair prompt and the checks above don't point to a cause:

1. Retry with `cleanup: false` and `cleanupPlaceholders: false`.
2. Remove exotic elements (animations, embedded videos, ThinkCell remnants) from the template.
3. Isolate the failing slide by bisecting your `addSlide` calls — comment out half of them, re-run, and narrow down until the offending slide (or modification) is found.

Another powerful check: run the repo's [OOXML schema validator](https://singerla.github.io/pptx-automizer/testing.md#ooxml-schema-validation-validatepptx) (`yarn validate:pptx`) on your output — it usually names the exact part and element PowerPoint chokes on.

If none of these could help, please don't hesitate to [talk about it](https://github.com/singerla/pptx-automizer/issues/new).

### Testing

You can run all unit tests using these commands:

```
yarn test
yarn test-coverage
```

See [Testing and Validation Tools](https://singerla.github.io/pptx-automizer/testing.md) for the full test system — archive invariants, the docs-examples compile gate, schema validation and visual regression.

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
