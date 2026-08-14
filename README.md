# pptx-automizer: A Powerful .pptx Modifier for Node.js

[![npm version](https://img.shields.io/npm/v/pptx-automizer.svg)](https://www.npmjs.com/package/pptx-automizer)
[![CI](https://github.com/singerla/pptx-automizer/actions/workflows/ci.yml/badge.svg)](https://github.com/singerla/pptx-automizer/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://github.com/singerla/pptx-automizer/blob/main/LICENSE)

`pptx-automizer` is a Node.js-based PowerPoint (.pptx) generator that automates the manipulation of existing .pptx files. With `pptx-automizer`, you can import your library of .pptx templates, merge templates, and customize slide content. `pptx-automizer` will edit and merge existing pptx files. You can style template slides within PowerPoint, and these templates will be seamlessly integrated into the output presentation. Most of the content can be modified by using callbacks with [xmldom](https://github.com/xmldom/xmldom).

If you require to create elements from scratch, `pptx-automizer` wraps around [PptxGenJS](https://github.com/gitbrent/PptxGenJS). Use the powerful syntax of `PptxGenJS` to add dynamic content to your existing .pptx template files.

`pptx-automizer` is particularly well-suited for users who aim to manage their own library of .pptx template files, making it an ideal choice for those who work with intricate, well-designed customized layouts. With this tool, any existing slide or even a single element can serve as a data-driven template for generating output .pptx files.

## Documentation

**The full documentation lives at [singerla.github.io/pptx-automizer](https://singerla.github.io/pptx-automizer/)** — guides for every feature area plus the generated API reference:

- [Getting started](https://singerla.github.io/pptx-automizer/getting-started) · [Concepts](https://singerla.github.io/pptx-automizer/concepts) · [Selecting slides and shapes](https://singerla.github.io/pptx-automizer/selectors)
- Modify [Text](https://singerla.github.io/pptx-automizer/text) · [Tables](https://singerla.github.io/pptx-automizer/tables) · [Charts](https://singerla.github.io/pptx-automizer/charts) · [Images](https://singerla.github.io/pptx-automizer/images) · [Masters & Layouts](https://singerla.github.io/pptx-automizer/masters-layouts)
- [Generate shapes with PptxGenJS](https://singerla.github.io/pptx-automizer/generation) · [Hyperlinks](https://singerla.github.io/pptx-automizer/hyperlinks) · [Slide management](https://singerla.github.io/pptx-automizer/slide-management)
- [Helper utilities](https://singerla.github.io/pptx-automizer/helpers) · [Output](https://singerla.github.io/pptx-automizer/output) · [Requirements and limitations](https://singerla.github.io/pptx-automizer/limitations) · [Troubleshooting](https://singerla.github.io/pptx-automizer/troubleshooting)
- [API reference](https://singerla.github.io/pptx-automizer/api)

AI assistants consuming this library should read [AI-INSTRUCTOR.md](https://github.com/singerla/pptx-automizer/blob/main/AI-INSTRUCTOR.md), a condensed, self-contained guide that is also shipped in the npm package.

## Installation

```
$ yarn add pptx-automizer
```

or

```
$ npm install pptx-automizer
```

To explore the library from a clone instead, see [Getting started](https://singerla.github.io/pptx-automizer/getting-started).

## Example

```ts
import Automizer, { modify } from 'pptx-automizer';

const automizer = new Automizer({
  // where your template pptx files are coming from:
  templateDir: 'my/pptx/templates',
  // where to write the final pptx output files:
  outputDir: 'my/pptx/output',
  // truncate the root presentation and start with zero slides:
  removeExistingSlides: true,
});

const pres = automizer
  .loadRoot('RootTemplate.pptx')
  .load('SlideWithGraph.pptx', 'graph')
  .load('SlideWithShapes.pptx', 'shapes');

pres
  .addSlide('graph', 1)
  .addSlide('shapes', 1, (slide) => {
    slide.modifyElement('Drum', [modify.setText('Hello world!')]);
  });

// Nothing has happened yet: all calls above only queue work.
// write() executes the whole queue and produces the output file.
const summary = await pres.write('myPresentation.pptx');
console.log(summary);
```

The single most important concept: calls like `addSlide()` and `modifyElement()` only **queue** work — including every modification callback you pass. Nothing is applied until `await pres.write(...)` (or `.stream()` / `.getJSZip()`) executes the whole queue in order. Read more in [Concepts](https://singerla.github.io/pptx-automizer/concepts).

Take a look into the [`__tests__`-directory](https://github.com/singerla/pptx-automizer/blob/main/__tests__) to see a lot of examples for several use cases. You will also find the example .pptx-files there.

## Ecosystem and Support

This project is accompanied by [automizer-data](https://github.com/singerla/automizer-data). You can use `automizer-data` to import, browse and transform .xlsx- or .sav-data into perfectly fitting graph or table data.

Thanks to all contributors! You are always welcome to share code, tipps and ideas. We appreciate all levels of expertise and encourage everyone to get involved. Whether you're a seasoned pro or just starting out, your contributions are invaluable. [Get started](https://github.com/singerla/pptx-automizer/issues/new)

# Special Thanks

This project was inspired by:

- [PptxGenJS](https://github.com/gitbrent/PptxGenJS)
- [officegen](https://github.com/Ziv-Barber/officegen)
- [node-pptx](https://github.com/heavysixer/node-pptx)
- [docxtemplater](https://github.com/open-xml-templating/docxtemplater)
