---
slug: /
title: Introduction
sidebar_label: Introduction
---

# pptx-automizer

`pptx-automizer` is a Node.js-based PowerPoint (.pptx) generator that automates
the manipulation of existing .pptx files. Import your library of .pptx
templates, merge templates, and customize slide content. You can style template
slides within PowerPoint, and these templates will be seamlessly integrated
into the output presentation. Most of the content can be modified by using
callbacks with [xmldom](https://github.com/xmldom/xmldom).

If you need to create elements from scratch, `pptx-automizer` wraps around
[PptxGenJS](https://github.com/gitbrent/PptxGenJS) to add dynamic content to
your existing .pptx template files.

## Installation

```bash
yarn add pptx-automizer
# or
npm install pptx-automizer
```

## Quick example

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

The single most important concept: calls like `addSlide()` and
`modifyElement()` only **queue** work — including every modification callback
you pass. Nothing is applied until `await pres.write(...)` (or `.stream()` /
`.getJSZip()`) executes the whole queue in order.

## Where to go next

The docs site is being migrated section by section. Until every page has
landed here, the complete feature documentation lives in the
[README on GitHub](https://github.com/singerla/pptx-automizer#readme).

- [API reference](./api/index.md) — generated from the TypeScript sources.
- [AI-INSTRUCTOR.md](https://github.com/singerla/pptx-automizer/blob/main/AI-INSTRUCTOR.md)
  — a condensed, self-contained guide for AI assistants consuming this
  library (also shipped in the npm package).
