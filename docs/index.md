---
slug: /
title: Introduction
sidebar_label: Introduction
description: What pptx-automizer is — a template-based .pptx generator for Node.js that modifies existing PowerPoint templates instead of building slides from scratch.
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

`pptx-automizer` is particularly well-suited for users who aim to manage their
own library of .pptx template files, making it an ideal choice for those who
work with intricate, well-designed customized layouts. With this tool, any
existing slide or even a single element can serve as a data-driven template for
generating output .pptx files.

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
`.getJSZip()`) executes the whole queue in order. Read more in
[Concepts](./concepts.md).

## Where to go next

- [Getting started](./getting-started.md) — installation and a fully commented example.
- [Concepts](./concepts.md) — the mental model: deferred execution, the template/root model, 1-based numbering.
- [Selecting slides and shapes](./selectors.md) — how to address the things you want to change.
- Feature guides: [Text](./text.md), [Tables](./tables.md), [Charts](./charts.md), [Images](./images.md), [Masters & Layouts](./masters-layouts.md), [Shape generation](./generation.md), [Hyperlinks](./hyperlinks.md), [Slide management](./slide-management.md).
- [Helper Utilities](./helpers.md) — shape, cleanup, unit and XML helpers.
- [Output](./output.md) — write, stream, JSZip access, status tracking.
- [Requirements and Limitations](./limitations.md) and [Troubleshooting](./troubleshooting.md).
- [Testing and Validation Tools](./testing.md) — the test suite, the OOXML schema validator and visual regression.
- [API reference](./api/index.md) — generated from the TypeScript sources.
- [AI-INSTRUCTOR.md](https://github.com/singerla/pptx-automizer/blob/main/AI-INSTRUCTOR.md)
  — a condensed, self-contained guide for AI assistants consuming this
  library (also shipped in the npm package).

## Ecosystem

This project is accompanied by [automizer-data](https://github.com/singerla/automizer-data). You can use `automizer-data` to import, browse and transform .xlsx- or .sav-data into perfectly fitting graph or table data.

Thanks to all contributors! You are always welcome to share code, tipps and ideas. We appreciate all levels of expertise and encourage everyone to get involved. [Get started](https://github.com/singerla/pptx-automizer/issues/new)
