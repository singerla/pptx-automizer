---
title: Getting started
description: Install pptx-automizer and build your first presentation.
---

There are basically two ways to use `pptx-automizer`.

## Installation

### As a Package

If you are working on an existing project, you can add `pptx-automizer` to it using npm or yarn. Run

```
$ yarn add pptx-automizer
```

or

```
$ npm install pptx-automizer
```

in the root folder of your project. This will download and install the most recent version into your existing project.

### As a Cloned Repository

If you want to see how it works and you like to run own tests, you should clone the [repository](https://github.com/singerla/pptx-automizer) and install the dependencies:

```
$ git clone git@github.com:singerla/pptx-automizer.git
$ cd pptx-automizer
$ yarn install
```

You can now run

```
$ yarn dev
```

and see the most recent feature from `src/dev.ts`. Every time you change & save this file, you will see new console output and a pptx file in the destination folder. Take a look into the [`__tests__`-directory](https://github.com/singerla/pptx-automizer/blob/main/__tests__) to see a lot of examples for several use cases!

## Basic Example

This is a basic example on how to use `pptx-automizer` in your code. Before you dive in, make sure you know about [deferred execution](./concepts.md) — nothing is applied until `write()` is called. Most of the examples in these docs make use of [these template files](https://github.com/singerla/pptx-automizer/blob/main/__tests__/pptx-templates).

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

See [Output](./output.md) for the other output formats — a `ReadableStream` or the underlying JSZip instance.

## More examples

Take a look into the [`__tests__`-directory](https://github.com/singerla/pptx-automizer/blob/main/__tests__) to see a lot of examples for several use cases, e.g.:

- [Style chart series or datapoints](https://github.com/singerla/pptx-automizer/blob/main/__tests__/modify-existing-chart-styled.test.ts)
- [Use tags inside text to replace contents](https://github.com/singerla/pptx-automizer/blob/main/__tests__/replace-tagged-text.test.ts)
- [Modify vertical line charts](https://github.com/singerla/pptx-automizer/blob/main/__tests__/modify-chart-vertical-lines.test.ts)
- [Set table cell and border styles](https://github.com/singerla/pptx-automizer/blob/main/__tests__/modify-existing-table.test.ts)
- [Update chart plot area coordinates](https://github.com/singerla/pptx-automizer/blob/main/__tests__/modify-chart-plot-area.test.ts)
- [Update chart legend](https://github.com/singerla/pptx-automizer/blob/main/__tests__/modify-chart-legend.test.ts)
