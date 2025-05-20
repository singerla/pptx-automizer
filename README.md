# pptx-automizer: A Powerful .pptx Modifier for Node.js

`pptx-automizer` is a Node.js-based PowerPoint (.pptx) generator that automates the manipulation of existing .pptx files. With `pptx-automizer`, you can import your library of .pptx templates, merge templates, and customize slide content. `pptx-automizer` will edit and merge existing pptx files. You can style template slides within PowerPoint, and these templates will be seamlessly integrated into the output presentation. Most of the content can be modified by using callbacks with [xmldom](https://github.com/xmldom/xmldom).

If you require to create elements from scratch, `pptx-automizer` wraps around [PptxGenJS](https://github.com/gitbrent/PptxGenJS). Use the powerful syntax of `PptxGenJS` to add dynamic content to your existing .pptx template files. See an example on [how to add a chart from scratch](https://github.com/singerla/pptx-automizer/blob/main/__tests__/generate-pptxgenjs-charts.test.ts).

`pptx-automizer` is particularly well-suited for users who aim to manage their own library of .pptx template files, making it an ideal choice for those who work with intricate, well-designed customized layouts. With this tool, any existing slide or even a single element can serve as a data-driven template for generating output .pptx files.

This project is accompanied by [automizer-data](https://github.com/singerla/automizer-data). You can use `automizer-data` to import, browse and transform .xlsx- or .sav-data into perfectly fitting graph or table data.

Thanks to all contributors! You are always welcome to share code, tipps and ideas. We appreciate all levels of expertise and encourage everyone to get involved. Whether you're a seasoned pro or just starting out, your contributions are invaluable. [Get started](https://github.com/singerla/pptx-automizer/issues/new)

If you require commercial support for complex .pptx automation, you can explore [ensemblio.com](https://ensemblio.com). Ensemblio is a web application that leverages `pptx-automizer` and `automizer-data` to provide an accessible and convenient solution for automating .pptx files. Engaging with Ensemblio is likely to enhance and further develop this library.

## Table of contents

<!-- TOC -->

- [Requirements and Limitations](#requirements-and-limitations)
  - [Shape Types](#shape-types)
  - [Chart Types](#chart-types)
  - [Animations](#animations)
  - [Slide Masters and -Layouts](#slide-masters-and--layouts)
  - [Direct Manipulation of Elements](#direct-manipulation-of-elements)
  - [PowerPoint Version](#powerpoint-version)
- [Installation](#installation)
  - [As a Cloned Repository](#as-a-cloned-repository)
  - [As a Package](#as-a-package)
- [Usage](#usage)

  - [Basic Example](#basic-example)
  - [How to Select Slides Shapes](#how-to-select-slides-shapes)
    - [Select slide by number and shape by name](#select-slide-by-number-and-shape-by-name)
    - [Select slides by creationId](#select-slides-by-creationid)
  - [Find and Modify Shapes](#find-and-modify-shapes)
  - [Modify Text](#modify-text)
  - [Modify Images](#modify-images)
  - [Modify Tables](#modify-tables)
  - [Modify Charts](#modify-charts)
  - [Modify Extended Charts](#modify-extended-charts)
  - [Generate shapes with PptxGenJs](#generate-shapes-with-pptxgenjs)
  - [Remove elements from a slide](#remove-elements-from-a-slide)
  - [Hyperlink Management](#hyperlink-management)

- [Tipps and Tricks](#tipps-and-tricks)
  - [Loop through the slides of a presentation](#loop-through-the-slides-of-a-presentation)
  - [Quickly get all slide numbers of a template](#quickly-get-all-slide-numbers-of-a-template)
  - [Find all text elements on a slide](#find-all-text-elements-on-a-slide)
  - [Sort output slides](#sort-output-slides)
  - [Import and modify slide Masters](#import-and-modify-slide-masters)
  - [Track status of automation process](#track-status-of-automation-process)
  - [More examples](#more-examples)
  - [Troubleshooting](#troubleshooting)
  - [Testing](#testing)
- [Special Thanks](#special-thanks)

<!-- TOC -->

# Requirements and Limitations

This generator can only be used on the server-side and requires a [Node.js](https://nodejs.org/en/download/package-manager/) environment.

## Shape Types

At the moment, you might encounter difficulties with special shape types that require additional relations (e.g., hyperlinks, video and audio may not work correctly). However, most shape types, including connection shapes, tables, and charts, are already supported. If you encounter any issues, please feel free to [report any issue](https://github.com/singerla/pptx-automizer/issues/new).

## Chart Types

Extended chart types, like waterfall or map charts, are basically supported. You might need additional modifiers to handle extended properties, which are not implemented yet. Please help to improve `pptx-automizer` and [report](https://github.com/singerla/pptx-automizer/issues/new) issues regarding extended charts.

## Animations

Animations are currently out of scope of this library. You might get errors on opening an output .pptx when there are added or removed shapes. This is because `pptx-automizer` doesn't synchronize `id`-attributes of animations with the existing shapes on a slide.

## Slide Masters and -Layouts

`pptx-automizer` supports importing slide masters and their associated slide layouts into the output presentation. It is important to note that you cannot add, modify, or remove individual slideLayouts directly. However, you have the flexibility to modify the underlying slideMaster, which can serve as a workaround for certain changes.

Please be aware that importing slideLayouts containing complex contents, such as charts and images, is currently not supported. For instance, if a slideLayout includes an icon that is not present on the slideMaster, this icon will break when the slideMaster is auto-imported into an output presentation. To avoid this issue, ensure that all images and charts are placed exclusively on a slideMaster and not on a slideLayout.

## Direct Manipulation of Elements

It is also important to know that `pptx-automizer` is currently limited to _adding_ things to the output presentation. If you require the ability to, for instance, modify a specific element on a slide within an existing presentation and leave the rest untouched, you will need to include all the other slides in the process. Find some workarounds [below](#loop-through-the-slides-of-a-presentation).

## PowerPoint Version

All testing focuses on PowerPoint 2019 .pptx file format.

# Installation

There are basically two ways to use `pptx-automizer`.

## As a Cloned Repository

If you want to see how it works and you like to run own tests, you should clone this repository and install the dependencies:

```
$ git clone git@github.com:singerla/pptx-automizer.git
$ cd pptx-automizer
$ yarn install
```

You can now run

```
$ yarn dev
```

and see the most recent feature from `src/dev.ts`. Every time you change & save this file, you will see new console output and a pptx file in the destination folder. Take a look into `__tests__`-directory to see a lot of examples for several use cases!

## As a Package

If you are working on an existing project, you can add `pptx-automizer` to it using npm or yarn. Run

```
$ yarn add pptx-automizer
```

or

```
$ npm install pptx-automizer
```

in the root folder of your project. This will download and install the most recent version into your existing project.

# Usage

Take a look into [**tests**-directory](https://github.com/singerla/pptx-automizer/blob/main/__tests__) to see a lot of examples for several use cases. You will also find example .pptx-files there. Most of the examples shown below make use of [those files](https://github.com/singerla/pptx-automizer/blob/main/__tests__/pptx-templates).

## Basic Example

This is a basic example on how to use `pptx-automizer` in your code:

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
  // Powerpoint's creationIds instead of slide numbers
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

  // Use 1 to show warnings or 2 for detailed information
  // 0 disables logging
  verbosity: 1,

  // Remove all unused placeholders to prevent unwanted overlays:
  cleanupPlaceholders: false,
  
  // Use a customized version of pptxGenJs if required:
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
const mySlides = presInfo.slidesByTemplate('shapes');
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

## How to Select Slides Shapes

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
import { FindElementSelector } from './types/types';

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

## Find and Modify Shapes

There are basically to ways to access a target shape on a slide:

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

## Modify Text

You can select and import generic shapes from any loaded template. It is possible to update the containing text in several ways:

```ts
import { ModifyTextHelper } from 'pptx-automizer';

pres.addSlide('SlideWithImages.pptx', 1, (slide) => {
  // You can directly modify the child nodes of <p:sp>
  slide.addElement('shapes', 2, 'Arrow', (element) => {
    element.getElementsByTagName('a:t').item(0).firstChild.data =
      'Custom content';
  });

  // You might prefer a built-in function to set text:
  slide.addElement('shapes', 2, 'Arrow', [
    ModifyTextHelper.setText('This is my text'),
  ]);
});
```

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

Find out more about text replacement:

- [Replace and style by tags](https://github.com/singerla/pptx-automizer/blob/main/__tests__/replace-tagged-text.test.ts)
- [Modify text elements using getAllTextElementIds](https://github.com/singerla/pptx-automizer/blob/main/__tests__/get-all-text-element-ids.test.ts)

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

Find more examples on image manipulation:

- [Add external image](https://github.com/singerla/pptx-automizer/blob/main/__tests__/add-external-image.test.ts)
- [Modify duotone color overlay for images](https://github.com/singerla/pptx-automizer/blob/main/__tests__/modify-image-duotone.test.ts)
- [Swap image source on a slide master](https://github.com/singerla/pptx-automizer/blob/main/__tests__/modify-master-external-image.test.ts)

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

Find out more about formatting cells:

- [Modify and style table cells](https://github.com/singerla/pptx-automizer/blob/main/__tests__/modify-existing-table.test.ts)
- [Insert data into table with empty cells](https://github.com/singerla/pptx-automizer/blob/main/__tests__/modify-existing-table-create-text.test.ts)

## Modify Charts

All data and styles of a chart can be modified. Please note that if your template contains more data than your data object, Automizer will remove these extra nodes. Conversely, if you provide more data, new nodes will be cloned from the first existing one in the template.

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

## Modify Extended Charts

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

## Generate shapes with PptxGenJs

This library wraps around the [PptxGenJS](https://github.com/gitbrent/PptxGenJS) to generate shapes from scratch. It is possible to use the `pptxGenJS` wrapper to generate shapes on a slide.

Here's an example of how to use `pptxGenJS` to add a text shape to a slide:
```ts
pres.addSlide('empty', 1, (slide) => {
  // Use pptxgenjs to add text from scratch:
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
  // Use pptxgenjs to add generated contents from scratch:
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

You can use the following functions to generate shapes with `pptxGenJS`:
* addChart
* addImage
* addShape
* addTable
* addText


## Remove elements from a slide

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

## Hyperlink Management

PowerPoint presentations often use hyperlinks to connect to external websites or internal slides. The `pptx-automizer` provides simple and powerful functions to manage hyperlinks in your presentations.

### Add Hyperlinks to existing shapes

You can add hyperlinks to template text shapes using the `addHyperlink` helper function. The function accepts either a URL string for external links or a slide number for internal slide links:

```ts
// Add an external hyperlink
slide.modifyElement('TextShape', modify.addHyperlink('https://example.com'));

// Add an internal slide link (to slide 3)
slide.modifyElement('TextShape', (element, relation) => {
  modify.addHyperlink(3)(element, relation);
});
```

The `addHyperlink` function will automatically detect whether the target is an external URL or an internal slide number and set up the appropriate relationship type and attributes.

### Create a new hyperlinked text shape with pptxGenJS

It is also possible to create a new hyperlink from scratch with the `pptxGenJS` wrapper. This is useful if you want to add hyperlinks to shapes that are not part of the template.

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

# Tipps and Tricks

## Loop through the slides of a presentation

If you would like to modify elements in a single .pptx file, it is important to know that `pptx-automizer` is not able to directly "jump" to a shape to modify it.

This is how it works internally:

- Load a root template to append slides to
- (Probably) load root template again to modify slides
- Load other templates
- Append a loaded slide to (probably truncated) root template
- Modify the recently added slide
- Write root template and appended slides as output presentation.

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
    // Defining a "name" as second params makes it a little easier
    .load(`SlideWithShapes.pptx`, 'myTemplate');

  // This is brandnew: get useful information about loaded templates:
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

  // We can loop all slides an apply the callbacks if defined
  mySlides.forEach((mySlide) => {
    pres.addSlide('myTemplate', mySlide.number, getCallbacks(mySlide.number));
  });

  // This will result to an output presentation containing all slides of "SlideWithShapes.pptx"
  pres.write(`myOutputPresentation.pptx`).then((summary) => {
    console.log(summary);
  });
};

run().catch((error) => {
  console.error(error);
});
```

## Quickly get all slide numbers of a template

When calling `pres.getInfo()`, it will gather information about all elements on all slides of all templates. In case you just want to loop through all slides of a certain template, you can use this shortcut:

```ts
const slideNumbers = await pres
  .getTemplate('myTemplate.pptx')
  .getAllSlideNumbers();

for (const slideNumber of slideNumbers) {
  // do the thing
}
```

## Find all text elements on a slide

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

## Sort output slides

There are three ways to arrange slides in an output presentation.

1. By default, all slides will be appended to the existing slides in your root template. The order of `addSlide`-calls will define slide sortation in output presentation.

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
  // We load this twice to make it available for sorting slide
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
import ModifyPresentationHelper from './helper/modify-presentation-helper';

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

## Import and modify slide Masters

You can import, modify and use one or more slideMasters and the related slideLayouts.
It is only supported to add and modify shapes on the underlying slideMaster, you cannot modify something on a slideLayout. This means, each modification on a slideMaster will appear on all related slideLayouts.

To specify the target index of the required slide master to import, you need to count slideMasters in your _template_ presentation.
To specify another slideLayout for an added output slide, you need to count slideLayouts in your _output_ presentation

To add and modify shapes on a slide master, please take a look at [Add and modify shapes](https://github.com/singerla/pptx-automizer#add-and-modify-shapes).

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

## Track status of automation process

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

## More examples

Take a look into [**tests**-directory](https://github.com/singerla/pptx-automizer/blob/main/__tests__) to see a lot of examples for several use cases, e.g.:

- [Style chart series or datapoints](https://github.com/singerla/pptx-automizer/blob/main/__tests__/modify-existing-chart-styled.test.ts)
- [Use tags inside text to replace contents](https://github.com/singerla/pptx-automizer/blob/main/__tests__/replace-tagged-text.test.ts)
- [Modify vertical line charts](https://github.com/singerla/pptx-automizer/blob/main/__tests__/modify-chart-vertical-lines.test.ts)
- [Set table cell and border styles](https://github.com/singerla/pptx-automizer/blob/main/__tests__/modify-existing-table.test.ts)
- [Update chart plot area coordinates](https://github.com/singerla/pptx-automizer/blob/main/__tests__/modify-chart-plot-area.test.ts)
- [Update chart legend](https://github.com/singerla/pptx-automizer/blob/main/__tests__/modify-chart-legend.test.ts)

## Troubleshooting

If you encounter problems when opening a `.pptx`-file modified by this library, you might worry about PowerPoint not giving any details about the error. It can be hard to find the cause, but there are some things you can check:

- **Broken relation**: There are still unsupported shape types and `pptx-automizer` wil not copy required relations of those. You can inflate `.pptx`-output and check `ppt/slides/_rels/slide[#].xml.rels`-files to find possible missing files.
- **Unsupported media**: You can also take a look at the `ppt/media`-directory of an inflated `.pptx`-file. If you discover any unusual file formats, remove or replace the files by one of the [known types](https://github.com/singerla/pptx-automizer/blob/main/src/enums/content-type-map.ts).
- **Broken animation**: Pay attention to modified/removed shapes which are part of an animation. In case of doubt, (temporarily) remove all animations from your template. (see [#78](https://github.com/singerla/pptx-automizer/issues/78))
- **Proprietary/Binary contents** (e.g. ThinkCell): Walk through all slides, slideMasters and slideLayouts and seek for hidden Objects. Hit `ALT+F10` to toggle the sidebar.
- **Chart datasheet won't open** If you encounter an error message on opening a chart's datasheet, please make sure that the data table (blue bordered rectangle in worksheet view) of your template starts at cell A:1. If not, open worksheet in Excel mode and edit the table size in the table draft tab.

If none of these could help, please don't hesitate to [talk about it](https://github.com/singerla/pptx-automizer/issues/new).

## Testing

You can run all unit tests using these commands:

```
yarn test
yarn test-coverage
```

# Special Thanks

This project was inspired by:

- [PptxGenJS](https://github.com/gitbrent/PptxGenJS)
- [officegen](https://github.com/Ziv-Barber/officegen)
- [node-pptx](https://github.com/heavysixer/node-pptx)
- [docxtemplater](https://github.com/open-xml-templating/docxtemplater)
