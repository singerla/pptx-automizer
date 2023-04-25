# pptx-automizer

This is a pptx generator for Node.js based on templates. It can read pptx files and insert selected slides or single slide elements into another presentation. `pptx-automizer` will not write files from scratch, but edit and merge existing pptx files. Template slides are styled within PowerPoint and will be merged into the output presentation. Most of the content can be modified by using callbacks with [xmldom](https://github.com/xmldom/xmldom).

`pptx-automizer` will fit best to users who try to maintain their own library of pptx template files. This is perfect to anyone who uses complex and well-styled customized layouts. Any existing slide and even a single element can be a data driven template for output pptx files.

This project comes along with [automizer-data](https://github.com/singerla/automizer-data). You can use `automizer-data` to import, browse and transform XSLX-data into perfectly fitting graph or table data.

## Requirements

This generator can only be used on the server-side and requires a [Node.js](https://nodejs.org/en/download/package-manager/) environment.

## Limitations

### Shape types

Please note that this project is _work in progress_. At the moment, you might encounter difficulties for special shape types that require further relations (e.g. links will not work properly). Although, most shape types are already supported, such as connection shapes, tables or charts. You are welcome to [report any issue](https://github.com/singerla/pptx-automizer/issues/new).

### Chart types

Extended chart types, like waterfall or map charts, are basically supported. You might need additional modifiers to handle extended properties, which are not implemented yet. Please help to improve `pptx-automizer` and [report](https://github.com/singerla/pptx-automizer/issues/new) issues regarding extended charts.

### PowerPoint Version

All testing focuses on PowerPoint 2019 pptx file format.

### Slide Masters and -Layouts

It is basically supported to import slide masters and related slide layouts into the root presentation, but you can only import a master together with its related layouts. Any appended slide can use any of the available layouts afterwards. It is currently not possible to add, modify or remove a single slideLayout, but you can modify the underlying slideMaster.

## Install

There are basically two ways to use `pptx-automizer`.

### As a cloned repository

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

### As a package

If you are working on an existing project, you can add `pptx-automizer` to it using npm or yarn. Run

```
$ yarn add pptx-automizer
```

or

```
$ npm install pptx-automizer
```

in the root folder of your project. This will download and install the most recent version into your existing project.

## General Example

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
  // Powerpoint's creationIds instead of slide-numbers
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
});

// Now we can start and load a pptx template.
// With removeExistingSlides set to 'false', each addSlide will append to
// any existing slide in RootTemplate.pptx. Otherwise, we are going to start
// with a truncated root template.
let pres = automizer
  .loadRoot(`RootTemplate.pptx`)
  // We want to make some more files available and give them a handy label.
  .load(`SlideWithShapes.pptx`, 'shapes')
  .load(`SlideWithGraph.pptx`, 'graph')
  // Skipping the second argument will not set a label.
  .load(`SlideWithImages.pptx`);

// addSlide takes two arguments: The first will specify the source
// presentation's label to get the template from, the second will set the
// slide number to require.
pres
  .addSlide('graph', 1)
  .addSlide('shapes', 1)
  .addSlide(`SlideWithImages.pptx`, 2);

// Finally, we want to write the output file.
pres.write(`myPresentation.pptx`).then((summary) => {
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

## Modify shapes with built-in functions

It is possible to modify an existing element on a newly added slide.

```ts
import { modify } from 'pptx-automizer';

pres.addSlide('shapes', 2, (slide) => {
  slide.modifyElement('Drum', [
    // You can use some of the builtin modifiers to edit a shape's xml:
    modify.setPosition({ x: 1000000, h: 5000000, w: 5000000 }),
    // Log your target xml into the console:
    modify.dump,
  ]);
});
```

## Add and modify shapes

You can also select and import a single element from a template slide. The desired shape will be identified by its name from slide-xml's `p:cNvPr`-element.

```ts
pres.addSlide('SlideWithImages.pptx', 1, (slide) => {
  // Pass the template name, the slide number, the element's name and
  // (optionally) a callback function to directly modify the child nodes
  // of <p:sp>
  slide.addElement('shapes', 2, 'Arrow', (element) => {
    element.getElementsByTagName('a:t')[0].firstChild.data = 'Custom content';
  });
});
```

## Modify charts

All data and styles can be modified. Please notice: If your template has more data than your data object, automizer will remove these nodes. The other way round, new nodes will be created from the existing ones in case you provide more data.

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

## Modify extended charts

If you need to modify extended chart types, such like waterfall or map charts, you need to use `modify.setExtendedChartData`.

```ts
// Add and modify a waterfall chart on slide.
pres.addSlide('charts', 2, (slide) => {
  slide.addElement('ChartWaterfall.pptx', 1, 'Waterfall 1', [
    modify.setExtendedChartData(<ChartData>{
      series: [{ label: 'series 1' }],
      categories: [
        { label: 'cat 2-1', values: [100] },
        { label: 'cat 2-2', values: [20] },
        { label: 'cat 2-3', values: [50] },
        { label: 'cat 2-4', values: [-40] },
        { label: 'cat 2-5', values: [130] },
        { label: 'cat 2-6', values: [-60] },
        { label: 'cat 2-7', values: [70] },
        { label: 'cat 2-8', values: [140] },
      ],
    }),
  ]);
});
```

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

## Sort output slides

There are three ways to arrange slides in an output presentation.

1. By default, all slides will be appended to the existing slides in your root template. The order of `addSlide`-calls will define slide sortation in output presentation.

2. You can alternatively remove all existing slides by setting the `removeExistingSlides` flag to true. The first slide added with `addSlide` will be first slide in the output presentation. If you want to insert slides from root template, you need to load it a second time.

```ts
import Automizer from "pptx-automizer";

const automizer = new Automizer({
  templateDir: `my/pptx/templates`,
  outputDir: `my/pptx/output`,

  // truncate root presentation and start with zero slides
  removeExistingSlides: true,
});


let pres = automizer.loadRoot(`RootTemplate.pptx`)
  // We load this twice to make it available for sorting slide
  .load(`RootTemplate.pptx`, 'root')
  .load(`SlideWithShapes.pptx`, 'shapes')
  .load(`SlideWithGraph.pptx`, 'graph')

pres.addSlide('root', 1)  // First slide will be taken from root
  .addSlide('graph', 1)
  .addSlide('shapes', 1)
  .addSlide('root', 3)    // Third slide from root will be appended
  .addSlide('root', 2);    // Second and third slide will switch position

pres.write(`mySortedPresentation.pptx`).then(summary => {
  console.log(summary)
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
pres.addSlide('charts', 1)
  .addSlide('charts', 2)
  .addSlide('images', 1)
  .addSlide('images', 2);

const order = [3, 2, 4, 1];
pres.modify(ModifyPresentationHelper.sortSlides(order));
```

## Import and modify slide Masters
You can import, modify and use one or more slideMasters and the related slideLayouts.
It is only supported to add and modify shapes on the underlying slideMaster, you cannot modify something on a slideLayout. This means, each modification on a slideMaster will appear on all related slideLayouts.

To specify the target index of the required slide master to import, you need to count slideMasters in your *template* presentation.
To specify another slideLayout for an added output slide, you need to count slideLayouts in your *output* presentation 

To add and modify shapes on a slide master, please take a look at [Add and modify shapes](#Add and modify shapes).

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

If you have trouble with messed up slideMasters, and if you don't worry about the impact on performance, you can try and set ``autoImportSlideMasters: true`` to always import all required files:

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

### Testing

You can run all unit tests using these commands:

```
yarn test
yarn test-coverage
```

### Special Thanks

This project is deeply inspired by:

- [PptxGenJS](https://github.com/gitbrent/PptxGenJS)
- [officegen](https://github.com/Ziv-Barber/officegen)
- [node-pptx](https://github.com/heavysixer/node-pptx)
- [docxtemplater](https://github.com/open-xml-templating/docxtemplater)
