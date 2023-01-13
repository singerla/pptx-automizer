# pptx-automizer
This is a pptx generator for Node.js based on templates. It can read pptx files and insert selected slides or single slide elements into another presentation. `pptx-automizer` will not write files from scratch, but edit and merge existing pptx files. Template slides are styled within PowerPoint and will be merged into the output presentation. Most of the content can be modified by using callbacks with [xmldom](https://github.com/xmldom/xmldom).

`pptx-automizer` will fit best to users who try to maintain their own library of pptx template files. This is perfect to anyone who uses complex and well-styled customized layouts. Any existing slide and even a single element can be a data driven template for output pptx files.

This project comes along with [automizer-data](https://github.com/singerla/automizer-data). You can use `automizer-data` to import, browse and transform XSLX-data into perfectly fitting graph or table data.

## Requirements
This generator can only be used on the server-side and requires a [Node.js](https://nodejs.org/en/download/package-manager/) environment.

## Limitations
### Shape types
Please note that this project is *work in progress*. At the moment, you might encounter difficulties for special shape types that require further relations (e.g. links will not work properly). Although, most shape types are already supported, such as connection shapes, tables or charts. You are welcome to [report any issue](https://github.com/singerla/pptx-automizer/issues/new).

### Chart types
Extended chart types, like waterfall or map charts, are basically supported. You might need additional modifiers to handle extended properties, which are not implemented yet. Please help to improve `pptx-automizer` and [report](https://github.com/singerla/pptx-automizer/issues/new) issues regarding extended charts.

### PowerPoint Version
All testing focuses on PowerPoint 2019 pptx file format.

### Slide Masters and -Layouts
It is currently not supported to import slide masters or slide layouts into the root presentation. You can append any slide in any order, but you need to assure all pptx files to have the same set of master slides. Imported slide masters will be matched by number, e.g. if your imported slide uses master #3, your root presentation also needs to have at least three master slides.


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
import Automizer from "pptx-automizer"
  
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
  
  // truncate root presentation and start with zero slides
  removeExistingSlides: true,
  
  // use a callback function to track pptx generation process.
  // statusTracker: myStatusTracker,
})

// Now we can start and load a pptx template.
// With removeExistingSlides set to 'false', each addSlide will append to 
// any existing slide in RootTemplate.pptx. Otherwise, we are going to start
// with a truncated root template.
let pres = automizer.loadRoot(`RootTemplate.pptx`)
  // We want to make some more files available and give them a handy label.
  .load(`SlideWithShapes.pptx`, 'shapes')
  .load(`SlideWithGraph.pptx`, 'graph')
  // Skipping the second argument will not set a label.
  .load(`SlideWithImages.pptx`)

// addSlide takes two arguments: The first will specify the source 
// presentation's label to get the template from, the second will set the 
// slide number to require.
pres.addSlide('graph', 1)
  .addSlide('shapes', 1)
  .addSlide(`SlideWithImages.pptx`, 2)

// Finally, we want to write the output file.
pres.write(`myPresentation.pptx`).then(summary => {
  console.log(summary)
})
```


## Modify shapes with built-in functions
It is possible to modify an existing element on a newly added slide.

```ts
import { modify } from "pptx-automizer"

pres.addSlide('shapes', 2, (slide) => {
  slide.modifyElement('Drum', [
    // You can use some of the builtin modifiers to edit a shape's xml:
    modify.setPosition({x: 1000000, h:5000000, w:5000000}),
    // Log your target xml into the console:
    modify.dump
  ])
})
```

## Add and modify shapes
You can also select and import a single element from a template slide. The desired shape will be identified by its name from slide-xml's `p:cNvPr`-element.

```ts
pres.addSlide('SlideWithImages.pptx', 1, (slide) => {
  // Pass the template name, the slide number, the element's name and 
  // (optionally) a callback function to directly modify the child nodes 
  // of <p:sp>
  slide.addElement('shapes', 2, 'Arrow', (element) => {
    element.getElementsByTagName('a:t')[0]
      .firstChild
      .data = 'Custom content'
  })
})
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
        { label: 'cat 2-1', values: [ 50, 50, 20 ] },
        { label: 'cat 2-2', values: [ 14, 50, 20 ] },
        { label: 'cat 2-3', values: [ 15, 50, 20 ] },
        { label: 'cat 2-4', values: [ 26, 50, 20 ] }
      ]
    })
  ])
})
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
})
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
})
```


## Sort output slides

There are three ways to arrange slides in an output presentation. 

1. By default, all slides will be appended to the existing slides in your root template. The order of `addSlide`-calls will define slide sortation in output presentation.

2. You can alternatively remove all existing slides by setting the `removeExistingSlides` flag to true. The first slide added with `addSlide` will be first slide in the output presentation. If you want to insert slides from root template, you need to load it a second time.

```ts
import Automizer from "pptx-automizer"

const automizer = new Automizer({
  templateDir: `my/pptx/templates`,
  outputDir: `my/pptx/output`,

  // truncate root presentation and start with zero slides
  removeExistingSlides: true,
})


let pres = automizer.loadRoot(`RootTemplate.pptx`)
  // We load this twice to make it available for sorting slide
  .load(`RootTemplate.pptx`, 'root')  
  .load(`SlideWithShapes.pptx`, 'shapes')
  .load(`SlideWithGraph.pptx`, 'graph')

pres.addSlide('root', 1)  // First slide will be taken from root
  .addSlide('graph', 1)
  .addSlide('shapes', 1)
  .addSlide('root', 3)    // Third slide from root will be appended
  .addSlide('root', 2)    // Second and third slide will switch position
})

pres.write(`mySortedPresentation.pptx`).then(summary => {
  console.log(summary)
})
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
pres.addSlide('charts', 1);
  .addSlide('charts', 2);
  .addSlide('images', 1);
  .addSlide('images', 2);
  
const order = [3, 2, 4, 1];
pres.modify(ModifyPresentationHelper.sortSlides(order));
```


## Track status of automation process

When creating large presentations, you might want to have some information about the current status. Use a custom status tracker:

```ts
import Automizer, { StatusTracker } from "pptx-automizer"

// If you want to track the steps of creation process,
// you can use a custom callback:
const myStatusTracker = (status: StatusTracker) => {
  console.log(status.info + ' (' + status.share + '%)');
};

const automizer = new Automizer({
  // ...
  statusTracker: myStatusTracker,
})
```

## More examples
Take a look into [__tests__-directory](https://github.com/singerla/pptx-automizer/blob/main/__tests__) to see a lot of examples for several use cases, e.g.:
* [Style chart series or datapoints](https://github.com/singerla/pptx-automizer/blob/main/__tests__/modify-existing-chart-styled.test.ts)
* [Use tags inside text to replace contents](https://github.com/singerla/pptx-automizer/blob/main/__tests__/replace-tagged-text.test.ts)
* [Modify vertical line charts](https://github.com/singerla/pptx-automizer/blob/main/__tests__/modify-chart-vertical-lines.test.ts)
* [Set table cell and border styles](https://github.com/singerla/pptx-automizer/blob/main/__tests__/modify-existing-table.test.ts)


### Testing
You can run all unit tests using these commands:
```
yarn test
yarn test-coverage
```

### Special Thanks
This project is deeply inspired by:

* [PptxGenJS](https://github.com/gitbrent/PptxGenJS)
* [officegen](https://github.com/Ziv-Barber/officegen)
* [node-pptx](https://github.com/heavysixer/node-pptx)
* [docxtemplater](https://github.com/open-xml-templating/docxtemplater)
