# pptx-automizer
This is a pptx generator for Node.js based on templates. It can read pptx files and insert selected slides or single slide elements into another presentation. *pptx-automizer* will not write files from scratch, but edit and merge existing pptx files. Template slides are styled within PowerPoint and will be merged into the output presentation. Most of the content can be modified by using callbacks with [xmldom](https://github.com/xmldom/xmldom).

*pptx-automizer* will fit best to users who try to maintain their own library of pptx template files. This is perfect to anyone who uses complex and well-styled customized layouts. Any existing slide and even a single element can be a data driven template for output pptx files.

## Requirements
This generator can only be used on the server-side and requires a [Node.js](https://nodejs.org/en/download/package-manager/) environment.

## Limitations
Please note that this project is *work in progress*. At the moment, you might encounter difficulties for special shape types that require further relations (e.g. links will not work properly). Although, most shape types are already supported, such as connection shapes, tables or charts. You are welcome to [report any issue](https://github.com/singerla/pptx-automizer/issues/new).

All testing focuses on PowerPoint 2019 pptx file format.

## Install
There are basically two ways to use *pptx-automizer*.

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
If you are working on an existing project, you can add *pptx-automizer* to it using npm or yarn. Run
```
$ yarn add pptx-automizer
```
or
```
$ npm install pptx-automizer
```
in the root folder of your project. This will download and install the most recent version into your existing project.

## Example
```js
import Automizer, { modify } from "pptx-automizer"

// First, let's set some preferences
const automizer = new Automizer({
  templateDir: `my/pptx/templates`,
  outputDir: `my/pptx/output`
})

// Now we can start and load a pptx template.
// Each addSlide will append to any existing slide in RootTemplate.pptx.
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

// You can also select and import a single element from a template slide. 
// The desired shape will be identified by its name from slide-xml's 
// 'p:cNvPr'-element.
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

// It is possible to modify an existing element on a newly added slide.
pres.addSlide('shapes', 2, (slide) => {
  slide.modifyElement('Drum', [
    // You can use some of the builtin modifiers to edit a shape's xml:
    modify.setPosition({x: 1000000, h:5000000, w:5000000}),
    // Log your target xml into the console:
    modify.dump
  ])
})

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
    // Please notice: If your template has more data than your data
    // object, automizer will remove these nodes.
  ])
})

// Finally, we want to write the output file.
pres.write(`myPresentation.pptx`).then(summary => {
  console.log(summary)
})
```

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
