# pptx-automizer
This is a pptx generator for Node.js based on templates. It can read pptx files and insert selected slides into another presentation. Compared to other pptx libraries (such as [PptxGenJS](https://github.com/gitbrent/PptxGenJS), [officegen](https://github.com/Ziv-Barber/officegen) or [node-pptx](https://github.com/heavysixer/node-pptx)), *pptx-automizer* will not write files from scratch, but edit and merge existing pptx files. Any type of shape can be styled within PowerPoint itself and will be copied 1:1 into the output presentation.

## Requirements
This generator can only be used on the server-side and requires a [Node.js](https://nodejs.org/en/download/package-manager/) environment.

## Limitations
Please note that this project is *work in progress*. At the moment, it is *not* possible to handle slides containing:
* images
* links
* notes

Although, most other shape types are already supported, such as connection shapes or charts.

## Install
You can add this package to your own project using npm:
```
yarn add pptx-generator
```
or
```
npm install pptx-generator
```

## Example
```js
import Automizer from "pptx-automizer"
const automizer = new Automizer

// First, lets set some preferences
const templateFolder = `my/pptx/templates/`
const outputFolder = `my/pptx/output/`

// Let's start and import a root template. All slides will be appended to 
// any existing slide in RootTemplate.pptx
let pres = automizer.importRootTemplate(`${templateFolder}RootTemplate.pptx`)
  // We want to make two files available and give them a handy label.
  .importTemplate(`${templateFolder}SlideWithShapes.pptx`, 'shapes')
  .importTemplate(`${templateFolder}SlideWithGraph.pptx`, 'graph')

// addSlide takes two arguments: The first will specify the source presentation
// where your template should come from, the second will set the slide number.
pres.addSlide('graph', 1)

// You can also loop through something and add slides in a batch.
for(let i=0; i<=10; i++) {
  pres.addSlide('shapes', 1)
}

// Finally, we want to write the output file.
pres.write(`${outputFolder}myPresentation.pptx`)
```
