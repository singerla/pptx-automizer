# pptx-automizer
This is a pptx generator for Node.js based on templates. It can read pptx files and insert selected slides into another presentation. Compared to other pptx libraries (such as [PptxGenJS](https://github.com/gitbrent/PptxGenJS), [officegen](https://github.com/Ziv-Barber/officegen) or [node-pptx](https://github.com/heavysixer/node-pptx)), *pptx-automizer* will not write files from scratch, but edit and merge existing pptx files. Template slides are styled with PowerPoint and will be merged 1:1 into the output presentation.

It is also possible to choose a certain element by its name from a slide and insert it to another template slide. Take a look at <code>slide.addElement()</code> from the example below.

## Requirements
This generator can only be used on the server-side and requires a [Node.js](https://nodejs.org/en/download/package-manager/) environment.

## Limitations
Please note that this project is *work in progress*. At the moment, you might encounter difficulties for special shape types that require internal relations.
Although, most shape types are already supported, such as connection shapes or charts.

Importing a single element is limited to shapes that also do not require further relations.

## Install
You can add this package to your own project using npm or yarn:
```
yarn add pptx-automizer
```
or
```
npm install pptx-automizer
```

## Example
```js
import Automizer from "pptx-automizer"

// First, let's set some preferences
const automizer = new Automizer({
  templateDir: `my/pptx/templates`,
  outputDir: `my/pptx/output`
})

// Now we can start and load a pptx template.
// Skipping the second argument will set the root template.
// Each addSlide will append to any existing slide in RootTemplate.pptx.
let pres = automizer.loadRoot(`RootTemplate.pptx`)
  // We want to make some more files available and give them a handy label.
  .load(`SlideWithShapes.pptx`, 'shapes')
  .load(`SlideWithGraph.pptx`, 'graph')
  .load(`SlideWithImages.pptx`, 'images')

// addSlide takes two arguments: The first will specify the source presentation's
// label to take the template from, the second will set the slide number to require.
pres.addSlide('graph', 1)
  .addSlide('shapes', 1)
  .addSlide('images', 2)

// You can also select and import a single element from a template slide. The desired
// shape will be identified by its name from slide-xml's 'p:cNvPr'-element.
pres.addSlide('image', 1, (slide) => {
  // Pass the template name, the slide number, the element's name and (optionally)
  // a callback function to directly modify the child nodes of <p:sp>
  slide.addElement('shapes', 2, 'Arrow', (element) => {
    element.getElementsByTagName('a:t')[0]
      .firstChild
      .data = text
  })
})

// Finally, we want to write the output file.
pres.write(`myPresentation.pptx`).then(summary => {
  console.log(summary)
})
```
