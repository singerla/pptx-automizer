# pptx-automizer
This is a pptx generator for Node.js based on templates. It can read pptx files and insert selected slides or single slide elements into another presentation. *pptx-automizer* will not write files from scratch, but edit and merge existing pptx files. Template slides are styled with PowerPoint and will be merged into the output presentation. Most of the content can be modified by using callbacks based on [xmldom](https://github.com/xmldom/xmldom).

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
// Each addSlide will append to any existing slide in RootTemplate.pptx.
let pres = automizer.loadRoot(`RootTemplate.pptx`)
  // We want to make some more files available and give them a handy label.
  .load(`SlideWithShapes.pptx`, 'shapes')
  .load(`SlideWithGraph.pptx`, 'graph')
  // Skipping the second argument will not set a label.
  .load(`SlideWithImages.pptx`)

// addSlide takes two arguments: The first will specify the source presentation's
// label to get the template from, the second will set the slide number to require.
pres.addSlide('graph', 1)
  .addSlide('shapes', 1)
  .addSlide(`SlideWithImages.pptx`, 2)

// You can also select and import a single element from a template slide. The desired
// shape will be identified by its name from slide-xml's 'p:cNvPr'-element.
pres.addSlide('image', 1, (slide) => {
  // Pass the template name, the slide number, the element's name and (optionally)
  // a callback function to directly modify the child nodes of <p:sp>
  slide.addElement('shapes', 2, 'Arrow', (element) => {
    element.getElementsByTagName('a:t')[0]
      .firstChild
      .data = 'Custom content'
  })
})

// Finally, we want to write the output file.
pres.write(`myPresentation.pptx`).then(summary => {
  console.log(summary)
})
```

### Testing
You can run tests using these commands:
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