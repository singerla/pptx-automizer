import Automizer from "../src/index"
const automizer = new Automizer

// First, lets set some preferences
const templateFolder = `${__dirname}/../__tests__/pptx-templates/`
const outputFolder = `${__dirname}/../__tests__/pptx-output/`

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