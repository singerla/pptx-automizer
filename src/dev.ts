import Automizer from "../src/index"


const automizer = new Automizer({
  templateDir: `${__dirname}/../__tests__/pptx-templates`,
  outputDir: `${__dirname}/../__tests__/pptx-output`
})

let pres = automizer.importRootTemplate(`RootTemplateWithCharts.pptx`)
  .importTemplate(`SlideWithCharts.pptx`, 'charts')

pres.addSlide('charts', 1)

pres.write(`myPresentation.pptx`).then(result => {
  console.log(result)
})