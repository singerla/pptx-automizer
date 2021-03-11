import Automizer from "../src/index"
import FileHelper from "./helper/file"


const automizer = new Automizer({
  templateDir: `${__dirname}/../__tests__/pptx-templates`,
  outputDir: `${__dirname}/../__tests__/pptx-output`
})

let pres = automizer.importRootTemplate(`RootTemplateWithCharts.pptx`)
  .importTemplate(`SlideWithImage.pptx`, 'image')

pres.addSlide('image', 3)
pres.addSlide('image', 3)

pres.write(`myPresentation.pptx`).then(result => {
  console.log(result)
})