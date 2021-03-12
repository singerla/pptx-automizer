import Automizer from "./index"
import Slide from "./slide"
import { setSolidFill, setText } from "./helper/modify"

const automizer = new Automizer({
  templateDir: `${__dirname}/../__tests__/pptx-templates`,
  outputDir: `${__dirname}/../__tests__/pptx-output`
})

let pres = automizer.loadRoot(`RootTemplateWithCharts.pptx`)
  .load(`SlideWithImage.pptx`, 'image')
  .load(`SlideWithShapes.pptx`, 'shapes')
  .load(`SlideWithCharts.pptx`, 'charts')

pres
  .addSlide('image', 1, (slide: Slide) => {
    slide.addElement('shapes', 2, 'Cloud', [ setSolidFill, setText('my cloudy thoughts')] )
    slide.addElement('shapes', 2, 'Arrow', setText('my text'))
    slide.addElement('shapes', 2, 'Drum')
  })
  .addSlide('image', 1, (slide: Slide) => {
    slide.addElement('shapes', 2, 'Cloud', [ setSolidFill, setText('my cloudy thoughts 2')] )
    slide.addElement('shapes', 2, 'Arrow', setText('my text'))
    slide.addElement('shapes', 2, 'Drum')
  })
  .addSlide('charts', 1)
  .write(`myPresentation.pptx`).then(result => {
    console.info(result)
  }).catch(error => {
    console.error(error)
  })
