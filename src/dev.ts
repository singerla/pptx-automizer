import Automizer from "./index"
import Slide from "./slide"
import { setSolidFill, setText } from "./helper/modify"

const automizer = new Automizer({
  templateDir: `${__dirname}/../__tests__/pptx-templates`,
  outputDir: `${__dirname}/../__tests__/pptx-output`
})

let pres = automizer.loadRoot(`RootTemplate.pptx`)
  .load(`SlideWithImage.pptx`, 'image')
  .load(`SlideWithShapes.pptx`, 'shapes')
  .load(`SlideWithCharts.pptx`, 'charts')

pres
  .addSlide('image', 1, (slide: Slide) => {
    slide.addElement('charts', 1, 'StackedBars')
    // slide.addElement('charts', 1, 'StackedBars')
  })
  // .addSlide('image', 1, (slide: Slide) => {
  //   // slide.addElement('charts', 1, 'StackedBars')
  //   // slide.addElement('charts', 1, 'StackedBars')
  // })
  .write(`myPresentation.pptx`).then(result => {
    console.info(result)
  }).catch(error => {
    console.error(error)
  })
