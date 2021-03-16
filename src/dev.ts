import Automizer from "./index"
import Slide from "./slide"
import { setSolidFill, setText } from "./helper/modify"

const automizer = new Automizer({
  templateDir: `${__dirname}/../__tests__/pptx-templates`,
  outputDir: `${__dirname}/../__tests__/pptx-output`
})

let pres = automizer.loadRoot(`RootTemplateWithCharts.pptx`)
  .load(`SlideWithImages.pptx`, 'images')
  .load(`SlideWithCharts.pptx`, 'charts')

pres
  // .addSlide('charts', 1)
  .addSlide('charts', 1, (slide: Slide) => {
    // slide.addElement('charts', 2, 'PieChart')
    slide.addElement('charts', 2, 'PieChart')
    // slide.addElement('charts', 2, 'PieChart')
    slide.addElement('images', 2, 'imageJPG')
    slide.addElement('charts', 1, 'StackedBars')
  })
  .addSlide('images', 2)

  .write(`myPresentation.pptx`).then(result => {
    console.info(result)
  }).catch(error => {
    console.error(error)
  })
