import Automizer from "./index"
import Slide from "./slide"
import { setSolidFill, setText } from "./helper/modify"

const automizer = new Automizer({
  templateDir: `${__dirname}/../__tests__/pptx-templates`,
  outputDir: `${__dirname}/../__tests__/pptx-output`
})

let pres = automizer.loadRoot(`RootTemplateWithCharts.pptx`)
  .load(`SlideWithImage.pptx`, 'image')
  .load(`SlideWithCharts.pptx`, 'charts')

pres
  // .addSlide('charts', 1)
  .addSlide('charts', 1, (slide: Slide) => {
    // slide.addElement('charts', 1, 'StackedBars')
    // slide.addElement('charts', 2, 'PieChart')
    slide.addElement('image', 3, 'imageJPG')
  })
  // .addSlide('charts', 2)

  // .addSlide('image', 1, (slide: Slide) => {
  //   slide.addElement('charts', 1, 'StackedBars')
  //   // slide.addElement('charts', 1, 'StackedBars')
  // })
  .write(`myPresentation.pptx`).then(result => {
    console.info(result)
  }).catch(error => {
    console.error(error)
  })
