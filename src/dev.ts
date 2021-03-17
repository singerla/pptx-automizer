import Automizer from "./index"
import Slide from "./slide"
import { setSolidFill, setText } from "./helper/modify"

const automizer = new Automizer({
  templateDir: `${__dirname}/../__tests__/pptx-templates`,
  outputDir: `${__dirname}/../__tests__/pptx-output`
})

let pres = automizer.loadRoot(`RootTemplate.pptx`)
  .load(`SlideWithImages.pptx`, 'images')
  .load(`EmptySlide.pptx`, 'empty')

pres
  .addSlide('images', 2)
  .addSlide('empty', 1, (slide: Slide) => {
    // slide.addElement('charts', 2, 'PieChart')
    // slide.addElement('charts', 2, 'PieChart')
    // slide.addElement('charts', 2, 'PieChart')
    slide.addElement('images', 2, 'imageSVG')
    // slide.addElement('images', 2, 'imageSVG')
    // slide.addElement('images', 2, 'imageSVG')
    // slide.addElement('images', 2, 'imageSVG')
    // slide.addElement('charts', 1, 'StackedBars')
  })

  .write(`myPresentation.pptx`).then(result => {
    console.info(result)
  }).catch(error => {
    console.error(error)
  })
