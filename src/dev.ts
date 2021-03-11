import Automizer from "./index"
import Slide from "./slide"
import { setSolidFill, setText } from "./helper/modify"

const automizer = new Automizer({
  templateDir: `${__dirname}/../__tests__/pptx-templates`,
  outputDir: `${__dirname}/../__tests__/pptx-output`
})

let pres = automizer.load(`RootTemplate.pptx`)
  .load(`SlideWithImage.pptx`, 'image')
  .load(`SlideWithShapes.pptx`, 'shapes')

pres
  .addSlide('image', 1, (slide: Slide) => {
    slide.addElement('shapes', 2, 'Cloud', [ setSolidFill, setText('my cloudy thoughts')] )
    slide.addElement('shapes', 2, 'Arrow', setText('my text'))
  })
  .write(`myPresentation.pptx`).then(result => {
    console.info(result)
  }).catch(error => {
    console.error(error)
  })
