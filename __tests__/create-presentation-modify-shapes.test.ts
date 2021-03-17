import Automizer from "../src/automizer"
import { setSolidFill, setText } from "../src/helper/modify"

test("create presentation, add some elements and modify content", async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`
  })

  let pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`SlideWithImages.pptx`, 'images')
    .load(`SlideWithShapes.pptx`, 'shapes')

  let result = await pres
    .addSlide('images', 1, (slide) => {
      slide.addElement('shapes', 2, 'Cloud', [ setSolidFill, setText('my cloudy thoughts')] )
      slide.addElement('shapes', 2, 'Arrow', setText('my text'))
      slide.addElement('shapes', 2, 'Drum')
    })
    .write(`create-presentation-modify-shapes.test.pptx`)

  expect(result.slides).toBe(2)
})
