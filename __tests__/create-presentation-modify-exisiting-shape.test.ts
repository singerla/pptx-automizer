import Automizer from "../src/automizer"
import { setPosition } from "../src/helper/modify"

test("create presentation, add slide with shapes from template and modify existing shape.", async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`
  })

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`SlideWithShapes.pptx`, 'shapes')

  const result = await pres
    .addSlide('shapes', 2, (slide) => {
      slide.modifyElement('Drum', [setPosition({x: 1000000, h:5000000, w:5000000})])
    })
    .write(`create-presentation-modify-existing-shape.test.pptx`)

  expect(result.slides).toBe(2)
})
