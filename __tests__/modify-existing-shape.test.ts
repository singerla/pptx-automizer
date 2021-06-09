import Automizer, { modify } from '../src/index';

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
      slide.modifyElement('Drum', [modify.setPosition({x: 1000000, h:5000000, w:5000000})])
    })
    .write(`modify-existing-shape.test.pptx`)

  expect(result.slides).toBe(2)
})
