import Automizer from "../src/automizer"

test("create presentation and append slides with images", async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`
  })

  let pres = automizer.loadRoot(`RootTemplateWithCharts.pptx`)
    .load(`SlideWithImages.pptx`, 'images')

  pres.addSlide('images', 1)
  pres.addSlide('images', 2)

  let result = await pres.write(`myPresentation.pptx`)

  expect(result.images).toBe(5)
})
