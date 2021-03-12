import Automizer from "../src/automizer"

test("create presentation and append slides with images", async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`
  })

  let pres = automizer.loadRoot(`RootTemplateWithCharts.pptx`)
    .load(`SlideWithImage.pptx`, 'image')

  pres.addSlide('image', 1)
  pres.addSlide('image', 2)
  pres.addSlide('image', 3)

  let result = await pres.write(`myPresentation.pptx`)

  expect(result.images).toBe(6)
})
