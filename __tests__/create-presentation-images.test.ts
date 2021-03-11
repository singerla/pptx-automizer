import Automizer from "../src/automizer"

test("create presentation and append slides with images", async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`
  })

  let pres = automizer.load(`RootTemplateWithCharts.pptx`)
    .load(`SlideWithImage.pptx`, 'image')

  pres.addSlide('image', 1)
  pres.addSlide('image', 2)
  pres.addSlide('image', 3)

  await pres.write(`myPresentation.pptx`)

  expect(pres).toBeInstanceOf(Automizer)
})
