import Automizer from "../src/automizer"

test("create presentation and append charts to existing charts", async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`
  })

  let pres = automizer.importRootTemplate(`RootTemplateWithCharts.pptx`)
    .importTemplate(`SlideWithCharts.pptx`, 'charts')

  pres.addSlide('charts', 1)

  await pres.write(`myPresentation.pptx`)

  expect(pres).toBeInstanceOf(Automizer)
})
