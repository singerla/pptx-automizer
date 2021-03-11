import Automizer from "../src/automizer"

test("create presentation and append charts to existing charts", async () => {
  const templateFolder = `${__dirname}/../__tests__/pptx-templates/`
  const outputFolder = `${__dirname}/../__tests__/pptx-output/`
  
  const automizer = new Automizer()
  let pres = automizer.importRootTemplate(`${templateFolder}RootTemplateWithCharts.pptx`)
    .importTemplate(`${templateFolder}SlideWithCharts.pptx`, 'charts')

  pres.addSlide('charts', 1)

  await pres.write(`${outputFolder}myPresentation.pptx`)

  expect(pres).toBeInstanceOf(Automizer)
})
