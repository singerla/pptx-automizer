import Automizer from "../src/automizer"

test("create presentation and append slides with notes", async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`
  })

  let pres = automizer.loadRoot(`RootTemplate.pptx`)
    .load(`SlideWithNotes.pptx`, 'notes')

  pres.addSlide('notes', 1)

  let result = await pres.write(`create-presentation-notes.test.pptx`)

  expect(result.slides).toBe(2)
})
