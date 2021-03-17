import Automizer from "../src/automizer"
import { revertElements } from "../src/helper/modify"

test("create presentation and add some single images", async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`
  })

  let pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`EmptySlide.pptx`, 'empty')
    .load(`SlideWithImages.pptx`, 'images')

  let result = await pres
    .addSlide('empty', 1, (slide) => {
      slide.addElement('images', 2, 'imageJPG')
      slide.addElement('images', 2, 'imagePNG')

      slide.modify(revertElements)
    })
    .write(`create-presentation-insert-single-images.test.pptx`)

  expect(result.slides).toBe(2)
})
