import Automizer from "../src/index"

test("throw error if template file not found", () => {
  const automizer = new Automizer()

  expect(() => {
    automizer.importRootTemplate(`non/existing/Template.pptx`)
  }).toThrow()
});