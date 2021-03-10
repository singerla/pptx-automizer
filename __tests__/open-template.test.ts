import Automizer from "../src/index"

test("open pptx template file", () => {
  const automizer = new Automizer()

  expect(automizer).toBeInstanceOf(Automizer)
});