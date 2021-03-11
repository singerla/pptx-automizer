import Automizer from "../src/index"

test("return automizer instance", () => {
  const automizer = new Automizer()

  expect(automizer).toBeInstanceOf(Automizer)
});